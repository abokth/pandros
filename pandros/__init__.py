# -*- coding: utf-8 -*-

# Copyright 2020 Kungliga Tekniska högskolan

# Permission is hereby granted, free of charge, to any person
# obtaining a copy of this software and associated documentation files
# (the "Software"), to deal in the Software without restriction,
# including without limitation the rights to use, copy, modify, merge,
# publish, distribute, sublicense, and/or sell copies of the Software,
# and to permit persons to whom the Software is furnished to do so,
# subject to the following conditions:

# The above copyright notice and this permission notice shall be
# included in all copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
# NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
# BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
# ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
# CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

from __future__ import print_function, unicode_literals

import pandas as pd
import re

import defusedxml
from defusedxml.common import EntitiesForbidden

def defuse():
    defusedxml.defuse_stdlib()

class ValidationException(Exception):
    def readable(self, indent=0):
        if self.__cause__:
            if isinstance(self.__cause__, ValidationException):
                yield "   "*indent + str(self) + ":"
                yield from self.__cause__.readable(indent=indent+1)
            else:
                yield "   "*indent + str(self) + ": " + str(self.__cause__)
        else:
            yield "   "*indent + str(self)

    @property
    def long_message(self):
        return "\n".join(self.readable())

    def __eq__(self, other):
        return self.long_message == other.long_message

    def __lt__(self, other):
        return self.long_message < other.long_message

    def __hash__(self):
        return hash(self.long_message)

class MultiValidationException(ValidationException):
    def __init__(self, multi, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.multi = list(set(multi))
        self.multi.sort()

    def readable(self, indent=0):
        yield "   "*indent + f"{str(self)}:"
        for e in self.multi:
            yield from e.readable(indent=indent+1)

class AddResults:
    def __init__(self, person, result_collector, **results_to_add):
        self.person = person
        self.result_collector = result_collector
        self.results_to_add = results_to_add

    def describe(self):
        print(f"For person {self.person.index}:")
        for (key,value) in self.results_to_add.items():
            print(f" - {key}: {value}")

    def doit(self):
        self.result_collector.set_results(self.person.index, **self.results_to_add)

class ResultCollector:
    def __init__(self, sheet, fileupdater):
        self.sheet = sheet
        self.results = self.sheet
        self.fileupdater = fileupdater

    def get_value(self, row, key):
        if key not in self.results.keys():
            return None
        value = self.results[key][row]
        if pd.isnull(value) or value == "":
            return None
        return value

    def set_results(self, row, **new_results):
        for key in new_results.keys():
            if key not in self.results.keys():
                self.results = self.results.assign(**{key:""})
            col = self.results[key].copy()
            col[row] = new_results[key]
            self.results = self.results.assign(**{key:col})
        self.fileupdater.write_callback(self.results)
        print(self.results)

class FileAnalysis:
    def __init__(self, path):
        interpretations = [ValidOr(SheetReadAnalysis, path, header_row_shift=header) for header in range(4)]
        valid_interpretations = [interpretation for interpretation in interpretations if interpretation.res]

        if len(valid_interpretations) == 0:
            raise MultiValidationException([c.e for c in interpretations], "No valid interpretatons as a file with a header column")
        if len(valid_interpretations) > 1:
            raise ValidationException("Too many valid interpretatons")

        interpretation = valid_interpretations[0].res

        self.sheet = interpretation.sheet
        self.columns = interpretation.columns
        self.interpretation = self.columns.interpretation

        self.fileupdater = interpretation.fileupdater

    def get_writer(self, path):
        return ResultCollector(self.sheet, self.fileupdater)

    def print(self):
        self.interpretation.print()

class SheetReadAnalysis:
    def __init__(self, path, header_row_shift=0):
        sheet = read_file(path, header=header_row_shift)
        try:
            interpretation = Analysis(sheet)
        except ValidationException as e:
            raise ValidationException(f"Read with header shifted {header_row_shift} rows down failed") from e

        orig_sheet = read_file(path, header=None)
        self.fileupdater = SheetUpdater(path, orig_sheet, startrow=header_row_shift)

        self.sheet = interpretation.sheet
        self.columns = interpretation.columns
        self.interpretation = self.columns.interpretation

class SheetUpdater:
    def __init__(self, path, orig_sheet, **update_kwargs):
        def write_callback(new_sheet):
            writer = pd.ExcelWriter(path)
            # Write orig_sheet exactly as inputed
            orig_sheet.to_excel(writer, header=None, index=False)
            # Write new_sheet at the location where we read it before.
            new_sheet.to_excel(writer, index=False, **update_kwargs)
            writer.save()
        self.write_callback = write_callback

class Analysis:
    def __init__(self, sheet):
        self.sheet = sheet
        self.columns = ColumnsAnalysis(sheet)
        self.interpretation = self.columns.interpretation

class ColumnsAnalysis:
    def __init__(self, sheet):
        self.sheet = sheet
        columnanalyses_or = [ValidOr(ColumnAnalysis, sheet[column]) for column in sheet.columns]

        self.interpretation = InterpretationCandidates([PersonList]).find_one(columnanalyses_or)

class PersonList:
    def __init__(self, columnanalyses_or):
        columns = {columnanalysis_or.res.interpretation.key:columnanalysis_or.res for columnanalysis_or in columnanalyses_or if columnanalysis_or.res}

        keys = {
            'pnr': {'required': True },
            'family_name': {'required': True},
            'given_name': {'required': True},
            'email': {'required': False}
        }

        for (key,keyinfo) in keys.items():
            if keyinfo['required']:
                matching_columns = [columnanalysis_or.res for columnanalysis_or in columnanalyses_or if columnanalysis_or.res and columnanalysis_or.res.interpretation.key == key]
                if len(matching_columns) > 1:
                    raise ValidationException("multiple columns for {key}")

        for (key,keyinfo) in keys.items():
            if keyinfo['required'] and key not in columns:
                raise MultiValidationException([columnanalysis_or.e for columnanalysis_or in columnanalyses_or if columnanalysis_or.res is None], f"not enough data, missing {key}")
            if key in columns:
                keys[key]['column'] = columns[key]
        self.columns = keys

        column_row_sizes = [len(self.columns[key]['column'].column) for key in self.columns.keys() if 'column' in self.columns[key]]
        if min(column_row_sizes) != max(column_row_sizes):
            raise ValidationException("mismatched columns")
            return

        def renamed_column(key):
            if 'column' in self.columns[key]:
                column = self.columns[key]['column'].column
                return column
                #return column.rename(key)
            return None
        rows = pd.concat([column for column in [renamed_column(key) for key in self.columns.keys()] if column is not None], axis=1)

        for key in self.columns.keys():
            rows = rows.assign(**{key:self.columns[key]['column'].interpretation.found_data})

        self.new_sheet = rows

        valid_rows = None
        for (key,keyinfo) in keys.items():
            if keyinfo['required']:
                column_valid_rows = set(keys[key]['column'].interpretation.valid_rows)
                if valid_rows is None:
                    valid_rows = column_valid_rows
                else:
                    valid_rows = valid_rows.intersection(column_valid_rows)

        valid_rows = list(valid_rows)
        valid_rows.sort()
        persons = [Person(rows,i) for i in valid_rows]
        self.persons = persons

        self.items_type = 'persons'

    def print(self):
        print("Person information:")
        def column_for_print(key):
            if 'column' in self.columns[key]:
                column = self.columns[key]['column'].column
                return column.rename(f"{key} ({column.name})")
            return None
        key_names = sorted(self.columns.keys())
        c = pd.concat([column for column in [column_for_print(key) for key in key_names] if column is not None], axis=1)
        print(c)
        #print(self.new_sheet)
        for person in self.persons:
            person.print()

class Person:
    def __init__(self, rows, index):
        self.index = index
        row = rows[index:index+1]
        self.pnr = row['pnr'].values[0].replace("-", "").replace(" ", "")
        if len(self.pnr) == 12:
            self.pnr = self.pnr[2:]
        if len(self.pnr) == 10 and self.pnr[6:8] == "TF":
            self.pnr = self.pnr[0:6]
        self.given_name = row['given_name'].values[0]
        self.family_name = row['family_name'].values[0]
        self.email = row['email'].values[0] if 'email' in row else None

    def print(self):
        print(f"pnr {self.pnr}, given name {self.given_name}, family name {self.family_name}, email {self.email}")

class ColumnAnalysis:
    def __init__(self, column):
        self.column = column

        self.interpretation = InterpretationCandidates([FamilyNameColumn, GivenNameColumn, PnrColumn, EmailColumn]).find_one(column)

class ValidOr:
    def __init__(self, f, *args, **kwargs):
        self.res = None
        self.e = None
        try:
            self.res = f(*args, **kwargs)
        except ValidationException as e:
            self.e = e

class InterpretationCandidates:
    def __init__(self, classes):
        self.classes = classes

    def find_one(self, *args, **kwargs):
        interpretation_candidates = [ValidOr(interpretation, *args, **kwargs) for interpretation in self.classes]
        valid_interpretations = [interpretation_candidate.res for interpretation_candidate in interpretation_candidates if interpretation_candidate.res]

        if len(valid_interpretations) == 0:
            raise MultiValidationException([c.e for c in interpretation_candidates], "No valid interpretatons")
        if len(valid_interpretations) > 1:
            raise ValidationException("Too many valid interpretatons")

        return valid_interpretations[0]

class NameColumn:
    KEY = None
    NAME_RE = None

    def __init__(self, column):
        try:
            name = str(column.name.strip())
        except Exception as e:
            raise ValidationException(f"Could not parse column name '{column.name}'") from e
        if not self.NAME_RE.match(name):
            raise ValidationException(f"Unrecognized column name '{column.name}'")

        def istext(s):
            if pd.isna(s):
                return False
            s = str(s).strip()
            return len(s) > 1 and all([c.isalpha() or c.isspace() for c in s])

        num_rows = len(column)
        valid_rows = [i for i in column.index if istext(column[i])]
        if 100 * len(valid_rows) / num_rows < 80:
            raise ValidationException(f"Content of column '{column.name}' is not mostly alphabetical")

        self.column = column
        self.names = [str(row).strip() for row in column.convert_dtypes()]
        self.found_data = self.names
        self.key = self.KEY
        self.valid_rows = valid_rows

class FamilyNameColumn(NameColumn):
    KEY = "family_name"
    NAME_RE = re.compile("((last|family).*name|efternamn)", flags=re.I)

class GivenNameColumn(NameColumn):
    KEY = "given_name"
    NAME_RE = re.compile("((first|given).*name|förnamn)", flags=re.I)

class PnrColumn:
    NAME_RE = re.compile("((person|t).*(number|nmr|nr|nummer))|(birth(day|date)|födelse(dag|datum))", flags=re.I)

    def __init__(self, column):
        try:
            name = str(column.name.strip())
        except Exception as e:
            raise ValidationException(f"Could not parse column name '{column.name}'") from e
        if not self.NAME_RE.match(name): 
            raise ValidationException(f"Unrecognized column name '{column.name}'")

        pnrs = column.astype("string").str.extract(r'(((19|20)\d\d|\d\d)[01]\d[0-3]\d *((-|) *[T\d][\dF]\d\d|))')[0]
        valid_rows = [i for i in pnrs.index if not pd.isna(pnrs[i])]
        if 100 * len(valid_rows) / len(pnrs) < 80:
            raise ValidationException("Content does not match pnr data")

        self.column = column
        self.pnrs = pnrs
        self.found_data = self.pnrs
        self.key = "pnr"
        self.valid_rows = valid_rows

class EmailColumn:
    NAME_RE = re.compile("(e*-*mail|e*-*post)(adress|address)", flags=re.I)

    def __init__(self, column):
        try:
            name = str(column.name.strip())
        except Exception as e:
            raise ValidationException(f"Could not parse column name '{column.name}'") from e
        if not self.NAME_RE.match(name):
            raise ValidationException(f"Unrecognized column name '{column.name}'")

        emails = column.convert_dtypes().str.extract('([\w\.]+@\w[\w\.]*\w\w)', flags=re.U)[0]
        valid_rows = [i for i in emails.index if not pd.isna(emails[i])]
        if 100 * len(valid_rows) / len(emails) < 80:
            raise ValidationException("Content is not valid email addresses")

        self.column = column
        self.emails = emails
        self.found_data = self.emails
        self.key = "email"
        self.valid_rows = valid_rows

def read_file(path, *args, **kwargs):
    if path.endswith(".xlsx"):
        return pd.read_excel(path, *args, **kwargs)
    if path.endswith(".xls"):
        return pd.read_excel(path, *args, **kwargs)
    if path.endswith(".odf"):
        return pd.read_excel(path, *args, **kwargs)
    if path.endswith(".csv"):
        return pd.read_csv(path, *args, **kwargs)
    raise Exception("Unknown input format")

def write_file(sheet, path, **kwargs):
    if path.endswith(".xlsx"):
        sheet.to_excel(path, **kwargs)
        return
    if path.endswith(".xls"):
        sheet.to_excel(path, **kwargs)
        return
    if path.endswith(".odf"):
        sheet.to_excel(path, **kwargs)
        return
    if path.endswith(".csv"):
        sheet.to_csv(path, **kwargs)
        return
    raise Exception("Unknown input format")

