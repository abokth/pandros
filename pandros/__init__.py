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
    def __init__(self, sheet, write_callback):
        self.sheet = sheet
        self.results = self.sheet
        self.write_callback = write_callback

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
        self.write_callback(self.results)
        print(self.results)

class Analysis:
    def __init__(self, sheet):
        self.sheet = sheet
        self.columns = ColumnsAnalysis(sheet)
        self.interpretation = self.columns.interpretation

    def print(self):
        self.columns.print()

class ColumnsAnalysis:
    def __init__(self, sheet):
        self.sheet = sheet
        self.columnanalyses = [ColumnAnalysis(sheet[column]) for column in sheet.columns]

        self.interpretation = InterpretationCandidates([PersonList]).find_one(self.columnanalyses)
        self.valid = self.interpretation is not None

    def print(self):
        self.interpretation.print()

class PersonList:
    def __init__(self, columnanalyses):
        columnanalysis_keys = [columnanalysis.interpretation.key for columnanalysis in columnanalyses if columnanalysis.interpretation is not None]
        if len(columnanalysis_keys) > len(set(columnanalysis_keys)):
            self.is_valid = False
            return

        columns = {columnanalysis.interpretation.key:columnanalysis for columnanalysis in columnanalyses if columnanalysis.interpretation is not None}

        keys = {
            'pnr': {'required': True },
            'family_name': {'required': True},
            'given_name': {'required': True},
            'email': {'required': False}
        }
        for (key,keyinfo) in keys.items():
            if keyinfo['required'] and key not in columns:
                self.is_valid = False
                return
            if key in columns:
                keys[key]['column'] = columns[key]
        self.columns = keys

        column_row_sizes = [len(self.columns[key]['column'].column) for key in self.columns.keys() if 'column' in self.columns[key]]
        if min(column_row_sizes) != max(column_row_sizes):
            self.is_valid = False
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

        persons = [Person(rows,i) for i in range(len(rows))]
        self.persons = persons

        self.items_type = 'persons'
        self.is_valid = True

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
        self.pnr = row['pnr'].values[0].replace("-", "")
        if len(self.pnr) == 12:
            self.pnr = self.pnr[2:]
        self.given_name = row['given_name'].values[0]
        self.family_name = row['family_name'].values[0]
        self.email = row['email'].values[0] if 'email' in row else None

    def print(self):
        print(f"pnr {self.pnr}, given name {self.given_name}, family name {self.family_name}, email {self.email}")

class ColumnAnalysis:
    def __init__(self, column):
        self.column = column

        self.interpretation = InterpretationCandidates([FamilyNameColumn, GivenNameColumn, PnrColumn, EmailColumn]).find_one(column)
        self.valid = self.interpretation is not None

    def print(self):
        self.interpretation.print()

class InterpretationCandidates:
    def __init__(self, classes):
        self.classes = classes

    def find_one(self, *args, **kwargs):
        interpretation_candidates = [interpretation(*args, **kwargs) for interpretation in self.classes]
        valid_interpretations = [interpretation_candidate for interpretation_candidate in interpretation_candidates if interpretation_candidate.is_valid]

        if len(valid_interpretations) != 1:
            return None

        return valid_interpretations[0]

class NameColumn:
    KEY = None
    NAME_RE = None

    def __init__(self, column):
        name = column.name.lower().strip()
        if not self.NAME_RE.match(name):
            self.is_valid = False
            return

        num_rows = len(column)
        num_alpha = len([row for row in column if row.strip().isalpha()])
        if 100 * num_alpha / num_rows < 80:
            self.is_valid = False
            return

        self.column = column
        self.names = [row.strip() for row in column]
        self.found_data = self.names
        self.key = self.KEY
        self.is_valid = True

    def print(self):
        print(f"{self.column.name} is a {self.KEY}")

class FamilyNameColumn(NameColumn):
    KEY = "family_name"
    NAME_RE = re.compile("((last|family).*name|efternamn)")

class GivenNameColumn(NameColumn):
    KEY = "given_name"
    NAME_RE = re.compile("((first|given).*name|förnamn)")

class PnrColumn:
    NAME_RE = re.compile("((person|t).*(number|nmr|nummer))|(birth(day|date)|födelse(dag|datum))")

    def __init__(self, column):
        name = column.name.lower().strip()
        if not self.NAME_RE.match(name):
            self.is_valid = False
            return

        pnrs = column.str.extract(r'(((19|20)\d\d|\d\d)[01]\d[0-3]\d((-|)[T\d]\d\d\d|))')[0]
        if pnrs.hasnans:
            self.is_valid = False
            return

        self.column = column
        self.pnrs = pnrs
        self.found_data = self.pnrs
        self.key = "pnr"
        self.is_valid = True

    def print(self):
        print(f"{self.column.name} is a pnr/birthdate column")

class EmailColumn:
    NAME_RE = re.compile("(e*-*mail|e*-*post)")

    def __init__(self, column):
        name = column.name.lower().strip()
        if not self.NAME_RE.match(name):
            self.is_valid = False
            return

        emails = column.str.extract('([\w\.]+@\w[\w\.]*\w\w)', flags=re.U)[0]
        if emails.hasnans:
            self.is_valid = False
            return

        self.column = column
        self.emails = emails
        self.found_data = self.emails
        self.key = "email"
        self.is_valid = True

    def print(self):
        print(f"{self.column.name} is an email column")

def read_file(path):
    if path.endswith(".xlsx"):
        return pd.read_excel(path)
    if path.endswith(".xls"):
        return pd.read_excel(path)
    if path.endswith(".odf"):
        return pd.read_excel(path)
    if path.endswith(".csv"):
        return pd.read_csv(path)
    raise Exception("Unknown input format")

def write_file(sheet, path):
    if path.endswith(".xlsx"):
        sheet.to_excel(path)
        return
    if path.endswith(".xls"):
        sheet.to_excel(path)
        return
    if path.endswith(".odf"):
        sheet.to_excel(path)
        return
    if path.endswith(".csv"):
        sheet.to_csv(path)
        return
    raise Exception("Unknown input format")
