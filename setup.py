import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="pandros",
    version="0.0.1",
    author="Alexander Boström",
    author_email="abo@kth.se",
    description="Pandas based routines for interpreting account list shreadsheets",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://gita.sys.kth.se/abo/pandros",
    packages=['pandros'],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
    install_requires=["pandas", "openpyxl", "xlrd", "odf", "defusedxml"],
)

