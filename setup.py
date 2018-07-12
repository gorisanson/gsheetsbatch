import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="gsheetsbatch",
    version="0.0.1",
    author="Lee Kyutae",
    author_email="gorisanson@gmail.com",
    description="a wrapper for Google Sheets API (PYTHON)",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/gorisanson/gsheetsbatch",
    packages=setuptools.find_packages(),
    classifiers=(
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ),
)
