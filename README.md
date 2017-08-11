# PythonExcelScripts [![Build Status](https://travis-ci.org/tyl-/PythonExcelScripts.svg?branch=master)](https://travis-ci.org/tyl-/PythonExcelScripts) [![License](http://img.shields.io/:license-mit-blue.svg?style=flat-square)](http://badges.mit-license.org)

A collection of Python scripts that make it easier to generate and update excel spreadsheets.

---

## Main Goals

The main goals for this program are:
- Completed: Python script to generate excel spreadsheets based on templates.
- Completed: Python script to update monthly reports based on daily entries.
- Completed: Python script to update specific reports based on a specified range of daily reports.

## Start Date

- March 26, 2017

## Initial Completion Date

- March 26, 2017

## Goal Changes

- Completed: Python scripts to update files can be run both at the root directory and in the file directory.

---

## Updates

- **August 11, 2017**
> - Minor template fixes.
> - Added simple GUI to generate.py
> - Added option to select year (up to 2030) to generate.py
> - Added option to overwrite/skip existing files to generate.py

- **March 17, 2017**
> - Initial completion.

---

## Possible Improvements

- Refactor code.
- Add unit tests.
- Add GUI.
- Make it more general.

---

## Notes

- Can update excel files by clicking python scripts in the file folder or in the root folder.
- Specific reports use a date range in the form of MM-DD-YY to MM-DD-YY and will calculate data based on that date range.
- The scripts require the template structure and folder structure used by the generate script.

![Screenshot](/screenshots/ss1.jpg?raw=true "Screenshot of generate.py running")
![Screenshot](/screenshots/ss2.jpg?raw=true "Screenshot of monthlyupdate.py running")
![Screenshot](/screenshots/ss3.jpg?raw=true "Screenshot of salesupdate.py running")

---

## Special Thanks To

- <a href="http://fvcproductions.com" target="_blank">FVCproductions</a> for their README Template.
- <a href="https://travis-ci.org/" target="_blank">Travis CI</a>.

---

## License

[![License](http://img.shields.io/:license-mit-blue.svg?style=flat-square)](http://badges.mit-license.org)

- **[MIT license](http://opensource.org/licenses/mit-license.php)**