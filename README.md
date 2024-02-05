# **Exercise Sheet Corrector (README file)**

## Description
Welcome! You are in the README file of the exercise sheet corrector project! If you are not viewing this on Github, then it is recommended that you navigate to **{insert github link here}** for the latest updates and most accurate information. The GUI is also refined there so you are certain to have a better experience with the project.

## Table of Contents
- [Disclaimers](#Disclaimers)
- [Included Files](#included-files)
- [Requirements](#Requirements)
- [Usage](#Usage)
- [Contributing](#Contributing)
- [Authors](#Authors)


## Disclaimers

* The function **quotes()** makes the running of the file much slower because the quotes exercise requires the fetching of data from 10 pages of a website. If you want faster processing time, consider commenting out the section in **extract_sheets()** where it is called (line 136). PLEASE DO NOT DELETE OR TAMPER WITH THE FUNCTION!!!

* An internet connection is required on the PC in which you run the exercise sheets program due to the quotes() function. Again, if an internet connection is not available, head to line 136 and comment out the calling of the function.

* This project assumes the ownership of a directory of files in this format:

```lua
├── A Mother folder
|   ├── folder1
|   |   ├── ExerciseSheet.zip
|   |   ├── AnotherExerciseSheet.zip
|   ├── folder2
|   |   ├── ExerciseSheet.zip
|   |   ├── AnotherExerciseSheet.zip
|   |   ├── AThirdExerciseSheet.zip
|   ├── folder3
|   |   ├── ExerciseSheet.zip

```

The folders `folder1`, `folder2`, and `folder3` should ideally be named in a way that represents a student. This means that through reading the name, it should be identifiable who the student is. Additionally, the number of exercise sheets in the folders does not matter. The code is designed to skip the correction of exercises that don't exist.

* The code has been tested on several examples, but if the case does occur that the code throws an error on your PC, please attempt it on the sample directory **{name of directory}** provided in the submission.

## Included files

* [Project File](/Final%20Project%201.py): This is the main project file. This is what you will run 
* [README File](README.md): This is this file. The one that gives you context or important information
* [Sample Directory](/Test%20Parent%20Folder%20-%20Backup.zip): This is the folder that you can use to test the code if you don't want to test on your own directories.
* [Requirements file](requirements.txt): This is the file that has only the Python version inside. This is just to ensure that the library versions would be the same.

## Requirements

Although there is a `requirements.txt` available for this project, only the Python version *(v3.12.1)* is present inside. This is because all libraries used are a part of Python's standard libraries. This means that as long as you have the same Python version, you should have the same version for all libraries. One exception that may be worth checking is the tkinter library, but if you type in the installation command, then the latest version should be the same. To install pip, one can run:

```bash
pip install tk
```
To check to see if your Python version is correct, run the command:

```bash
python --version
```
If the result was something other than `3.12.1`, then you might consider running the code anyway, and if it does not work, then you might consider down/updating your Python to the `3.12.1` version.

## Usage

IMPORTANT! Please make a backup of the folders that you run the code on. The code is designed to make a backup, but having an extra safeguard shouldn't hurt.

To use this Python file, please run it in your preferred way, after making sure you have directories available like in the example present in the [Disclaimer](#disclaimers) section. To run it with the console, just run:

```bash
python -u {filepath}
```
Be sure to replace the {filepath} placeholder with the real file path. If you are already in the correct directory in your terminal (using the `cd` command), then the {filename} would just suffice, with the .py suffix.

Upon running, be sure to click on the main mother folder directory, and click on "Select Parent Folder", once that has been done, please wait. You will know the code is done when you see that the terminal once again becomes available for commands. Early exits can be done with **Ctrl + C**, but waiting until it is ready is recommended.

After the running is finished, you may navigate into the parent directory and open the child folders, you should see no zip files there anymore, and instead:

* A **Points log** which tells you how many points the grader was able to successfully validate. This would mean that the remaining sheets are available in the "Manual Correction Needed" folder for further checking.

* A **"Successful Sheets"** folder, which contains the .py pages for all sheets that passed the test conditions of the grader.

* An **"Manual Correction Needed"** folder, containing the .py files that either did not pass or were automatically placed there since an automatic correction was not possible / would not make sense.

* An **"Already Extracted Sheets"** folder which contains all the zip files for future use, if necessary.

* An **"Unrecognized Sheets"** folder, containing files with names that the project code could not recognize. This can include typos, cache files, accidentally added files, and further.

## Contributing

Although this is a one-time submission, many man-hours have been spent on this project, which means that we are dedicated to having the code run optimally and correctly. For this reason, it would be great if we received any feedback on the code or suggestions on how to run it more effectively. Our communication channels are mainly:

### 1. **Github**:
Click on the `Issues` tab of this GitHub page to submit a query to us.
### 2. **Email**:
Assuming this is being read by a course professor, it is assumed that you have the emails of the group members, feel free to reach out in case of any questions/problems
### 3. **ILIAS**:
Once again, assuming that this is being seen by the course professors, then ILIAS is also a method of communication.

## Authors

Let's have some appreciation for those involved in the project!
The collaborators are:

*Husam Al Ahmadieh*

*Fe Bossert*

*George Rohan Pottamkulam*
