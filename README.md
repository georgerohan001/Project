# **Exercise Sheet Corrector (README file)**

## Description
Welcome! You are in the README file of the exercise sheet corrector project! If you are not viewing this on Github, then it is recommended that you navigate to **{insert github link here}** for the latest updates and most accurate information. The GUI is also refined there so you are certain to have a better experience with the project.

## Table of Contents
- [Disclaimers](#Disclaimers)
- [Requirements](#Requirements)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)

## Disclaimer

* The function **quotes()** makes the running of the file much slower because the quotes exercise requires the fetching of data from 10 pages of a website. If you want faster processing time, consider commenting out the section in **extract_sheets()** where it is called (line 136). PLEASE DO NOT DELETE OR TAMPER WITH THE FUNCTION!!!

* An internet connection is required on the PC in which you run the exercise sheets program due to the quotes() function. Again, if internet connection is not available, head to line 136 and comment out the calling of the function.

* This project assumes the ownership of a directory of files in this format:

```lua
- A Mother folder
  |-- folder1
  |   |-- ExerciseSheet.zip
  |   |-- AnotherExerciseSheet.zip
  |-- folder2
  |   |-- ExerciseSheet.zip
  |   |-- AnotherExerciseSheet.zip
  |   |-- AThirdExerciseSheet.zip
  |-- folder3
  |   |-- ExerciseSheet.zip

```

The folders folder1, folder2 and folder3 should ideally be named in a way that it represents a student. This means that through reading the name, it should be identiable who the student is. Additionally, the number of exercise sheets in the folders does not matter. The code is designed to skip the correction of exercises that don't exist.