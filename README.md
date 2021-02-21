# Excel2MoodleXML

This python script converts Quiz in an Excel File (in .xlsx format) to Moodle XML format. 

License: GPL V3 or at your convenience later versions of GPL

### Features
- Every sheet is converted in to a question category
- Support for images in both questions and answers
- Automatic insertion of HTML line break tag `<br>` for non equation texts
- Automatic check for missing keys

### Limitations of this version:
- It is a preliminary version and is limited to Multiple Choice Questions alone. 
- Feeback for questions/answers are not yet supported

### Command:
` python3 Excel2MoodleXML.py <<input.xlsx>> <<output.xml>> `



Example:
` python3 Excel2MoodleXML.py Quiz.xlsx Quiz.xml`

Even if any argument is not given, by default, Quiz.xml is considered as input file and Quiz.xml is considered as name of output file

### Dependencies:
- Python 3.6 and above
- Packages
  - Openpyxl
  - Openpyxl-image-loader

### Background:
I am a teacher, who use moodle as learning management system in classroom. Creating Quizzes in Moodle Online version is exhaustive. Even though quiz can be created in AIKEN format, I felt it difficult to manage large set of questions and AIKEN does not support images. After exhaustive search, I found about a free software called QuestionMachine. I used it and have recommended it to my friends. It runs on .NET and was developed around 2012. Since then, it was not updated. I liked that software. But it supported images only for questions. Also, copy pasting questions from one quiz to another is very diffcult. So I thought about this converter. So that, any spreadsheet application which is capable of creating .xlsx file can be used for creating quiz in Moodle.
