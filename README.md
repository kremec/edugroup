# EduGroup

This program is designed to take a list of students and grouping them:
- by students' subject groups
- by number of total groups

Both ways of grouping support defining exceptions for students who cannot work together.

## Usage

Start the program by running the .exe from Releases.

### Input

The program will display a file dialog to select the Excel file with student data to use.

Excel format - grouping by subject groups:
- First sheet: subject names as column headers in 1st row of sheet, below each is a column of student names from given subject group
- Second sheet: exception groups with group's student names in columns

Excel format - groupng by number of total groups:
- First sheet: one long column of student names starting in top left corner (cell A1)
- Second sheet: (same as above)

Names of sheets are not important, only ordering matters.

Examples of Excel input and output files can be found in `/examples`.

### Output

The program will then display the second file dialog to save the Excel file with generated student groups, each row representing one team.

The newly created Excel file will be then opened automatically.
