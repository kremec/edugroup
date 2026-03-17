# EduGroup

This program is designed to take a list of students and grouping them:

- by students' **subject groups**
- by **number of total groups**

Both ways of grouping support defining:

- **exceptions** for students who cannot work together
- **required groups** for students who should stay together

## Usage

Start the program by running the .exe from Releases.

The program asks the user to select the grouping mode:

- 0\) Group students by subject groups
- n\) Group students into 'n' groups

IF the user inputs '0' and then ENTER, the program will group the students by subject groups.

IF the user inputs any other number and then ENTER, the program will group the students into given number of groups.

### Input

After the mode selection, the program will display a file dialog to select the Excel file with student data to use.

Excel format - grouping by subject groups:

- First sheet: subject names as column headers in 1st row of sheet, below each is a column of student names from given subject group
- Second sheet: exception groups with group's student names in columns
- Third sheet: required groups with group's student names in columns

Excel format - grouping by number of total groups:

- First sheet: one long column of student names starting in top left corner (cell A1)
- Second sheet: (same as above)
- Third sheet: (same as above)

Names of sheets are not important, only ordering matters.
Second and third sheets are optional. If omitted, the program assumes there are no constraints of that type.

Examples of Excel input and output files can be found in `/examples`.

### Output

The program will then display the second file dialog to save the Excel file with generated student groups, each row representing one team.

The newly created Excel file will be then opened automatically.
