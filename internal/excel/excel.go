package excel

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/kremec/edugroup/internal/types"

	"github.com/xuri/excelize/v2"
)

const (
	errNotifyDeveloper     = "Please notify the developer of this error!"
	errOpeningExcelFile    = "Error opening Excel file:"
	errNoSheetsInExcelFile = "No sheets found in Excel file, make sure to create at least one sheet and fill it with student data!"
	errNoDataInExcelFile   = "No data found in the first sheet, make sure to add subject headers and student data!"
	errParsingExcelFile    = "Error reading data from Excel file:"
	errInvalidExcelInput   = "Invalid Excel input:"
	errSavingExcelFile     = "Error saving Excel file:"
	groupsSheetName        = "Groups"
)

type cellValueRef struct {
	value string
	cell  string
}

type validationErrors struct {
	issues []string
}

func (v *validationErrors) add(format string, args ...any) {
	v.issues = append(v.issues, fmt.Sprintf(format, args...))
}

func (v *validationErrors) err() error {
	if len(v.issues) == 0 {
		return nil
	}

	return fmt.Errorf("%s\n- %s", errInvalidExcelInput, strings.Join(v.issues, "\n- "))
}

// ReadExcelSubjectGroups loads the data from the specified Excel file.
func ReadExcelSubjectGroups(filename string) (*types.GroupingData, error) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, fmt.Errorf("%s %s\n%s", errOpeningExcelFile, err, errNotifyDeveloper)
	}
	defer f.Close()

	subjectStudents, err := getSubjectsStudents(f)
	if err != nil {
		return nil, err
	}

	exclusions, err := getExclusions(f, flattenSubjectStudents(subjectStudents))
	if err != nil {
		return nil, err
	}

	data := &types.GroupingData{
		SubjectStudents: subjectStudents,
		Exclusions:      exclusions,
	}

	return data, nil
}

// Read subjects and their students from the 1st sheet of Excel file
func getSubjectsStudents(f *excelize.File) (map[string][]string, error) {
	subjectStudents := make(map[string][]string)

	// If there is no 1st sheet, throw an error
	if f.GetSheetName(0) == "" {
		return nil, fmt.Errorf("%s", errNoSheetsInExcelFile)
	}

	rows, err := f.GetRows(f.GetSheetName(0))
	if err != nil {
		return nil, fmt.Errorf("%s %s\n%s", errParsingExcelFile, err, errNotifyDeveloper)
	}
	if len(rows) == 0 {
		return nil, fmt.Errorf("%s", errNoDataInExcelFile)
	}

	issues := &validationErrors{}
	if countNonEmptyCells(rows[0]) == 0 {
		issues.add("row 1 of the first sheet is empty; subject headers must start in row 1")
		return nil, issues.err()
	}

	if countNonEmptyCells(rows[0]) == 1 && len(rows) > 1 && countNonEmptyCells(rows[1]) > 1 {
		issues.add("row 1 of the first sheet looks like a title row; subject headers must start in row 1")
	}

	subjects := rows[0]
	maxCols := maxColumnCount(rows)
	seenSubjects := make(map[string]cellValueRef)
	seenStudents := make(map[string]cellValueRef)
	seenStudentsNormalized := make(map[string]cellValueRef)

	// Read students for each subject from the columns
	for colIndex := 0; colIndex < maxCols; colIndex++ {
		cell := spreadsheetCell(colIndex, 0)
		rawSubject := ""
		if colIndex < len(subjects) {
			rawSubject = subjects[colIndex]
		}
		subject := trimmedValue(rawSubject)

		if rawSubject != "" && rawSubject != subject {
			issues.add("subject header at %s contains leading or trailing spaces", cell)
		}

		hasStudentsBelow := false
		for rowIndex := 1; rowIndex < len(rows); rowIndex++ {
			if colIndex < len(rows[rowIndex]) && trimmedValue(rows[rowIndex][colIndex]) != "" {
				hasStudentsBelow = true
				break
			}
		}

		if subject == "" {
			if hasStudentsBelow {
				issues.add("%s is missing a subject name while cells below it contain student names", cell)
			}
			continue
		}

		subjectKey := strings.ToLower(subject)
		if first, exists := seenSubjects[subjectKey]; exists {
			if first.value == subject {
				issues.add("subject %q is duplicated at %s and %s", subject, first.cell, cell)
			} else {
				issues.add("subject names %q (%s) and %q (%s) differ only by letter case", first.value, first.cell, subject, cell)
			}
			continue
		}
		seenSubjects[subjectKey] = cellValueRef{value: subject, cell: cell}

		for rowIndex := 1; rowIndex < len(rows); rowIndex++ {
			if colIndex >= len(rows[rowIndex]) {
				continue
			}

			studentCell := spreadsheetCell(colIndex, rowIndex)
			rawStudentName := rows[rowIndex][colIndex]
			studentName := trimmedValue(rawStudentName)
			if rawStudentName != "" && rawStudentName != studentName {
				issues.add("student name at %s contains leading or trailing spaces", studentCell)
			}

			if studentName != "" {
				if first, exists := seenStudents[studentName]; exists {
					issues.add("student %q is duplicated at %s and %s", studentName, first.cell, studentCell)
					continue
				}

				studentKey := strings.ToLower(studentName)
				if first, exists := seenStudentsNormalized[studentKey]; exists {
					issues.add("student names %q (%s) and %q (%s) differ only by letter case", first.value, first.cell, studentName, studentCell)
					continue
				}

				seenStudents[studentName] = cellValueRef{value: studentName, cell: studentCell}
				seenStudentsNormalized[studentKey] = cellValueRef{value: studentName, cell: studentCell}
				subjectStudents[subject] = append(subjectStudents[subject], studentName)
			}
		}
	}

	if len(subjectStudents) == 0 {
		issues.add("no student names were found below the subject headers on the first sheet")
	}

	if err := issues.err(); err != nil {
		return nil, err
	}

	return subjectStudents, nil
}

// Read exclusions from the 2nd sheet of Excel file
func getExclusions(f *excelize.File, knownStudents []string) ([][]string, error) {

	// If there is no 2nd sheet assume no exlusions
	if f.GetSheetName(1) == "" {
		return make([][]string, 0), nil
	}

	columns, err := f.GetCols(f.GetSheetName(1))
	if err != nil {
		return nil, fmt.Errorf("%s %s\n%s", errParsingExcelFile, err, errNotifyDeveloper)
	}

	issues := &validationErrors{}
	knownStudentsExact := make(map[string]struct{}, len(knownStudents))
	knownStudentsNormalized := make(map[string]string, len(knownStudents))
	for _, student := range knownStudents {
		knownStudentsExact[student] = struct{}{}
		knownStudentsNormalized[strings.ToLower(student)] = student
	}

	seenAcrossGroups := make(map[string]cellValueRef)
	exclusions := make([][]string, 0, len(columns))
	for colIndex, column := range columns {
		exclusionGroup := make([]string, 0, len(column))
		seenInGroup := make(map[string]cellValueRef)

		for rowIndex, rawName := range column {
			cell := spreadsheetCell(colIndex, rowIndex)
			name := trimmedValue(rawName)

			if rawName != "" && rawName != name {
				issues.add("exclusion name at %s contains leading or trailing spaces", cell)
			}

			if name == "" {
				continue
			}

			canonicalName, exists := knownStudentsNormalized[strings.ToLower(name)]
			if !exists {
				issues.add("exclusion name %q at %s does not match any student from the first sheet", name, cell)
				continue
			}

			if canonicalName != name {
				issues.add("exclusion name %q at %s must match the first-sheet student name exactly: %q", name, cell, canonicalName)
				continue
			}

			if first, exists := seenInGroup[name]; exists {
				issues.add("student %q is listed twice in the same exclusion group at %s and %s", name, first.cell, cell)
				continue
			}

			if first, exists := seenAcrossGroups[name]; exists {
				issues.add("student %q appears in more than one exclusion group at %s and %s", name, first.cell, cell)
				continue
			}

			seenInGroup[name] = cellValueRef{value: name, cell: cell}
			seenAcrossGroups[name] = cellValueRef{value: name, cell: cell}
			exclusionGroup = append(exclusionGroup, name)
		}

		if len(exclusionGroup) > 0 {
			exclusions = append(exclusions, exclusionGroup)
		}
	}

	if err := issues.err(); err != nil {
		return nil, err
	}

	return exclusions, nil
}

func ReadExcelNumGroups(filename string) (*types.GroupingData, error) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, fmt.Errorf("%s %s\n%s", errOpeningExcelFile, err, errNotifyDeveloper)
	}
	defer f.Close()

	students, err := getStudents(f)
	if err != nil {
		return nil, err
	}

	exclusions, err := getExclusions(f, students)
	if err != nil {
		return nil, err
	}

	data := &types.GroupingData{
		Students:   students,
		Exclusions: exclusions,
	}

	return data, nil
}

func getStudents(f *excelize.File) ([]string, error) {

	// If there is no 1st sheet, throw an error
	if f.GetSheetName(0) == "" {
		return nil, fmt.Errorf("%s", errNoSheetsInExcelFile)
	}

	rows, err := f.GetRows(f.GetSheetName(0))
	if err != nil {
		return nil, fmt.Errorf("%s %s\n%s", errParsingExcelFile, err, errNotifyDeveloper)
	}
	if len(rows) == 0 {
		return nil, fmt.Errorf("%s", errNoDataInExcelFile)
	}

	issues := &validationErrors{}
	students := make([]string, 0, len(rows))
	seenStudents := make(map[string]cellValueRef)
	seenStudentsNormalized := make(map[string]cellValueRef)

	for rowIndex, row := range rows {
		if rowIndex == 0 && (len(row) == 0 || trimmedValue(row[0]) == "") {
			issues.add("cell A1 must contain the first student name in number-of-groups mode")
		}

		for colIndex := 1; colIndex < len(row); colIndex++ {
			if trimmedValue(row[colIndex]) != "" {
				issues.add("%s contains %q, but number-of-groups mode only reads student names from column A", spreadsheetCell(colIndex, rowIndex), row[colIndex])
			}
		}

		if len(row) == 0 {
			continue
		}

		cell := spreadsheetCell(0, rowIndex)
		rawStudent := row[0]
		student := trimmedValue(rawStudent)
		if rawStudent != "" && rawStudent != student {
			issues.add("student name at %s contains leading or trailing spaces", cell)
		}

		if student == "" {
			continue
		}

		if first, exists := seenStudents[student]; exists {
			issues.add("student %q is duplicated at %s and %s", student, first.cell, cell)
			continue
		}

		studentKey := strings.ToLower(student)
		if first, exists := seenStudentsNormalized[studentKey]; exists {
			issues.add("student names %q (%s) and %q (%s) differ only by letter case", first.value, first.cell, student, cell)
			continue
		}

		seenStudents[student] = cellValueRef{value: student, cell: cell}
		seenStudentsNormalized[studentKey] = cellValueRef{value: student, cell: cell}
		students = append(students, student)
	}

	if len(students) == 0 {
		issues.add("no student names were found in column A of the first sheet")
	}

	if err := issues.err(); err != nil {
		return nil, err
	}

	return students, nil
}

// ExportToExcel exports the groups to an Excel file.
func ExportToExcel(groups [][]string, filename string) error {
	f := excelize.NewFile()
	defer f.Close()

	err := f.SetSheetName(f.GetSheetName(0), groupsSheetName)
	if err != nil {
		return fmt.Errorf("%s %s\n%s", errOpeningExcelFile, err, errNotifyDeveloper)
	}

	for i, group := range groups {
		cell := fmt.Sprintf("%c%d", 'A', i+1)
		f.SetCellValue(groupsSheetName, cell, "Group "+strconv.Itoa(i+1))
		for j, student := range group {
			cell := fmt.Sprintf("%c%d", 'A'+j+1, i+1)
			f.SetCellValue(groupsSheetName, cell, student)
		}
	}
	f.SetActiveSheet(0)

	err = f.SaveAs(filename)
	if err != nil {
		return fmt.Errorf("%s %s\n%s", errSavingExcelFile, err, errNotifyDeveloper)
	}
	return nil
}

func flattenSubjectStudents(subjectStudents map[string][]string) []string {
	students := make([]string, 0)
	for _, subjectGroup := range subjectStudents {
		students = append(students, subjectGroup...)
	}

	return students
}

func countNonEmptyCells(row []string) int {
	count := 0
	for _, value := range row {
		if trimmedValue(value) != "" {
			count++
		}
	}

	return count
}

func maxColumnCount(rows [][]string) int {
	maxCols := 0
	for _, row := range rows {
		if len(row) > maxCols {
			maxCols = len(row)
		}
	}

	return maxCols
}

func spreadsheetCell(colIndex, rowIndex int) string {
	cell, err := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
	if err != nil {
		return fmt.Sprintf("row %d column %d", rowIndex+1, colIndex+1)
	}

	return cell
}

func trimmedValue(value string) string {
	return strings.TrimSpace(value)
}
