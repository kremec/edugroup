package excel

import (
	"fmt"
	"strconv"

	"github.com/kremec/edugroup/internal/types"

	"github.com/xuri/excelize/v2"
)

const (
	errNotifyDeveloper     = "Please notify the developer of this error!"
	errOpeningExcelFile    = "Error opening Excel file:"
	errNoSheetsInExcelFile = "No sheets found in Excel file, make sure to create at least one sheet and fill it with student data!"
	errParsingExcelFile    = "Error reading data from Excel file:"
	errSavingExcelFile     = "Error saving Excel file:"
	groupsSheetName        = "Groups"
)

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

	exclusions, err := getExclusions(f)
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

	subjects := rows[0]

	// Read students for each subject from the columns
	for colIndex, subject := range subjects {
		if subject == "" {
			continue
		}

		for rowIndex := 1; rowIndex < len(rows); rowIndex++ {
			studentName := rows[rowIndex][colIndex]
			if studentName != "" {
				subjectStudents[subject] = append(subjectStudents[subject], studentName)
			}
		}
	}

	return subjectStudents, nil
}

// Read exclusions from the 2nd sheet of Excel file
func getExclusions(f *excelize.File) ([][]string, error) {

	// If there is no 2nd sheet assume no exlusions
	if f.GetSheetName(1) == "" {
		return make([][]string, 0), nil
	}

	columns, err := f.GetCols(f.GetSheetName(1))
	if err != nil {
		return nil, fmt.Errorf("%s %s\n%s", errParsingExcelFile, err, errNotifyDeveloper)
	}

	return columns, nil
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

	exclusions, err := getExclusions(f)
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

	columns, err := f.GetCols(f.GetSheetName(0))
	if err != nil {
		return nil, fmt.Errorf("%s %s\n%s", errParsingExcelFile, err, errNotifyDeveloper)
	}

	return columns[0], nil
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
