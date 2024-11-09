package excel

import (
	"fmt"

	"github.com/kremec/edugroup/internal/types"

	"github.com/xuri/excelize/v2"
)

const (
	errNotifyDeveloper  = "Please notify the developer of this error!"
	errOpeningExcelFile = "Error opening Excel file:"
	errParsingExcelFile = "Error reading data from Excel file:"
	errSavingExcelFile  = "Error saving Excel file:"
	groupsSheetName     = "Groups"
)

// ReadExcelData loads the data from the specified Excel file.
func ReadExcelData(filename string) (*types.GroupingData, error) {
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

	columns, err := f.GetCols(f.GetSheetName(1))
	if err != nil {
		return nil, fmt.Errorf("%s %s\n%s", errParsingExcelFile, err, errNotifyDeveloper)
	}

	return columns, nil
}

// ExportToExcel exports the groups to an Excel file.
func ExportToExcel(groups [][]string, filename string) error {
	f := excelize.NewFile()
	sheet, err := f.NewSheet(groupsSheetName)
	if err != nil {
		return fmt.Errorf("%s %s\n%s", errOpeningExcelFile, err, errNotifyDeveloper)
	}

	for i, group := range groups {
		for j, student := range group {
			cell := fmt.Sprintf("%c%d", 'A'+j, i+1)
			f.SetCellValue(groupsSheetName, cell, student)
		}
	}
	f.SetActiveSheet(sheet)

	err = f.SaveAs(filename)
	if err != nil {
		return fmt.Errorf("%s %s\n%s", errSavingExcelFile, err, errNotifyDeveloper)
	}
	return nil
}
