package dialogs

import (
	"github.com/sqweek/dialog"
)

func OpenExcelFile() (string, error) {
	filename, err := dialog.File().Title("Open Excel file").Filter("Excel files", "xlsx").Filter("All files", "*").Load()

	return filename, err
}

func SaveExcelFile(filename string) (string, error) {
	filename, err := dialog.File().Title("Save Excel file").Filter("Excel files", "xlsx").Filter("All files", "*").Save()

	return filename, err
}

func ShowErrorDialog(err error) {
	dialog.Message("%s", err).Title("Error").Error()
}
