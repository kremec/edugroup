package main

import (
	"fmt"
	"math/rand"
	"os"
	"slices"

	"github.com/kremec/edugroup/internal/dialogs"
	"github.com/kremec/edugroup/internal/excel"
	"github.com/kremec/edugroup/internal/types"
)

var DEBUG bool = false

func main() {

	// Open Excel file
	inputFile, err := dialogs.OpenExcelFile()
	if err != nil {
		dialogs.ShowErrorDialog(err)
		os.Exit(1)
	}

	fmt.Println("Input file:", inputFile)

	// Read Excel file
	data, err := excel.ReadExcelData(inputFile)
	if err != nil {
		dialogs.ShowErrorDialog(err)
		return
	}

	// Create student groups based on the data
	groups := createGroups(data)

	// Export the groups to Excel file
	outputFile, err := dialogs.SaveExcelFile(inputFile)
	if err != nil {
		dialogs.ShowErrorDialog(err)
		return
	}
	excel.ExportToExcel(groups, outputFile)
}

// CreateGroups creates student groups based on the subjects and exclusions data.
func createGroups(data *types.GroupingData) [][]string {
	groups := [][]string{}

	// Shuffle exclusion groups for randomness
	for _, exclusionGroup := range data.Exclusions {
		rand.Shuffle(len(exclusionGroup), func(i, j int) { exclusionGroup[i], exclusionGroup[j] = exclusionGroup[j], exclusionGroup[i] })
	}

	// Map students to corresponding subjects
	studentSubject := make(map[string]string)
	for subject, students := range data.SubjectStudents {
		// Shuffle students for randomness
		rand.Shuffle(len(students), func(i, j int) { students[i], students[j] = students[j], students[i] })

		for _, student := range students {
			if student != "" {
				studentSubject[student] = subject
			}
		}
	}

	canAddToGroup := func(student string, group []string) bool {
		for _, studentInGroup := range group {
			// Dissallow students from the same subject
			if studentSubject[studentInGroup] == studentSubject[student] {
				return false
			}

			// Dissallow students from same exclusion group
			for _, exclusionGroup := range data.Exclusions {
				if slices.Contains(exclusionGroup, studentInGroup) && slices.Contains(exclusionGroup, student) {
					return false
				}
			}
		}
		return true
	}
	processStudent := func(student string) {

		subject := studentSubject[student]
		if subject == "" {
			return
		}

		// Add student to existing groups if possible
		for groupIndex, group := range groups {
			if canAddToGroup(student, group) {
				if DEBUG {
					fmt.Printf("Adding %s to group %s\n", student, group)
				}
				groups[groupIndex] = append(group, student)
				return
			}
		}

		// Else create a new group
		if DEBUG {
			fmt.Printf("Creating new group for %s\n", student)
		}
		groups = append(groups, []string{student})
	}

	// Process exclusions first
	for _, exclusions := range data.Exclusions {
		for _, student := range exclusions {
			if DEBUG {
				fmt.Printf("Processing exclusion: %s\n", student)
			}
			processStudent(student)
			if DEBUG {
				fmt.Println("Current groups:", groups)
				fmt.Println()
			}

			// Remove the student from SubjectStudents
			for subject, students := range data.SubjectStudents {
				students = slices.DeleteFunc(students, func(s string) bool { return s == student })
				data.SubjectStudents[subject] = students
			}
		}
	}

	// Process remaining students
	for _, students := range data.SubjectStudents {
		for _, student := range students {
			if DEBUG {
				fmt.Printf("Processing: %s\n", student)
			}
			processStudent(student)
			if DEBUG {
				fmt.Println("Current groups:", groups)
				fmt.Println()
			}
		}
	}

	if DEBUG {
		fmt.Println("Final groups:", groups)
	}

	return groups
}
