package main

import (
	"errors"
	"fmt"
	"math/rand"
	"os"
	"os/exec"
	"slices"
	"sort"
	"strconv"

	"github.com/kremec/edugroup/internal/dialogs"
	"github.com/kremec/edugroup/internal/excel"
	"github.com/kremec/edugroup/internal/types"
)

var DEBUG bool = true

func main() {

	numGroups := -1

	// Parse command line arguments
	args := os.Args[1:]
	if len(args) > 0 {
		argNumGroups, err := strconv.Atoi(args[0])
		if err != nil {
			dialogs.ShowErrorDialog(errors.New("Invalid number of groups"))
			return
		}
		numGroups = argNumGroups
	}

	// Open Excel file
	inputFile, err := dialogs.OpenExcelFile()
	if err != nil {
		dialogs.ShowErrorDialog(err)
		return
	}
	if DEBUG {
		fmt.Println("Input file:", inputFile)
	}

	var groups [][]string
	if numGroups == -1 {

		// Read Excel file
		data, err := excel.ReadExcelSubjectGroups(inputFile)
		if err != nil {
			dialogs.ShowErrorDialog(err)
			return
		}

		// Create student groups based on subjects and exclusions
		groups = createSubjectGroups(data)
	} else {

		// Read Excel file
		data, err := excel.ReadExcelNumGroups(inputFile)
		if err != nil {
			dialogs.ShowErrorDialog(err)
			return
		}

		// Create student groups based on number of groups
		groups = createNumGroups(data, numGroups)
	}

	// Export the groups to Excel file
	outputFile, err := dialogs.SaveExcelFile(inputFile)
	if err != nil {
		dialogs.ShowErrorDialog(err)
		return
	}
	if DEBUG {
		fmt.Println("Output file:", outputFile)
	}
	excel.ExportToExcel(groups, outputFile)

	// Open Excel file
	cmd := exec.Command("cmd", "/c", "start", outputFile)
	err = cmd.Start()
	if err != nil {
		dialogs.ShowErrorDialog(err)
		return
	}
}

// CreateGroups creates student groups based on the subjects and exclusions data.
func createSubjectGroups(data *types.GroupingData) [][]string {
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

// CreateGroups creates student groups based on the number of groups.
func createNumGroups(data *types.GroupingData, numGroups int) [][]string {
	groups := make([][]string, numGroups)

	// Shuffle exclusion groups for randomness
	for _, exclusionGroup := range data.Exclusions {
		rand.Shuffle(len(exclusionGroup), func(i, j int) { exclusionGroup[i], exclusionGroup[j] = exclusionGroup[j], exclusionGroup[i] })
	}

	// Shuffle students for randomness
	rand.Shuffle(len(data.Students), func(i, j int) { data.Students[i], data.Students[j] = data.Students[j], data.Students[i] })

	canAddToGroup := func(student string, group []string) bool {
		for _, studentInGroup := range group {
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

		// Add student to existing groups if possible
		sort.Slice(groups, func(i, j int) bool {
			return len(groups[i]) < len(groups[j])
		})
		foundGroupToAddTo := false
		for groupIndex, group := range groups {
			if canAddToGroup(student, group) {
				foundGroupToAddTo = true
				if DEBUG {
					fmt.Printf("Adding %s to group %s\n", student, group)
				}
				groups[groupIndex] = append(group, student)
				return
			}
		}
		if !foundGroupToAddTo {
			dialogs.ShowErrorDialog(errors.New("Exception constraints cannot be met for this number of groups"))
			os.Exit(0)
		}
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
			data.Students = slices.DeleteFunc(data.Students, func(s string) bool { return s == student })
		}
	}

	// Process remaining students
	for _, student := range data.Students {
		if DEBUG {
			fmt.Printf("Processing: %s\n", student)
		}
		processStudent(student)
		if DEBUG {
			fmt.Println("Current groups:", groups)
			fmt.Println()
		}
	}

	if DEBUG {
		fmt.Println("Final groups:", groups)
	}

	return groups
}
