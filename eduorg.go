package main

import (
	"bufio"
	"errors"
	"fmt"
	"math/rand"
	"os"
	"os/exec"
	"slices"
	"sort"
	"strconv"
	"strings"

	"github.com/kremec/edugroup/internal/dialogs"
	"github.com/kremec/edugroup/internal/excel"
	"github.com/kremec/edugroup/internal/types"
)

const (
	redText   = "\033[31m"
	resetText = "\033[0m"
)

var DEBUG bool = false

func main() {
	// Parse command line arguments
	args := os.Args[1:]
	if len(args) > 0 {
		if args[0] == "-debug" {
			DEBUG = true
			args = args[1:]
		}
	}

	fmt.Println("Welcome to EduGroup!")
	for {
		// Give user instructions
		fmt.Println("Available grouping modes:")
		fmt.Printf("0) %s\n", "Group students by subject groups")
		fmt.Printf("n) %s\n", "Group students into 'n' groups")
		fmt.Println("<ENTER>) Exit")
		fmt.Print("Enter your choice: ")

		// Read user input
		reader := bufio.NewReader(os.Stdin)
		input, _ := reader.ReadString('\n')
		input = strings.TrimSpace(input) // Remove any leading/trailing whitespace

		// Exit if user presses ENTER
		if input == "" {
			fmt.Println("Exiting the program.")
			return
		}

		groupMode, err := strconv.Atoi(input)
		if err != nil || groupMode < 0 {
			fmt.Printf("%sInvalid input. Please enter 0, a positive integer, or press ENTER to exit.%s\n", redText, resetText)
			restartProgramDelimiter()
			continue
		}

		// Open Excel file
		inputFile, err := dialogs.OpenExcelFile()
		if err != nil {
			dialogs.ShowErrorDialog(err)
			restartProgramDelimiter()
			continue
		}
		if DEBUG {
			fmt.Println("Input file:", inputFile)
		}

		var groups [][]string
		if groupMode == 0 {
			// Read Excel file
			data, err := excel.ReadExcelSubjectGroups(inputFile)
			if err != nil {
				dialogs.ShowErrorDialog(err)
				restartProgramDelimiter()
				continue
			}

			// Create student groups based on subjects and exclusions
			groups, err = createSubjectGroups(data)
			if err != nil {
				dialogs.ShowErrorDialog(err)
				restartProgramDelimiter()
				continue
			}
		} else {
			numGroups := groupMode
			// Read Excel file
			data, err := excel.ReadExcelNumGroups(inputFile)
			if err != nil {
				dialogs.ShowErrorDialog(err)
				restartProgramDelimiter()
				continue
			}

			// Create student groups based on number of groups
			groups, err = createNumGroups(data, numGroups)
			if err != nil {
				dialogs.ShowErrorDialog(err)
				restartProgramDelimiter()
				continue
			}
		}

		fmt.Printf("\nGrouping successful - %d groups created.\n", len(groups))

		// Export the groups to Excel file
		outputFile, err := dialogs.SaveExcelFile(inputFile)
		if err != nil {
			dialogs.ShowErrorDialog(err)
			restartProgramDelimiter()
			continue
		}
		if DEBUG {
			fmt.Println("Output file:", outputFile)
		}
		excel.ExportToExcel(groups, outputFile)

		fmt.Println("Groups exported to", outputFile)

		// Open Excel file
		cmd := exec.Command("cmd", "/c", "start", outputFile)
		err = cmd.Start()
		if err != nil {
			dialogs.ShowErrorDialog(err)
			restartProgramDelimiter()
			continue
		}
		restartProgramDelimiter()
	}
}

func restartProgramDelimiter() {
	fmt.Println()
	fmt.Println(" - - - - - ")
	fmt.Println()
}

// CreateGroups creates student groups based on the subjects and exclusions data.
func createSubjectGroups(data *types.GroupingData) ([][]string, error) {
	groups := [][]string{}
	exclusionLookup := buildExclusionLookup(data.Exclusions)

	// Map students to corresponding subjects
	studentSubject := make(map[string]string)
	for subject, students := range data.SubjectStudents {
		for _, student := range students {
			if student != "" {
				studentSubject[student] = subject
			}
		}
	}

	if err := validateSubjectInclusions(data.Inclusions, studentSubject, exclusionLookup); err != nil {
		return nil, err
	}

	allStudents := flattenSubjectStudentsBySubject(data.SubjectStudents)
	units := buildAssignmentUnits(allStudents, data.Inclusions, exclusionLookup)

	canAddUnitToGroup := func(unit []string, group []string) bool {
		for _, student := range unit {
			for _, studentInGroup := range group {
				// Dissallow students from the same subject
				if studentSubject[studentInGroup] == studentSubject[student] {
					return false
				}

				if studentsConflict(student, studentInGroup, exclusionLookup) {
					return false
				}
			}
		}
		return true
	}

	processUnit := func(unit []string) {
		// Add student to existing groups if possible
		for groupIndex, group := range groups {
			if canAddUnitToGroup(unit, group) {
				if DEBUG {
					fmt.Printf("Adding %v to group %s\n", unit, group)
				}
				groups[groupIndex] = append(group, unit...)
				return
			}
		}

		// Else create a new group
		if DEBUG {
			fmt.Printf("Creating new group for %v\n", unit)
		}
		groups = append(groups, slices.Clone(unit))
	}

	// Process inclusion groups and constrained students first
	for _, unit := range units {
		if DEBUG {
			fmt.Printf("Processing: %v\n", unit)
		}
		processUnit(unit)
		if DEBUG {
			fmt.Println("Current groups:", groups)
			fmt.Println()
		}
	}

	if DEBUG {
		fmt.Println("Final groups:", groups)
	}

	return groups, nil
}

// CreateGroups creates student groups based on the number of groups.
func createNumGroups(data *types.GroupingData, numGroups int) ([][]string, error) {
	groups := make([][]string, numGroups)
	exclusionLookup := buildExclusionLookup(data.Exclusions)

	if err := validateInclusionsAgainstExclusions(data.Inclusions, exclusionLookup); err != nil {
		return nil, err
	}

	units := buildAssignmentUnits(data.Students, data.Inclusions, exclusionLookup)

	canAddUnitToGroup := func(unit []string, group []string) bool {
		for _, student := range unit {
			for _, studentInGroup := range group {
				if studentsConflict(student, studentInGroup, exclusionLookup) {
					return false
				}
			}
		}
		return true
	}

	processUnit := func(unit []string) error {

		// Add student to existing groups if possible
		sort.Slice(groups, func(i, j int) bool {
			return len(groups[i]) < len(groups[j])
		})
		for groupIndex, group := range groups {
			if canAddUnitToGroup(unit, group) {
				if DEBUG {
					fmt.Printf("Adding %v to group %s\n", unit, group)
				}
				groups[groupIndex] = append(group, unit...)
				return nil
			}
		}
		return errors.New("exception and inclusion constraints cannot be met for this number of groups")
	}

	// Process inclusion groups and constrained students first
	for _, unit := range units {
		if DEBUG {
			fmt.Printf("Processing: %v\n", unit)
		}
		if err := processUnit(unit); err != nil {
			return nil, err
		}
		if DEBUG {
			fmt.Println("Current groups:", groups)
			fmt.Println()
		}
	}

	if DEBUG {
		fmt.Println("Final groups:", groups)
	}

	return groups, nil
}

func buildExclusionLookup(exclusions [][]string) map[string]map[string]struct{} {
	lookup := make(map[string]map[string]struct{})

	for _, exclusionGroup := range exclusions {
		for _, student := range exclusionGroup {
			if lookup[student] == nil {
				lookup[student] = make(map[string]struct{})
			}

			for _, otherStudent := range exclusionGroup {
				if otherStudent == student {
					continue
				}
				lookup[student][otherStudent] = struct{}{}
			}
		}
	}

	return lookup
}

func validateSubjectInclusions(inclusions [][]string, studentSubject map[string]string, exclusionLookup map[string]map[string]struct{}) error {
	if err := validateInclusionsAgainstExclusions(inclusions, exclusionLookup); err != nil {
		return err
	}

	for _, inclusionGroup := range inclusions {
		seenSubjects := make(map[string]string)
		for _, student := range inclusionGroup {
			subject := studentSubject[student]
			if firstStudent, exists := seenSubjects[subject]; exists {
				return fmt.Errorf("students %q and %q are required to be together but both belong to subject %q", firstStudent, student, subject)
			}
			seenSubjects[subject] = student
		}
	}

	return nil
}

func validateInclusionsAgainstExclusions(inclusions [][]string, exclusionLookup map[string]map[string]struct{}) error {
	for _, inclusionGroup := range inclusions {
		for i := 0; i < len(inclusionGroup); i++ {
			for j := i + 1; j < len(inclusionGroup); j++ {
				if studentsConflict(inclusionGroup[i], inclusionGroup[j], exclusionLookup) {
					return fmt.Errorf("students %q and %q are required to be together but are also listed in an exclusion group", inclusionGroup[i], inclusionGroup[j])
				}
			}
		}
	}

	return nil
}

func buildAssignmentUnits(students []string, inclusions [][]string, exclusionLookup map[string]map[string]struct{}) [][]string {
	units := make([][]string, 0, len(students))
	includedStudents := make(map[string]struct{}, len(students))

	for _, inclusionGroup := range inclusions {
		unit := slices.Clone(inclusionGroup)
		units = append(units, unit)
		for _, student := range inclusionGroup {
			includedStudents[student] = struct{}{}
		}
	}

	for _, student := range students {
		if _, exists := includedStudents[student]; exists {
			continue
		}
		units = append(units, []string{student})
	}

	rand.Shuffle(len(units), func(i, j int) {
		units[i], units[j] = units[j], units[i]
	})

	sort.SliceStable(units, func(i, j int) bool {
		iHasConstraints := unitHasExclusions(units[i], exclusionLookup)
		jHasConstraints := unitHasExclusions(units[j], exclusionLookup)
		if iHasConstraints != jHasConstraints {
			return iHasConstraints
		}

		if len(units[i]) != len(units[j]) {
			return len(units[i]) > len(units[j])
		}

		return false
	})

	return units
}

func unitHasExclusions(unit []string, exclusionLookup map[string]map[string]struct{}) bool {
	for _, student := range unit {
		if len(exclusionLookup[student]) > 0 {
			return true
		}
	}

	return false
}

func studentsConflict(student string, otherStudent string, exclusionLookup map[string]map[string]struct{}) bool {
	_, exists := exclusionLookup[student][otherStudent]
	return exists
}

func flattenSubjectStudentsBySubject(subjectStudents map[string][]string) []string {
	students := make([]string, 0)
	for _, subjectGroup := range subjectStudents {
		students = append(students, subjectGroup...)
	}

	return students
}
