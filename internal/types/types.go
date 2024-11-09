package types

type GroupingData struct {
	SubjectStudents map[string][]string
	Students        []string
	Exclusions      [][]string
}
