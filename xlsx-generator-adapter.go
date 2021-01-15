package toxlsx

import (
	"io"
	"os"
	"fmt"
)

// --- DummyGeneratorAdapter -----
type DummyGeneratorAdapter struct {}
func (a *DummyGeneratorAdapter) BeforeOutputXlsx() {}
func (a *DummyGeneratorAdapter) GetWriter() io.Writer { return nil; }
func (a *DummyGeneratorAdapter) GetSheet() string { return "Sheet1"; }
func (a *DummyGeneratorAdapter) GetTitles() []Title { return nil; }
func (a *DummyGeneratorAdapter) GetRows() (<-chan interface{}) { return nil; }
func (a *DummyGeneratorAdapter) BeforeOutputRow(row interface{}) {}
func (a *DummyGeneratorAdapter) GetColValue(row interface{}, idx, subIdx int, title Title) interface{} { return nil; }

// --- XlsxGeneratorAdapter -----
type XlsxGeneratorAdapter struct {
	DummyGeneratorAdapter
}

func (a *XlsxGeneratorAdapter) GetWriter() io.Writer {
	return os.Stdout
}

func (a *XlsxGeneratorAdapter) GetTitles() []Title {
	return []Title{
		Title{
			Name: "a",
		},
		Title{
			Name: "b",
		},
		Title{
			Name: "c",
		},
	}
}

func (a *XlsxGeneratorAdapter) GetRows() (<-chan interface{}) {
	rows := make(chan interface{})
	go func() {
		for i := 0; i < 10; i++ {
			row := make([]string, 3)
			for j := 0; j < 3; j++ {
				row[j] = fmt.Sprintf("%d%d", i+1, j+1)
			}
			rows <- row
		}

		close(rows)
	}()

	return rows
}

func (a *XlsxGeneratorAdapter) GetColValue(row interface{}, idx, subIdx int, title Title) interface{} {
	realRow, ok := row.([]string)
	if !ok {
		return ""
	}
	return realRow[idx]
}
