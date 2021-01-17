package toxlsx

import (
	"io"
	"os"
	"fmt"
)

// --- DummyXlsxGeneratorAdapter -----
type DummyXlsxGeneratorAdapter struct {}
func (a *DummyXlsxGeneratorAdapter) BeforeOutputXlsx() {}
func (a *DummyXlsxGeneratorAdapter) GetWriter() io.Writer { return nil; }
func (a *DummyXlsxGeneratorAdapter) GetSheet() string { return "Sheet1"; }
func (a *DummyXlsxGeneratorAdapter) GetTitles() []Title { return nil; }
func (a *DummyXlsxGeneratorAdapter) GetRows() (<-chan map[string]interface{}) { return nil; }

// --- XlsxGeneratorAdapter -----
type XlsxGeneratorAdapter struct {
	DummyXlsxGeneratorAdapter
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

func (a *XlsxGeneratorAdapter) GetRows() (<-chan map[string]interface{}) {
	rows := make(chan map[string]interface{})
	go func() {
		for i := 0; i < 10; i++ {
			row := make(map[string]interface{})
			for j := 0; j < 3; j++ {
				row[fmt.Sprintf("%c", 'a'+j)] = fmt.Sprintf("%d%d", i+1, j+1)
			}
			rows <- row
		}

		close(rows)
	}()

	return rows
}
