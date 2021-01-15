package toxlsx

import (
	"testing"
	"fmt"
	"log"
	"io"
	"os"
)

func TestXlsx(t *testing.T) {
	fp, err := os.Create("a.xlsx")
	if err != nil {
		t.Fatalf("%v\n", err)
		return
	}
	defer fp.Close()

	xg := &xlsxTest{fp: fp}
	GenerateXlsx(xg)
	log.Printf("testing TestXlsx done!\n")
}

type xlsxTest struct {
	XlsxGeneratorAdapter
	fp *os.File
}

func (t *xlsxTest) GetWriter() io.Writer {
	return t.fp
}

func (a *xlsxTest) GetTitles() []Title {
	return []Title{
		Title{
			Name: "a",
		},
		Title{
			Name: "b",
			SubTitles: []string {
				"b1", "b2", "b3",
			},
		},
		Title{
			Name: "c",
		},
	}
}

func (a *xlsxTest) GetRows() (<-chan interface{}) {
	rows := make(chan interface{})
	go func() {
		for i := 0; i < 10; i++ {
			row := make([]string, 5) // a, {b1,b2,b3}, c -> 5
			for j := 0; j < 5; j++ {
				row[j] = fmt.Sprintf("%d%d", i+1, j+1)
			}
			rows <- row
		}

		close(rows)
	}()

	return rows
}

func (t *xlsxTest) GetColValue(row interface{}, idx, subIdx int, title Title) interface{} {
	realRow, ok := row.([]string)
	if !ok {
		return ""
	}
	if subIdx < 0 {
		return fmt.Sprintf("%s_%d_%s", title.Name, idx, realRow[idx])
	}
	return fmt.Sprintf("%s_%d_%s", title.SubTitles[subIdx], subIdx, realRow[idx+subIdx])
}
