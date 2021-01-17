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
	DummyXlsxGeneratorAdapter
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

func (a *xlsxTest) GetRows() (<-chan map[string]interface{}) {
	rows := make(chan map[string]interface{})
	go func() {
		for i := 0; i < 10; i++ {
			r := i+1
			row := make(map[string]interface{}, 5) // a, {b1,b2,b3}, c -> 5
			row["a"]    = fmt.Sprintf("a%d%d", r, 1)
			row["b_b1"] = fmt.Sprintf("b%d%d", r, 2)
			row["b_b2"] = fmt.Sprintf("b%d%d", r, 3)
			row["b_b3"] = fmt.Sprintf("b%d%d", r, 4)
			row["c"]    = fmt.Sprintf("c%d%d", r, 5)

			rows <- row
		}

		close(rows)
	}()

	return rows
}

