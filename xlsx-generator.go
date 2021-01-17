package toxlsx

import (
	"github.com/360EntSecGroup-Skylar/excelize"
	"fmt"
	"io"
)

type Title struct {
	Name string
	SubTitles []string
}

func NewTitle(title string) Title {
	return Title{Name: title}
}

func NewTitleWithSubTitles(title string, subTitles []string) Title {
	return Title{
		Name: title,
		SubTitles: subTitles,
	}
}

func NewTitles(title ...string) []Title {
	if len(title) == 0 {
		return nil
	}
	titles := make([]Title, len(title))
	for i, t := range title {
		titles[i] = Title{Name:t}
	}
	return titles
}

type XlsxGenerator interface {
	/// 在输出整个xlsx之前调用，在这里可以做一些输出准备工作
	BeforeOutputXlsx()

	/// 获取输出目标
	GetWriter() io.Writer

	/// 获取Book的sheet名称
	GetSheet() string

	/// 获取的标题及子标题
	GetTitles() []Title

	// 获取所有的输出行channel
	GetRows() (<-chan map[string]interface{}) // 如果没有subtitle，key={title}; 其它key为{title}_{subtitle}
}

func GenerateXlsx(xg XlsxGenerator) {
	xg.BeforeOutputXlsx()

	rowsHandled := false
	rows := xg.GetRows()
	if rows == nil {
		return
	}
	defer func() {
		if rowsHandled {
			return
		}
		// rows必须读完，防止channel堵塞
		for _ = range rows {}
	}()

	writer := xg.GetWriter()
	if writer == nil {
		return
	}

	titles := xg.GetTitles()
	if len(titles) == 0 {
		return
	}

	fXls := excelize.NewFile()
	defer fXls.Write(writer)
	sheet := xg.GetSheet()
	rowNo := outputTitleRow(fXls, sheet, titles)
	rowsHandled = true

	for row := range rows {
		outputRow(fXls, sheet, titles, rowNo, row)
		rowNo += 1
	}
}

func outputTitleRow(fXls *excelize.File, sheet string, titles []Title) int {
	g := NewColumnGenerator()
	defer g.Stop()

	shouldMergeSubtitle := false
	for _, title := range titles {
		if len(title.SubTitles) > 0 {
			shouldMergeSubtitle = true
			break
		}
	}

	for _, title := range titles {
		if len(title.SubTitles) == 0 {
			outputTitleCell(g, fXls, sheet, title.Name, shouldMergeSubtitle)
		} else {
			outputTitleCellWithSubtitle(g, fXls, sheet, title.Name, title.SubTitles)
		}
	}

	lastCol := g.Last()

	style, _ := fXls.NewStyle(`{"alignment":{"horizontal":"center"}}`)
	if shouldMergeSubtitle {
		fXls.SetCellStyle(sheet, "A1", fmt.Sprintf("%s2", lastCol), style)
		return 3
	}
	fXls.SetCellStyle(sheet, "A1", fmt.Sprintf("%s1", lastCol), style)
	return 2
}

func outputTitleCell(g *columnGenerator, fXls *excelize.File, sheet, title string, shouldMergeSubtitle bool) {
	col := g.Next()
	axis1 := fmt.Sprintf("%s1", col)
	fXls.SetCellValue(sheet, axis1, title)
	if shouldMergeSubtitle {
		axis2 := fmt.Sprintf("%s2", col)
		fXls.SetCellValue(sheet, axis2, "")
		fXls.MergeCell(sheet, axis1, axis2)
	}
}

func outputTitleCellWithSubtitle(g *columnGenerator, fXls *excelize.File, sheet, title string, subtitles []string) {
	col := g.Next()
	axis1 := fmt.Sprintf("%s1", col)
	fXls.SetCellValue(sheet, axis1, title)
	var col2 string
	for i, st := range subtitles {
		if i == 0 {
			col2 = col
		} else {
			col2 = g.Next()
		}
		axis2 := fmt.Sprintf("%s2", col2)
		fXls.SetCellValue(sheet, axis2, st)
	}
	axis2 := fmt.Sprintf("%s1", col2)
	fXls.MergeCell(sheet, axis1, axis2)
}

func outputRow(fXls *excelize.File, sheet string, titles []Title, rowNo int, row map[string]interface{}) {
	g := NewColumnGenerator()
	defer g.Stop()

	for _, title := range titles {
		if len(title.SubTitles) > 0 {
			for _, subTitle := range title.SubTitles {
				key := fmt.Sprintf("%s_%s", title.Name, subTitle)
				val, ok := row[key]
				if !ok {
					val = ""
				}
				outputCell(g, rowNo, fXls, sheet, val)
			}
		} else {
			val, ok := row[title.Name]
			if !ok {
				val = ""
			}
			outputCell(g, rowNo, fXls, sheet, val)
		}
	}
}

func outputCell(g *columnGenerator, rowNo int, fXls *excelize.File, sheet string, val interface{}) {
	col := g.Next()
	axis1 := fmt.Sprintf("%s%d", col, rowNo)
	fXls.SetCellValue(sheet, axis1, val)
}
