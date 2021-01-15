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
	GetRows() (<-chan interface{})

	/// 在输出每行之前调用，在这里可以做一些判断条件收集
	BeforeOutputRow(row interface{})

	/// 获取某一列/子列的值。如果没有子列, subIdx为-1
	GetColValue(row interface{}, idx, subIdx int, title Title) interface{}
}

func GenerateXlsx(xg XlsxGenerator) {
	xg.BeforeOutputXlsx()

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

	rows := xg.GetRows()
	if rows == nil {
		return
	}

	for row := range rows {
		outputRow(xg, fXls, sheet, titles, rowNo, row)
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

func outputRow(xg XlsxGenerator, fXls *excelize.File, sheet string, titles []Title, rowNo int, row interface{}) {
	xg.BeforeOutputRow(row)

	g := NewColumnGenerator()
	defer g.Stop()

	for i, title := range titles {
		if len(title.SubTitles) > 0 {
			for j, _ := range title.SubTitles {
				val := xg.GetColValue(row, i, j, title)
				outputCell(g, rowNo, fXls, sheet, val)
			}
		} else {
			val := xg.GetColValue(row, i, -1, title)
			outputCell(g, rowNo, fXls, sheet, val)
		}
	}
}

func outputCell(g *columnGenerator, rowNo int, fXls *excelize.File, sheet string, val interface{}) {
	col := g.Next()
	axis1 := fmt.Sprintf("%s%d", col, rowNo)
	fXls.SetCellValue(sheet, axis1, val)
}
