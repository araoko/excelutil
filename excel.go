package excelutil

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"regexp"
	"strconv"
)

type ExcelModel struct {
	headers []string
	data    [][]interface{}
}

func NewExcelModel(h []string, l int) *ExcelModel {
	//test h items for duplicates
	if h == nil || len(h) == 0 {
		return nil
	}
	obj := ExcelModel{make([]string, len(h)), make([][]interface{}, 0, l)}
	copy(obj.headers, h)
	return &obj
}

func (e *ExcelModel) AddRow(row map[string]interface{}) error {
	if len(row) != len(e.headers) {
		return fmt.Errorf("Items count mismatch error: headers has %d items, data has %d items", len(e.headers), len(row))
	}
	r := make([]interface{}, len(e.headers))
	for i, h := range e.headers {
		if val, present := row[h]; present {
			r[i] = val
			continue
		}
		return fmt.Errorf("Error: field %s Not present in data", h)
	}
	e.data = append(e.data, r)
	return nil
}

func (e *ExcelModel) Write2File(f *excelize.File, sheetName, topLeft string) (string, error) {
	if f == nil {
		return "", fmt.Errorf("cannot write to %s file", "nil")
	}
	si := f.GetSheetIndex(sheetName)
	if si == 0 {
		si = f.NewSheet(sheetName)
	}
	c, r := S2cr(topLeft)
	if c == -1 {
		return "", fmt.Errorf("cell addess  %s is wrong", topLeft)
	}
	writeStringRow2Excel(f, sheetName, e.headers, c, r)
	for _, row := range e.data {
		r++
		writeRow2Excel(f, sheetName, row, c, r)
	}
	return Cr2s(c+len(e.headers)-1, r), nil
}

func writeStringRow2Excel(x *excelize.File, sheet string, row []string, c, r int) {
	for _, cell := range row {
		x.SetCellStr(sheet, Cr2s(c, r), cell)
		c++
	}
}

func writeRow2Excel(x *excelize.File, sheet string, row []interface{}, c, r int) {
	for _, cell := range row {
		x.SetCellValue(sheet, Cr2s(c, r), cell)
		c++
	}
}

func Cr2s(c, r int) string {
	colString := excelize.ToAlphaString(c)
	rowString := strconv.FormatInt(int64(r), 10)
	return colString + rowString
}
func S2cr(s string) (int, int) {
	re := regexp.MustCompile(`^([A-Za-z]+)(\d+)$`)
	sa := re.FindStringSubmatch(s)
	//fmt.Println("sa= ", sa)

	if sa == nil || len(sa) != 3 {
		return -1, -1
	}

	c := excelize.TitleToNumber(sa[1])
	r, err := strconv.ParseInt(sa[2], 10, 32)
	if err != nil {
		return -1, -1
	}

	return int(c), int(r)

}

func AdvanceRow(current string, step int) (string, error) {
	c, r := S2cr(current)
	if c == -1 {
		return "", fmt.Errorf("Error: invalid format for axis value %v", current)
	}
	return Cr2s(c, r+step), nil
}

func AdvanceCol(current string, step int) (string, error) {
	c, r := S2cr(current)
	if c == -1 {
		return "", fmt.Errorf("Error: invalid format for axis value %v", current)
	}
	return Cr2s(c+step, r), nil
}
