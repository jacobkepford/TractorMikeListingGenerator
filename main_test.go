package main

import "testing"

func TestGetColumnHeaders(t *testing.T) {
	columnHeaders := GetColumnHeaders()

	if len(columnHeaders) == 0 {
		t.Error("Did not get column headers")
	}
}

func TestWritingColumns(t *testing.T) {
	excelFile := CreateExcelFile()
	defer excelFile.CloseFile()

	excelFile.columnHeaders = GetColumnHeaders()

	excelFile.WriteColumnHeaders()

	firstColumn := excelFile.ReadCell("A1")

	if firstColumn == "" {
		t.Errorf("Expected Cell to contain %q but contained %q instead", excelFile.columnHeaders[0], firstColumn)
	}

}
