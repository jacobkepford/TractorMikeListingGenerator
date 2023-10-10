package main

import (
	"testing"
)

func TestGetColumnHeaders(t *testing.T) {
	columnHeaders := GetColumnHeaders()

	if len(columnHeaders) == 0 {
		t.Error("Did not get column headers")
	}
}

func TestWritingColumnHeaders(t *testing.T) {
	excelFile := CreateExcelFile()
	defer excelFile.CloseFile()

	excelFile.columnHeaders = GetColumnHeaders()

	excelFile.WriteColumnHeaders()

	cellValue := excelFile.ReadCell("B1")

	if cellValue == "" {
		t.Errorf("Expected Cell to contain %q but contained %q instead", excelFile.columnHeaders[1], cellValue)
	}

}

func TestWritingVariableTypeCells(t *testing.T) {
	excelFile := CreateExcelFile()
	defer excelFile.CloseFile()

	excelFile.WriteVariableCells(3)

	cellValue := excelFile.ReadCell("B5")

	if cellValue != "variable" {
		t.Errorf("Expected cell value to be variable, instead got %q", cellValue)
	}

}

func TestWritingVariationTypeCells(t *testing.T) {
	excelFile := CreateExcelFile()
	defer excelFile.CloseFile()

	excelFile.WriteVariationCells([]string{"Sku", "One", "Two"})

	variation1 := excelFile.ReadCell("B6")
	variation2 := excelFile.ReadCell("B7")

	if (variation1 != "variation") || (variation2 != "variation") {
		t.Errorf("Expected cell values to be variation, instead got %q and %q", variation1, variation2)
	}
}

func TestWritingVariationSkuCells(t *testing.T) {
	excelFile := CreateExcelFile()
	defer excelFile.CloseFile()

	excelFile.WriteVariationCells([]string{"Sku", "One", "Two"})

	actualSku1 := excelFile.ReadCell("C6")
	actualSku2 := excelFile.ReadCell("C7")

	wantSku1 := "Two-BOQF10HYD"
	wantSku2 := "Two-BOQF10"

	if actualSku1 != wantSku1 || actualSku2 != wantSku2 {
		t.Errorf("Sku values are not valid Sku 1 should be %q but is %q, Sku 2 should be %q but is %q", wantSku1, actualSku1, wantSku2, actualSku2)
	}
}
