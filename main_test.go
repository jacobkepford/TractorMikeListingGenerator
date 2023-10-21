package main

import (
	"fmt"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestGetColumnHeaders(t *testing.T) {
	columnHeaders := GetColumnHeaders()

	if len(columnHeaders) == 0 {
		t.Error("Did not get column headers")
	}
}

func TestWritingColumnHeaders(t *testing.T) {
	excelFile := CreateWriteFile()
	defer excelFile.CloseFile()

	excelFile.columnHeaders = GetColumnHeaders()

	excelFile.WriteColumnHeaders()

	cellValue := excelFile.ReadCell("B1")

	if cellValue == "" {
		t.Errorf("Expected Cell to contain %q but contained %q instead", excelFile.columnHeaders[1], cellValue)
	}

}

func TestWritingVariableTypeCells(t *testing.T) {
	excelFile := CreateWriteFile()
	defer excelFile.CloseFile()

	readFile := CreateTestReadFile(t)
	defer readFile.CloseFile()

	readFile.GetDataByColumn()

	excelFile.WriteVariableCells(readFile.dataByColumn)

	cellValue := excelFile.ReadCell("B5")

	if cellValue != "variable" {
		t.Errorf("Expected cell value to be variable, instead got %q", cellValue)
	}

}

func TestWritingVariationTypeCells(t *testing.T) {
	excelFile := CreateWriteFile()
	defer excelFile.CloseFile()

	readFile := CreateTestReadFile(t)
	defer readFile.CloseFile()

	readFile.GetDataByColumn()

	excelFile.WriteVariationCells(readFile.dataByColumn)

	variation1 := excelFile.ReadCell("B6")
	variation2 := excelFile.ReadCell("B7")

	if (variation1 != "variation") || (variation2 != "variation") {
		t.Errorf("Expected cell values to be variation, instead got %q and %q", variation1, variation2)
	}
}

func TestWritingVariationSkuCells(t *testing.T) {
	excelFile := CreateWriteFile()
	defer excelFile.CloseFile()

	readFile := CreateTestReadFile(t)
	defer readFile.CloseFile()

	readFile.GetDataByColumn()

	excelFile.WriteVariationCells(readFile.dataByColumn)

	actualSku1 := excelFile.ReadCell("C6")
	actualSku2 := excelFile.ReadCell("C7")

	wantSku1 := "Test2-3-BOQF10HYD"
	wantSku2 := "Test2-3-BOQF10"

	if actualSku1 != wantSku1 || actualSku2 != wantSku2 {
		t.Errorf("Sku values are not valid Sku 1 should be %q but is %q, Sku 2 should be %q but is %q", wantSku1, actualSku1, wantSku2, actualSku2)
	}
}

func TestWritingVariableNameCells(t *testing.T) {
	excelFile := CreateWriteFile()
	defer excelFile.CloseFile()

	readFile := CreateTestReadFile(t)
	defer readFile.CloseFile()

	readFile.GetDataByColumn()

	excelFile.WriteVariableCells(readFile.dataByColumn)

	actualName := excelFile.ReadCell("D5")
	wantName := "Test2-0 Test2-1 Skid Steer Quick Attach Replacement Faceplate"

	if actualName != wantName {
		t.Errorf("Expected name to be %q, instead got %q", wantName, actualName)
	}
}

func CreateTestReadFile(t *testing.T) ExcelWorker {
	t.Helper()

	worker := ExcelWorker{file: excelize.NewFile()}
	rowData := [][]string{}

	for i := 0; i < 10; i++ {
		row := make([]string, 10)
		for j := 0; j < 10; j++ {
			row[j] = fmt.Sprintf("Test%d-%d", i, j)
		}
		rowData = append(rowData, row)
	}

	for index, row := range rowData {
		startingCell, err := excelize.CoordinatesToCellName(1, index+1)

		if err != nil {
			t.Fatal("Unable to convert to cell name")
		}
		worker.file.SetSheetRow(WorkingSheet, startingCell, &row)
	}
	return worker
}
