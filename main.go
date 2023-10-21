package main

import (
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

const DataStartingRow int = 2
const WorkingSheet = "Sheet1"

type ExcelWorker struct {
	file          *excelize.File
	columnHeaders []string
	dataByColumn  [][]string
}

func (e *ExcelWorker) CloseFile() {
	if err := e.file.Close(); err != nil {
		fmt.Println(err)
	}
}

func (e *ExcelWorker) ReadCell(cell string) string {
	cell, err := e.file.GetCellValue(WorkingSheet, cell)
	if err != nil {
		fmt.Println(err)
		return ""
	}
	return cell
}

func (e *ExcelWorker) WriteCell(cell, cellValue string) {
	e.file.SetCellValue(WorkingSheet, cell, cellValue)
}

func (e *ExcelWorker) WriteColumnHeaders() {
	e.file.SetSheetRow(WorkingSheet, "A1", &e.columnHeaders)
}

func (e *ExcelWorker) WriteVariableCells(readFileData [][]string) {
	columnLength := len(readFileData[0])
	rowValue := DataStartingRow
	for i := 0; i < columnLength-1; i++ {
		e.writeVariable(rowValue)
		e.writeVariableName(rowValue, readFileData[0][i+1], readFileData[1][i+1])
		rowValue += 3
	}
}

func (e *ExcelWorker) writeVariable(rowValue int) {
	cell := fmt.Sprintf("B%d", rowValue)
	e.WriteCell(cell, "variable")
}

func (e *ExcelWorker) writeVariableName(rowValue int, brandData, modelData string) {
	cell := fmt.Sprintf("D%d", rowValue)
	writeData := fmt.Sprintf("%s %s Skid Steer Quick Attach Replacement Faceplate", brandData, modelData)
	e.WriteCell(cell, writeData)
}

func (e *ExcelWorker) WriteVariationCells(skuData []string) {
	rowValue := DataStartingRow + 1

	for i := 0; i < len(skuData)-1; i++ {
		e.writeVariation(rowValue)
		e.writeSku(rowValue, 1, skuData[i+1])
		rowValue++
		e.writeVariation(rowValue)
		e.writeSku(rowValue, 2, skuData[i+1])
		rowValue += 2
	}
}

func (e *ExcelWorker) writeVariation(rowValue int) {
	cellValue := fmt.Sprintf("B%d", rowValue)
	e.WriteCell(cellValue, "variation")
}

func (e *ExcelWorker) writeSku(rowValue, variationCount int, skuData string) {
	cellValue := fmt.Sprintf("C%d", rowValue)
	if variationCount == 1 {
		e.WriteCell(cellValue, fmt.Sprintf("%s-BOQF10HYD", skuData))
		return
	}
	e.WriteCell(cellValue, fmt.Sprintf("%s-BOQF10", skuData))
}

func (e *ExcelWorker) GetDataByColumn() {
	dataByColumn, err := e.file.GetCols(WorkingSheet)

	if err != nil {
		log.Fatal("Unable to get data by column")
	}

	e.dataByColumn = dataByColumn
}

func GetColumnHeaders() []string {
	columnHeaders := []string{"ID", "Type", "SKU", "Name", "Published", "Is featured?", "Visibility in catalog", "Short description",
		"Description", "Date sale price starts", "Date sale price ends", "Tax status", "Tax class",
		"In stock?", "Stock", "Low stock amount", "Backorders allowed?", "Sold individually?", "Weight (lbs)",
		"Length (in)", "Width (in)", "Height (in)", "Allow customer reviews?", "Purchase note", "Sale price",
		"Regular price", "Categories", "Tags", "Shipping class", "Images", "Download limit", "Download expiry days",
		"Parent", "Grouped products", "Upsells", "Cross-sells", "External URL", "Button text", "Position", "Supplier Id",
		"Supplier Name", "Supplier Slug", "Supplier Description", "Supplier Email", "Supplier Account Number",
		"Attribute 1 name", "Attribute 1 value(s)", "Attribute 1 visible", "Attribute 1 global",
		"Meta: product_custom_field_amazon_affiliate_id", "Meta: product_custom_field_amazon_product_id",
		"Meta: number_of_orders", "Meta: ali_product_url", "Meta: ali_store_url", "Meta: ali_store_name",
		"Meta: ali_store_price_range", "Meta: ali_currency", "Meta: _cost_of_goods", "Meta: _custom_product_text_field_description",
		"Meta: ltl_freight_class", "Meta: _is_ali_product", "Meta: wfm_custom_content_area", "Meta: _wfm_custom_content_area",
		"Meta: _yoast_wpseo_primary_product_cat", "Meta: _yoast_wpseo_estimated-reading-time-minutes",
		"Meta: _yoast_wpseo_wordproof_timestamp", "Meta: wps_zoho_allow_background_syncing", "Meta: supplier",
		"Meta: supplierid", "Meta: _yoast_wpseo_content_score", "Meta: _yoast_wpseo_focuskw", "Meta: _yoast_wpseo_title",
		"Meta: _yoast_wpseo_metadesc", "Meta: _yoast_wpseo_linkdex", "Meta: custom_field", "Meta: custom_field_description"}
	return columnHeaders
}

func CreateWriteFile() ExcelWorker {
	return ExcelWorker{file: excelize.NewFile()}
}

func CreateReadFile() ExcelWorker {
	f, err := excelize.OpenFile("Skid Loader QA Replacements.xlsx")
	if err != nil {
		log.Fatal("Unable to find read file")
	}
	return ExcelWorker{file: f}
}

func main() {
	excelFile := CreateWriteFile()
	defer excelFile.CloseFile()

	readFile := CreateReadFile()
	defer readFile.CloseFile()

	excelFile.columnHeaders = GetColumnHeaders()
	excelFile.WriteColumnHeaders()

	readFile.GetDataByColumn()

	skuData := readFile.dataByColumn[3]
	excelFile.WriteVariableCells(readFile.dataByColumn)
	excelFile.WriteVariationCells(skuData)

	excelFile.file.SaveAs("Book1.xlsx")
}
