package main

import (
	"fmt"
	"strconv"

	"github.com/xuri/excelize/v2"
)

const DataStartingRow int = 2

type ExcelWorker struct {
	file          *excelize.File
	columnHeaders []string
}

func (e *ExcelWorker) CloseFile() {
	if err := e.file.Close(); err != nil {
		fmt.Println(err)
	}
}

func (e *ExcelWorker) ReadCell(cell string) string {
	cell, err := e.file.GetCellValue("Sheet1", cell)
	if err != nil {
		fmt.Println(err)
		return ""
	}
	return cell
}

func (e *ExcelWorker) WriteCell(cell, cellValue string) {
	e.file.SetCellValue("Sheet1", cell, cellValue)
}

func (e *ExcelWorker) WriteColumnHeaders() {
	e.file.SetSheetRow("Sheet1", "A1", &e.columnHeaders)
}

func (e *ExcelWorker) WriteVariableTypeCells(skuLength int) {
	rowValue := DataStartingRow
	for i := 0; i < skuLength; i++ {
		rowString := strconv.Itoa(rowValue)
		cell := "B" + rowString
		e.WriteCell(cell, "variable")
		rowValue += 3
	}
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

func CreateExcelFile() ExcelWorker {
	return ExcelWorker{file: excelize.NewFile()}
}

func main() {
	f, err := excelize.OpenFile("Book1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		// Close the spreadsheet.
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
}
