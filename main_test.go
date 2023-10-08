package main

import "testing"

func TestGetColumnHeaders(t *testing.T) {
	columnHeaders := GetColumnHeaders()

	if len(columnHeaders) == 0 {
		t.Error("Did not get column headers")
	}
}
