package main

import (
	"testing"
)

func TestConvRC(t *testing.T) {
	row, col := ConvRC("B5")
	if row != 5 {
		t.Fatalf("row = %d\n", row)
		return
	}
	if col != 2 {
		t.Fatalf("col = %d\n", col)
		return
	}
}
