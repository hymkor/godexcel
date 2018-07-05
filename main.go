package main

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/zetamatta/pipe2excel/excel"
)

func main1(args []string) error {
	excel1, err := excel.New(true)
	if err != nil {
		return err
	}
	defer excel1.Release()

	var book1 *excel.Book
	if len(args) <= 0 {
		book1, err = excel1.NewBook()
	} else {
		fname, err := filepath.Abs(args[0])
		if err != nil {
			return err
		}
		book1, err = excel1.Open(fname)
	}
	if err != nil {
		return err
	}
	defer book1.Release()

	sheet, err := book1.Item(1)
	if err != nil {
		return err
	}
	_cell, err := sheet.GetProperty("Cells", 1, 1)
	if err != nil {
		return err
	}
	cell := _cell.ToIDispatch()
	cell.PutProperty("Value", "!!!!")
	cell.Release()
	return nil
}

func main() {
	if err := main1(os.Args[1:]); err != nil {
		fmt.Fprintln(os.Stderr, err.Error())
		os.Exit(1)
	}
	os.Exit(0)
}
