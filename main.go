package main

import (
	"flag"
	"fmt"
	"github.com/LindsayBradford/go-dbf/godbf"
	"github.com/tealeg/xlsx"
	"os"
)

func main() {
	var out string
	var enc string
	flag.StringVar(&out, "out", "", "result file. defualt folder with inputfile + inputfilename + .xlsx")
	flag.StringVar(&enc, "enc", "Cp866", "encoding. default Cp866")
	flag.Parse()

	inputPath := os.Args[1]
	println("open file ", inputPath)
	if out == "" {
		out = inputPath[:len(inputPath)-4] + ".xlsx"
		println("Out file: ", out)
	}

	err := export(enc, inputPath, out)
	if err != nil {
		fmt.Printf("%v\n", err)
	}
}

func export(enc, inputPath, outPath string) error {
	tbl, err := godbf.NewFromFile(inputPath, enc)
	if err != nil {
		return fmt.Errorf("Error on opening dbf file: %v", err)
	}

	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell
	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Sheet1")
	if err != nil {
		return fmt.Errorf("Error on adding new sheet: %v", err)
	}

	row = sheet.AddRow()
	for _, name := range tbl.FieldNames() {
		cell = row.AddCell()
		cell.Value = name
	}

	for i := 0; i < tbl.NumberOfRecords(); i++ {
		row = sheet.AddRow()
		for k := 0; k < len(tbl.Fields()); k++ {
			cell = row.AddCell()
			cell.Value = tbl.FieldValue(i, k)
		}
	}

	err = file.Save(outPath)
	if err != nil {
		return fmt.Errorf("Orror on saving xlsx: %v", err)
	}
	return nil
}
