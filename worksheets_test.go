package xlsx_test

import (
	"log"

	"github.com/go-ole/go-ole"
	"github.com/noypi/xlsx"
)

func ExamplePrintOut() {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	excel, err := xlsx.CreateObject()
	if nil != err {
		log.Fatal(err)
	}
	defer excel.Release()

	workbooks := excel.Workbooks()
	workbook := workbooks.Create()
	defer workbook.Close()

	//
	sheet1 := workbook.Worksheets(1)
	r := sheet1.Range("a1")
	r.PutValue("adrian guwapo")

	sheet1.PrintOut(0, 0, 0)

}
