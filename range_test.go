package xlsx

import (
	"fmt"

	"github.com/mattn/go-ole"
)

func ExampleRangeToString() {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	var err error
	defer func() {
		if nil != err {
			fmt.Println("err=", err)
		}
	}()

	excel, err := CreateObject()
	if nil != err {
		return
	}
	defer excel.Release()

	fpath := "m:\\gopath\\src\\github.com\\noypi\\xlsx\\dummy.xlsx"
	workbooks := excel.Workbooks()
	workbook := workbooks.Open(fpath)
	defer workbook.Close()

	//
	sheet1 := workbook.Worksheets(1)
	r := sheet1.Range("a1")
	a1Val := r.ToString()
	fmt.Println(a1Val)
}
