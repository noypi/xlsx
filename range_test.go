package xlsx

import (
	"fmt"
	"testing"

	"github.com/mattn/go-ole"
)

func TestRangeToString(t *testing.T) {
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

	filepath := "D:\\dev\\gopath\\src\\xlsx\\dummy.xlsx"
	workbooks := excel.Workbooks()

	workbook := workbooks.Open(filepath)
	if nil == workbook {
		t.Fatal("workbook is nil")
		return
	}
	defer workbook.Close()

	//
	sheet1 := workbook.Worksheets(1)
	r := sheet1.Range("a1")
	a1Val := r.ToString()

	if "a1" != a1Val {
		t.Fatal("a1Val=", a1Val)
	}
}
