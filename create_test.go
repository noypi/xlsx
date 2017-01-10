package xlsx

import (
	"testing"

	"github.com/mattn/go-ole"
)

func TestCreate(t *testing.T) {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	excel, err := CreateObject()
	if nil != err {
		return
	}
	defer excel.Release()

	workbooks := excel.Workbooks()

	workbook := workbooks.Create()
	if nil == workbook {
		t.Fatal("workbook is nil")
		return
	}
	defer workbook.Close()

	sheet := workbook.Worksheets(1)
	sheet.Range("A1").PutValue2("somevalue")

	filepath := "c:\\temp\\a.xlsx"
	workbook.Save(filepath)

}
