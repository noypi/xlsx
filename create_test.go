package xlsx

import (
	"io/ioutil"
	"os"

	"github.com/mattn/go-ole"
)

func ExampleCreate() {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	excel, err := CreateObject()
	if nil != err {
		return
	}
	defer excel.Release()

	workbooks := excel.Workbooks()

	workbook := workbooks.Create()
	defer workbook.Close()

	sheet := workbook.Worksheets(1)
	sheet.Range("A1").PutValue2("somevalue2")
	sheet.Range("A2").PutValue("bold")
	font := sheet.Range("A2").Font()
	font.SetBold(true)

	sheet.Range("A3").PutValue2(sheet.Range("A1").Font().GetBold())
	sheet.Range("A4").PutValue2(font.GetBold())

	sheet.Range("A1").Font().SetSize(14)

	sheet.Range("A7").PutValue2(sheet.Range("A1").Font().GetSize())
	sheet.Range("A8").PutValue2(sheet.Range("A1").Font().GetName())

	sheet.Range("A9").SetFormulaR1C1("=1+1")
	sheet.Range("A10").PutValue2("'" + sheet.Range("A9").GetFormulaR1C1())

	f, _ := ioutil.TempFile(os.TempDir(), "gotest-xlsx-")
	defer f.Close()
	workbook.Save(f.Name())

}
