package xlsx

import (
	"github.com/mattn/go-ole"
	"github.com/mattn/go-ole/oleutil"
)

func (this *Workbook) Close() {
	oleutil.MustCallMethod((*ole.IDispatch)(this), "Close")
}

func (this *Workbook) Worksheets(i int) *Worksheet {
	if v := oleutil.MustGetProperty((*ole.IDispatch)(this), "Worksheets", i); nil != v {
		return (*Worksheet)(v.ToIDispatch())
	}
	return nil
}
