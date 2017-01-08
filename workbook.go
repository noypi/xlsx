package xlsx

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func (this *Workbook) Close() {
	oleutil.MustCallMethod((*ole.IDispatch)(this), "Close", false)
}

func (this *Workbook) Worksheets(i int) *Worksheet {
	if v := oleutil.MustGetProperty((*ole.IDispatch)(this), "Worksheets", i); nil != v {
		return (*Worksheet)(v.ToIDispatch())
	}
	return nil
}

func (this *Workbook) WorksheetsCount() int {
	if v := oleutil.MustGetProperty((*ole.IDispatch)(this), "Worksheets"); nil != v {
		if v = oleutil.MustGetProperty(v.ToIDispatch(), "Count"); nil != v {
			return (int)(v.Val)
		}
	}
	return 0
}
