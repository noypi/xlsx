package xlsx

import (
	"github.com/mattn/go-ole"
	"github.com/mattn/go-ole/oleutil"
)

type Workbooks ole.IDispatch

func (this *Workbooks) Open(filepath string) *Workbook {
	if v := oleutil.MustCallMethod((*ole.IDispatch)(this), "Open", filepath); nil != v {
		return (*Workbook)(v.ToIDispatch())
	}
	return nil
}
