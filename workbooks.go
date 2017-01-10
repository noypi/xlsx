package xlsx

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type Workbooks ole.IDispatch

func (this *Workbooks) Open(filepath string) *Workbook {
	if v := oleutil.MustCallMethod((*ole.IDispatch)(this), "Open", filepath); nil != v {
		return (*Workbook)(v.ToIDispatch())
	}
	return nil
}

func (this *Workbooks) Create() *Workbook {
	if v := oleutil.MustCallMethod((*ole.IDispatch)(this), "Add"); nil != v {
		return (*Workbook)(v.ToIDispatch())
	}
	return nil
}
