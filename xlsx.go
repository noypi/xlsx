package xlsx

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type Xlsx ole.IDispatch
type Workbook ole.IDispatch

func CreateObject() (*Xlsx, error) {
	unknown, err := oleutil.CreateObject("Excel.Application")
	if nil != err {
		return nil, err
	}

	excel, err := unknown.QueryInterface(ole.IID_IDispatch)
	if nil != err {
		return nil, err
	}

	return (*Xlsx)(excel), nil
}

func (this *Xlsx) Release() {
	oleutil.MustCallMethod((*ole.IDispatch)(this), "Quit")
	(*ole.IDispatch)(this).Release()
}

func (this *Xlsx) Workbooks() *Workbooks {
	if v := oleutil.MustGetProperty((*ole.IDispatch)(this), "Workbooks"); nil != v {
		return (*Workbooks)(v.ToIDispatch())
	}
	return nil
}
