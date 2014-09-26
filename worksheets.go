package xlsx

import (
	"github.com/mattn/go-ole"
	"github.com/mattn/go-ole/oleutil"
)

type Worksheet ole.IDispatch

func (this *Worksheet) Range(r string) (out *Range) {
	if v := oleutil.MustGetProperty((*ole.IDispatch)(this), "Range", r); nil != v {
		return (*Range)(v.ToIDispatch())
	}
	return nil
}