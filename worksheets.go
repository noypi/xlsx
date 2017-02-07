package xlsx

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type Worksheet ole.IDispatch

func (this *Worksheet) Range(r string) (out *Range) {
	if v := oleutil.MustGetProperty((*ole.IDispatch)(this), "Range", r); nil != v {
		return (*Range)(v.ToIDispatch())
	}
	return nil
}

// fromPage = set to 0 for default
// toPage = set to 0 for default
// copies = default is 1.
func (this *Worksheet) PrintOut(fromPage, toPage, copies int, params ...interface{}) {
	oleutil.MustGetProperty((*ole.IDispatch)(this), "PrintOut", params...)
}
