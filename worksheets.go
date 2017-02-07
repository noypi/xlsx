package xlsx

import (
	"math"

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
	if fromPage < 1 {
		fromPage = 1
	}

	var ps = make([]interface{}, 3)
	ps[0] = fromPage
	if toPage < fromPage {
		ps[1] = math.MaxInt32
	} else {
		ps[1] = toPage
	}
	if copies <= 0 {
		copies = 1
	}
	ps[2] = copies
	if 0 == len(params) {
		params = ps
	} else {
		params = append(ps, params...)
	}
	oleutil.MustCallMethod((*ole.IDispatch)(this), "PrintOut", params...)
}
