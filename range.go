package xlsx

import (
	"github.com/mattn/go-ole"
	"github.com/mattn/go-ole/oleutil"
)

type Range ole.IDispatch

func (this *Range) ToString() string {
	if val := oleutil.MustGetProperty((*ole.IDispatch)(this), "Value"); nil != val {
		return val.ToString()
	}
	return ""
}
