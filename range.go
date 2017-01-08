package xlsx

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type Range ole.IDispatch

func (this *Range) ToString() string {
	if val := oleutil.MustGetProperty((*ole.IDispatch)(this), "Text"); nil != val {
		s := val.ToString()
		bb := make([]byte, len([]byte(s)))
		copy(bb, []byte(s))
		return string(bb)
	}
	return ""
}

func (this *Range) Format(fmt string) {
	oleutil.MustPutProperty((*ole.IDispatch)(this), "Format", fmt)
}
