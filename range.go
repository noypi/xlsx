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

func (this *Range) PutValue(o interface{}) {
	oleutil.MustPutProperty((*ole.IDispatch)(this), "Value", o)
}

func (this *Range) PutValue2(o interface{}) {
	oleutil.MustPutProperty((*ole.IDispatch)(this), "Value2", o)
}

func (this *Range) Format(fmt string) {
	oleutil.MustPutProperty((*ole.IDispatch)(this), "Format", fmt)
}
