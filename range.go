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

func (this *Range) Font() (out *Font) {
	if v := oleutil.MustGetProperty((*ole.IDispatch)(this), "Font"); nil != v {
		return (*Font)(v.ToIDispatch())
	}
	return nil
}

func (this *Range) PutRowHeight(o interface{}) {
	oleutil.MustPutProperty((*ole.IDispatch)(this), "RowHeight", o)
}

func (this *Range) PutValue(o interface{}) {
	oleutil.MustPutProperty((*ole.IDispatch)(this), "Value", o)
}

func (this *Range) PutValue2(o interface{}) {
	oleutil.MustPutProperty((*ole.IDispatch)(this), "Value2", o)
}

func (this *Range) SetFormulaR1C1(s string) {
	oleutil.MustPutProperty((*ole.IDispatch)(this), "FormulaR1C1", s)
}

func (this *Range) GetFormulaR1C1() string {
	if val := oleutil.MustGetProperty((*ole.IDispatch)(this), "FormulaR1C1"); nil != val {
		return val.Value().(string)
	}

	panic("should not be here")
}
