package xlsx

import (
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type Font ole.IDispatch

func (this *Font) SetBold(b bool) {
	oleutil.MustPutProperty((*ole.IDispatch)(this), "Bold", b)
}

func (this *Font) GetBold() bool {
	if val := oleutil.MustGetProperty((*ole.IDispatch)(this), "Bold"); nil != val {
		return val.Value().(bool)
	}

	panic("should not be here")
}

func (this *Font) SetItalic(b bool) {
	oleutil.MustPutProperty((*ole.IDispatch)(this), "Italic", b)
}

func (this *Font) GetItalic() bool {
	if val := oleutil.MustGetProperty((*ole.IDispatch)(this), "Italic"); nil != val {
		return val.Value().(bool)
	}

	panic("should not be here")
}

func (this *Font) SetName(s string) {
	oleutil.MustPutProperty((*ole.IDispatch)(this), "Name", s)
}

func (this *Font) GetName() string {
	if val := oleutil.MustGetProperty((*ole.IDispatch)(this), "Name"); nil != val {
		return val.Value().(string)
	}

	panic("should not be here")
}

func (this *Font) SetSize(n float64) {
	oleutil.MustPutProperty((*ole.IDispatch)(this), "Size", n)
}

func (this *Font) GetSize() float64 {
	if val := oleutil.MustGetProperty((*ole.IDispatch)(this), "Size"); nil != val {
		return val.Value().(float64)
	}

	panic("should not be here")
}
