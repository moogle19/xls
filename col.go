package xls

import (
	"math"
	"strconv"
	"strings"

	"time"

	"github.com/extrame/goyymmdd"
)

//content type
type contentHandler interface {
	String(*WorkBook) []string
	FirstCol() uint16
	LastCol() uint16
}

// Col is a worksheet column
type Col struct {
	RowB      uint16
	FirstColB uint16
}

// Coler defines a generic column
type Coler interface {
	Row() uint16
}

// Row returns the row index of the column
func (c *Col) Row() uint16 {
	return c.RowB
}

// FirstCol returns the first column
func (c *Col) FirstCol() uint16 {
	return c.FirstColB
}

// LastCol returns the last column
func (c *Col) LastCol() uint16 {
	return c.FirstColB
}

func (c *Col) String(wb *WorkBook) []string {
	return []string{"default"}
}

// XfRk ...
type XfRk struct {
	Index uint16
	Rk    RK
}

func (xf *XfRk) String(wb *WorkBook) string {
	idx := int(xf.Index)
	if len(wb.Xfs) > idx {
		fNo := wb.Xfs[idx].formatNo()
		if fNo >= 164 { // user defined format
			if formatter := wb.Formats[fNo]; formatter != nil {
				if strings.Contains(formatter.str, "#") || strings.Contains(formatter.str, ".00") {
					//If format contains # or .00 then this is a number
					return xf.Rk.String()
				}
				i, f, isFloat := xf.Rk.number()
				if !isFloat {
					f = float64(i)
				}
				t := timeFromExcelTime(f, wb.dateMode == 1)

				return yymmdd.Format(t, formatter.str)
			}
			// see http://www.openoffice.org/sc/excelfileformat.pdf Page #174
		} else if 14 <= fNo && fNo <= 17 || fNo == 22 || 27 <= fNo && fNo <= 36 || 50 <= fNo && fNo <= 58 { // jp. date format
			i, f, isFloat := xf.Rk.number()
			if !isFloat {
				f = float64(i)
			}
			t := timeFromExcelTime(f, wb.dateMode == 1)
			return t.Format(time.RFC3339) //TODO it should be international
		}
	}
	return xf.Rk.String()
}

// RK ...
type RK uint32

func (rk RK) number() (intNum int64, floatNum float64, isFloat bool) {
	multiplied := rk & 1
	isInt := rk & 2
	val := rk >> 2
	if isInt == 0 {
		isFloat = true
		floatNum = math.Float64frombits(uint64(val) << 34)
		if multiplied != 0 {
			floatNum = floatNum / 100
		}
		return
	}
	//+++ add lines from here
	if multiplied != 0 {
		isFloat = true
		floatNum = float64(val) / 100
		return
	}
	//+++end
	return int64(val), 0, false
}

func (rk RK) String() string {
	i, f, isFloat := rk.number()
	if isFloat {
		return strconv.FormatFloat(f, 'f', -1, 64)
	}
	return strconv.FormatInt(i, 10)
}

// Float returns the column value as float
func (rk RK) Float() (float64, error) {
	i, f, isFloat := rk.number()
	if !isFloat {
		return float64(i), nil
	}
	return f, nil
}

// MulrkCol ...
type MulrkCol struct {
	Col
	Xfrks    []XfRk
	LastColB uint16
}

// LastCol returns the last column
func (c *MulrkCol) LastCol() uint16 {
	return c.LastColB
}

func (c *MulrkCol) String(wb *WorkBook) []string {
	var res = make([]string, len(c.Xfrks))
	for i := 0; i < len(c.Xfrks); i++ {
		xfrk := c.Xfrks[i]
		res[i] = xfrk.String(wb)
	}
	return res
}

// MulBlankCol ...
type MulBlankCol struct {
	Col
	Xfs      []uint16
	LastColB uint16
}

// LastCol ...
func (c *MulBlankCol) LastCol() uint16 {
	return c.LastColB
}

func (c *MulBlankCol) String(wb *WorkBook) []string {
	return make([]string, len(c.Xfs))
}

// NumberCol ...
type NumberCol struct {
	Col
	Index uint16
	Float float64
}

func (c *NumberCol) String(wb *WorkBook) []string {
	return []string{strconv.FormatFloat(c.Float, 'f', -1, 64)}
}

// FormulaCol ...
type FormulaCol struct {
	Header struct {
		Col
		IndexXf uint16
		Result  [8]byte
		Flags   uint16
		_       uint32
	}
	Bts []byte
}

func (c *FormulaCol) String(wb *WorkBook) []string {
	return []string{"FormulaCol"}
}

// RkCol ...
type RkCol struct {
	Col
	Xfrk XfRk
}

func (c *RkCol) String(wb *WorkBook) []string {
	return []string{c.Xfrk.String(wb)}
}

// LabelsstCol ...
type LabelsstCol struct {
	Col
	Xf  uint16
	Sst uint32
}

func (c *LabelsstCol) String(wb *WorkBook) []string {
	return []string{wb.sst[int(c.Sst)]}
}

type labelCol struct {
	BlankCol
	Str string
}

func (c *labelCol) String(wb *WorkBook) []string {
	return []string{c.Str}
}

// BlankCol ...
type BlankCol struct {
	Col
	Xf uint16
}

func (c *BlankCol) String(wb *WorkBook) []string {
	return []string{""}
}
