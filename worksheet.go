package xls

import (
	"encoding/binary"
	"fmt"
	"io"
	"unicode/utf16"
)

type boundsheet struct {
	Filepos uint32
	Type    byte
	Visible byte
	Name    byte
}

//WorkSheet in one WorkBook
type WorkSheet struct {
	bs   *boundsheet
	wb   *WorkBook
	Name string
	rows map[uint16]*Row
	//NOTICE: this is the max row number of the sheet, so it should be count -1
	MaxRow uint16
	parsed bool
}

// Row returns the row at the specified index
func (w *WorkSheet) Row(i int) *Row {
	row := w.rows[uint16(i)]
	if row != nil {
		row.wb = w.wb
	}
	return row
}

func (w *WorkSheet) parse(buf io.ReadSeeker) error {
	w.rows = make(map[uint16]*Row)
	b := new(bof)
	var preBof *bof
	for {
		if err := binary.Read(buf, binary.LittleEndian, b); err == nil {
			preBof, err = w.parseBof(buf, b, preBof)
			if err != nil {
				return err
			}
			if b.ID == 0xa {
				break
			}
		} else {
			fmt.Println(err)
			break
		}
	}
	w.parsed = true
	return nil
}

func (w *WorkSheet) parseBof(buf io.ReadSeeker, b *bof, pre *bof) (*bof, error) {
	var col interface{}
	var err error
	switch b.ID {
	case 0x208: //ROW
		r := new(rowInfo)
		if err := binary.Read(buf, binary.LittleEndian, r); err != nil {
			return nil, err
		}
		w.addRow(r)
	case 0x0BD: //MULRK
		mc := new(MulrkCol)
		size := (b.Size - 6) / 6
		if err := binary.Read(buf, binary.LittleEndian, &mc.Col); err != nil {
			return nil, err
		}
		mc.Xfrks = make([]XfRk, size)
		for i := uint16(0); i < size; i++ {
			binary.Read(buf, binary.LittleEndian, &mc.Xfrks[i])
		}
		if err := binary.Read(buf, binary.LittleEndian, &mc.LastColB); err != nil {
			return nil, err
		}
		col = mc
	case 0x0BE: //MULBLANK
		mc := new(MulBlankCol)
		size := (b.Size - 6) / 2
		binary.Read(buf, binary.LittleEndian, &mc.Col)
		mc.Xfs = make([]uint16, size)
		for i := uint16(0); i < size; i++ {
			if err := binary.Read(buf, binary.LittleEndian, &mc.Xfs[i]); err != nil {
				return nil, err
			}
		}
		if err := binary.Read(buf, binary.LittleEndian, &mc.LastColB); err != nil {
			return nil, err
		}
		col = mc
	case 0x203: //NUMBER
		col = new(NumberCol)
		if err := binary.Read(buf, binary.LittleEndian, col); err != nil {
			return nil, err
		}
	case 0x06: //FORMULA
		c := new(FormulaCol)
		if err := binary.Read(buf, binary.LittleEndian, &c.Header); err != nil {
			return nil, err
		}
		c.Bts = make([]byte, b.Size-20)
		if err := binary.Read(buf, binary.LittleEndian, &c.Bts); err != nil {
			return nil, err
		}
		col = c
	case 0x27e: //RK
		col = new(RkCol)
		if err := binary.Read(buf, binary.LittleEndian, col); err != nil {
			return nil, err
		}
	case 0xFD: //LABELSST
		col = new(LabelsstCol)
		if err := binary.Read(buf, binary.LittleEndian, col); err != nil {
			return nil, err
		}
	case 0x204:
		c := new(labelCol)
		if err := binary.Read(buf, binary.LittleEndian, &c.BlankCol); err != nil {
			return nil, err
		}
		var count uint16
		if err := binary.Read(buf, binary.LittleEndian, &count); err != nil {
			return nil, err
		}
		c.Str, err = w.wb.getString(buf, count)
		if err != nil {
			return nil, err
		}
		col = c
	case 0x201: //BLANK
		col = new(BlankCol)
		if err := binary.Read(buf, binary.LittleEndian, col); err != nil {
			return nil, err
		}
	case 0x1b8: //HYPERLINK
		var hy HyperLink
		if err := binary.Read(buf, binary.LittleEndian, &hy.CellRange); err != nil {
			return nil, err
		}
		buf.Seek(20, 1)
		var flag uint32
		if err := binary.Read(buf, binary.LittleEndian, &flag); err != nil {
			return nil, err
		}
		var count uint32

		if flag&0x14 != 0 {
			if err := binary.Read(buf, binary.LittleEndian, &count); err != nil {
				return nil, err
			}
			hy.Description, err = b.utf16String(buf, count)
			if err != nil {
				return nil, err
			}
		}
		if flag&0x80 != 0 {
			if err := binary.Read(buf, binary.LittleEndian, &count); err != nil {
				return nil, err
			}
			hy.TargetFrame, err = b.utf16String(buf, count)
			if err != nil {
				return nil, err
			}
		}
		if flag&0x1 != 0 {
			var guid [2]uint64
			if err := binary.Read(buf, binary.BigEndian, &guid); err != nil {
				return nil, err
			}
			if guid[0] == 0xE0C9EA79F9BACE11 && guid[1] == 0x8C8200AA004BA90B { //URL
				hy.IsURL = true
				if err := binary.Read(buf, binary.LittleEndian, &count); err != nil {
					return nil, err
				}
				hy.URL, err = b.utf16String(buf, count/2)
				if err != nil {
					return nil, err
				}
			} else if guid[0] == 0x303000000000000 && guid[1] == 0xC000000000000046 { //URL{
				var upCount uint16
				if err := binary.Read(buf, binary.LittleEndian, &upCount); err != nil {
					return nil, err
				}
				if err := binary.Read(buf, binary.LittleEndian, &count); err != nil {
					return nil, err
				}
				bts := make([]byte, count)
				if err := binary.Read(buf, binary.LittleEndian, &bts); err != nil {
					return nil, err
				}
				hy.ShortedFilePath = string(bts)
				buf.Seek(24, 1)
				if err := binary.Read(buf, binary.LittleEndian, &count); err != nil {
					return nil, err
				}
				if count > 0 {
					if err := binary.Read(buf, binary.LittleEndian, &count); err != nil {
						return nil, err
					}
					buf.Seek(2, 1)
					hy.ExtendedFilePath, err = b.utf16String(buf, count/2+1)
					if err != nil {
						return nil, err
					}
				}
			}
		}
		if flag&0x8 != 0 {
			if err := binary.Read(buf, binary.LittleEndian, &count); err != nil {
				return nil, err
			}
			var bts = make([]uint16, count)
			if err := binary.Read(buf, binary.LittleEndian, &bts); err != nil {
				return nil, err
			}
			runes := utf16.Decode(bts[:len(bts)-1])
			hy.TextMark = string(runes)
		}

		w.addRange(&hy.CellRange, &hy)
	case 0x809:
		buf.Seek(int64(b.Size), 1)
	case 0xa:
	default:
		// log.Printf("Unknow %X,%d\n", b.Id, b.Size)
		buf.Seek(int64(b.Size), 1)
	}
	if col != nil {
		w.add(col)
	}
	return b, nil
}

func (w *WorkSheet) add(content interface{}) {
	if ch, ok := content.(contentHandler); ok {
		if col, ok := content.(Coler); ok {
			w.addCell(col, ch)
		}
	}

}

func (w *WorkSheet) addCell(col Coler, ch contentHandler) {
	w.addContent(col.Row(), ch)
}

func (w *WorkSheet) addRange(rang Ranger, ch contentHandler) {

	for i := rang.FirstRow(); i <= rang.LastRow(); i++ {
		w.addContent(i, ch)
	}
}

func (w *WorkSheet) addContent(rowNo uint16, ch contentHandler) {
	var row *Row
	var ok bool
	if row, ok = w.rows[rowNo]; !ok {
		info := new(rowInfo)
		info.Index = rowNo
		row = w.addRow(info)
	}
	row.cols[ch.FirstCol()] = ch
}

func (w *WorkSheet) addRow(info *rowInfo) (row *Row) {
	if info.Index > w.MaxRow {
		w.MaxRow = info.Index
	}
	var ok bool
	if row, ok = w.rows[info.Index]; ok {
		row.info = info
	} else {
		row = &Row{info: info, cols: make(map[uint16]contentHandler)}
		w.rows[info.Index] = row
	}
	return
}
