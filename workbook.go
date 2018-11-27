package xls

import (
	"bytes"
	"encoding/binary"
	"io"
	"os"
	"unicode/utf16"

	"golang.org/x/text/encoding/charmap"
)

// WorkBook contains an Excel workbook
type WorkBook struct {
	Is5ver   bool
	Type     uint16
	Codepage uint16
	Xfs      []stXfData
	Fonts    []Font
	Formats  map[uint16]*Format
	//All the sheets from the workbook
	sheets        []*WorkSheet
	Author        string
	rs            io.ReadSeeker
	sst           []string
	continueUTF16 uint16
	continueRich  uint16
	continueAPSB  uint32
	dateMode      uint16
}

//read workbook from ole2 file
func newWorkBookFromOle2(rs io.ReadSeeker) (*WorkBook, error) {
	wb := &WorkBook{
		Formats: make(map[uint16]*Format),
		rs:      rs,
		sheets:  make([]*WorkSheet, 0),
	}
	if err := wb.Parse(rs); err != nil {
		return nil, err
	}
	return wb, nil
}

// Parse parses the given reader into the workbook
func (w *WorkBook) Parse(buf io.Reader) error {
	b := new(bof)
	preBof := new(bof)
	offset := 0
	for {
		if err := binary.Read(buf, binary.LittleEndian, b); err == nil {
			preBof, b, offset, err = w.parseBof(buf, b, preBof, offset)
			if err != nil {
				return err
			}

		} else {
			break
		}
	}
	return nil
}

func (w *WorkBook) addXf(xf stXfData) {
	w.Xfs = append(w.Xfs, xf)
}

func (w *WorkBook) addFont(font *FontInfo, buf io.Reader) {
	name, _ := w.getString(buf, uint16(font.NameB))
	w.Fonts = append(w.Fonts, Font{Info: font, Name: name})
}

func (w *WorkBook) addFormat(format *Format) {
	if w.Formats == nil {
		os.Exit(1)
	}
	w.Formats[format.Head.Index] = format
}

func (w *WorkBook) parseBof(buf io.Reader, b *bof, pre *bof, preOffset int) (after *bof, afterUsing *bof, offset int, err error) {
	after = b
	afterUsing = pre
	var bts = make([]byte, b.Size)
	if err := binary.Read(buf, binary.LittleEndian, bts); err != nil {
		return nil, nil, 0, err
	}
	bufItem := bytes.NewReader(bts)
	switch b.ID {
	case 0x809:
		bif := new(biffHeader)
		if err := binary.Read(bufItem, binary.LittleEndian, bif); err != nil {
			return nil, nil, 0, err
		}

		if bif.Ver != 0x600 {
			w.Is5ver = true
		}
		w.Type = bif.Type
	case 0x042: // CODEPAGE
		if err := binary.Read(bufItem, binary.LittleEndian, &w.Codepage); err != nil {
			return nil, nil, 0, err
		}
	case 0x3c: // CONTINUE
		if pre.ID == 0xfc {
			var size uint16
			var err error
			if w.continueUTF16 >= 1 {
				size = w.continueUTF16
				w.continueUTF16 = 0
			} else {
				if err := binary.Read(bufItem, binary.LittleEndian, &size); err != nil {
					return nil, nil, 0, err
				}
			}
			for err == nil && preOffset < len(w.sst) {
				var str string
				if size > 0 {
					str, err = w.getString(bufItem, size)
					w.sst[preOffset] = w.sst[preOffset] + str
				}

				if err == io.EOF {
					break
				} else if err != nil {
					return nil, nil, 0, err
				}

				preOffset++
				if err := binary.Read(bufItem, binary.LittleEndian, &size); err != nil {
					return nil, nil, 0, err
				}
			}
		}
		offset = preOffset
		after = pre
		afterUsing = b
	case 0xfc: // SST
		info := new(SstInfo)
		if err := binary.Read(bufItem, binary.LittleEndian, info); err != nil {
			return nil, nil, 0, err
		}
		w.sst = make([]string, info.Count)
		var size uint16
		var i = 0
		for ; i < int(info.Count); i++ {
			var err error
			if err = binary.Read(bufItem, binary.LittleEndian, &size); err == nil {
				var str string
				str, err = w.getString(bufItem, size)
				w.sst[i] = w.sst[i] + str
			}

			if err == io.EOF {
				break
			} else if err != nil {
				return nil, nil, 0, err
			}
		}
		offset = i
	case 0x85: // bOUNDSHEET
		var bs = new(boundsheet)
		if err := binary.Read(bufItem, binary.LittleEndian, bs); err != nil {
			return nil, nil, 0, err
		}
		// different for BIFF5 and BIFF8
		w.addSheet(bs, bufItem)
	case 0x0e0: // XF
		if w.Is5ver {
			xf := new(Xf5)
			if err := binary.Read(bufItem, binary.LittleEndian, xf); err != nil {
				return nil, nil, 0, err
			}
			w.addXf(xf)
		} else {
			xf := new(Xf8)
			if err := binary.Read(bufItem, binary.LittleEndian, xf); err != nil {
				return nil, nil, 0, err
			}
			w.addXf(xf)
		}
	case 0x031: // FONT
		f := new(FontInfo)
		if err := binary.Read(bufItem, binary.LittleEndian, f); err != nil {
			return nil, nil, 0, err
		}
		w.addFont(f, bufItem)
	case 0x41E: //FORMAT
		font := new(Format)
		if err := binary.Read(bufItem, binary.LittleEndian, &font.Head); err != nil {
			return nil, nil, 0, err
		}
		font.str, err = w.getString(bufItem, font.Head.Size)
		if err != nil {
			return nil, nil, 0, err
		}
		w.addFormat(font)
	case 0x22: //DATEMODE
		if err := binary.Read(bufItem, binary.LittleEndian, &w.dateMode); err != nil {
			return nil, nil, 0, err
		}
	}
	return
}
func decodeWindows1251(enc []byte) string {
	dec := charmap.Windows1251.NewDecoder()
	out, _ := dec.Bytes(enc)
	return string(out)
}
func (w *WorkBook) getString(buf io.Reader, size uint16) (res string, err error) {
	if w.Is5ver {
		var bts = make([]byte, size)
		_, err = buf.Read(bts)
		res = decodeWindows1251(bts)
	} else {
		var richtextNum = uint16(0)
		var phoneticSize = uint32(0)
		var flag byte
		err = binary.Read(buf, binary.LittleEndian, &flag)
		if flag&0x8 != 0 {
			err = binary.Read(buf, binary.LittleEndian, &richtextNum)
		} else if w.continueRich > 0 {
			richtextNum = w.continueRich
			w.continueRich = 0
		}
		if flag&0x4 != 0 {
			err = binary.Read(buf, binary.LittleEndian, &phoneticSize)
		} else if w.continueAPSB > 0 {
			phoneticSize = w.continueAPSB
			w.continueAPSB = 0
		}
		if flag&0x1 != 0 {
			var bts = make([]uint16, size)
			var i = uint16(0)
			for ; i < size && err == nil; i++ {
				err = binary.Read(buf, binary.LittleEndian, &bts[i])
			}
			runes := utf16.Decode(bts[:i])
			res = string(runes)
			if i < size {
				w.continueUTF16 = size - i + 1
			}
		} else {
			var bts = make([]byte, size)
			var n int
			n, err = buf.Read(bts)
			if uint16(n) < size {
				w.continueUTF16 = size - uint16(n)
				err = io.EOF
			}

			var bts1 = make([]uint16, n)
			for k, v := range bts[:n] {
				bts1[k] = uint16(v)
			}
			runes := utf16.Decode(bts1)
			res = string(runes)
		}
		if richtextNum > 0 {
			var bts []byte
			var seekSize int64
			if w.Is5ver {
				seekSize = int64(2 * richtextNum)
			} else {
				seekSize = int64(4 * richtextNum)
			}
			bts = make([]byte, seekSize)
			err = binary.Read(buf, binary.LittleEndian, bts)
			if err == io.EOF {
				w.continueRich = richtextNum
			}
		}
		if phoneticSize > 0 {
			bts := make([]byte, phoneticSize)
			err = binary.Read(buf, binary.LittleEndian, bts)
			if err == io.EOF {
				w.continueAPSB = phoneticSize
			}
		}
	}
	return
}

func (w *WorkBook) addSheet(sheet *boundsheet, buf io.Reader) {
	name, _ := w.getString(buf, uint16(sheet.Name))
	w.sheets = append(w.sheets, &WorkSheet{bs: sheet, Name: name, wb: w})
}

//reading a sheet from the compress file to memory, you should call this before you try to get anything from sheet
func (w *WorkBook) prepareSheet(sheet *WorkSheet) {
	w.rs.Seek(int64(sheet.bs.Filepos), 0)
	sheet.parse(w.rs)
}

// GetSheet gets one sheet by its number
func (w *WorkBook) GetSheet(num int) *WorkSheet {
	if num >= len(w.sheets) {
		return nil
	}
	s := w.sheets[num]
	if !s.parsed {
		w.prepareSheet(s)
	}
	return s
}

// NumSheets gets the number of all sheets
func (w *WorkBook) NumSheets() int {
	return len(w.sheets)
}

// ReadAllCells is a helper function to read all cells from file
// Notice: the max value is the limit of the max capacity of lines.
// Warning: the helper function will need big memeory if file is large.
func (w *WorkBook) ReadAllCells(max int) (res [][]string) {
	res = make([][]string, 0)
	for _, sheet := range w.sheets {
		if len(res) < max {
			max = max - len(res)
			w.prepareSheet(sheet)
			if sheet.MaxRow != 0 {
				leng := int(sheet.MaxRow) + 1
				if max < leng {
					leng = max
				}
				temp := make([][]string, leng)
				for k, row := range sheet.rows {
					data := make([]string, 0)
					if len(row.cols) > 0 {
						for _, col := range row.cols {
							if uint16(len(data)) <= col.LastCol() {
								data = append(data, make([]string, col.LastCol()-uint16(len(data))+1)...)
							}
							str := col.String(w)

							for i := uint16(0); i < col.LastCol()-col.FirstCol()+1; i++ {
								data[col.FirstCol()+i] = str[i]
							}
						}
						if leng > int(k) {
							temp[k] = data
						}
					}
				}
				res = append(res, temp...)
			}
		}
	}
	return
}
