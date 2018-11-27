package xls

// FontInfo contains font information
type FontInfo struct {
	Height     uint16
	Flag       uint16
	Color      uint16
	Bold       uint16
	Escapement uint16
	Underline  byte
	Family     byte
	Charset    byte
	Notused    byte
	NameB      byte
}

// Font ...
type Font struct {
	Info *FontInfo
	Name string
}
