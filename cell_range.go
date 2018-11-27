package xls

import (
	"fmt"
)

// Ranger type of multi rows
type Ranger interface {
	FirstRow() uint16
	LastRow() uint16
}

// CellRange type of multi cells in multi rows
type CellRange struct {
	FirstRowB uint16
	LastRowB  uint16
	FristColB uint16
	LastColB  uint16
}

// FirstRow returns the first row index
func (c *CellRange) FirstRow() uint16 {
	return c.FirstRowB
}

// LastRow returns the last row index
func (c *CellRange) LastRow() uint16 {
	return c.LastRowB
}

// FirstCol returns the first column index
func (c *CellRange) FirstCol() uint16 {
	return c.FristColB
}

// LastCol returns the last column index
func (c *CellRange) LastCol() uint16 {
	return c.LastColB
}

// HyperLink type's content
type HyperLink struct {
	CellRange
	Description      string
	TextMark         string
	TargetFrame      string
	URL              string
	ShortedFilePath  string
	ExtendedFilePath string
	IsURL            bool
}

//get the hyperlink string, use the public variable Url to get the original Url
func (h *HyperLink) String(wb *WorkBook) []string {
	res := make([]string, h.LastColB-h.FristColB+1)
	var str string
	if h.IsURL {
		str = fmt.Sprintf("%s(%s)", h.Description, h.URL)
	} else {
		str = h.ExtendedFilePath
	}

	for i := uint16(0); i < h.LastColB-h.FristColB+1; i++ {
		res[i] = str
	}
	return res
}
