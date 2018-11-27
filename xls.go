package xls

import (
	"io"
	"os"

	"github.com/extrame/ole2"
)

// Open opens one xls file with the specified charset
func Open(file string, charset string) (*WorkBook, error) {
	fi, err := os.Open(file)
	if err != nil {
		return nil, err
	}

	return OpenReader(fi, charset)
}

// OpenWithCloser opens one xls file and return the closer
func OpenWithCloser(file string, charset string) (*WorkBook, io.Closer, error) {
	fi, err := os.Open(file)
	if err != nil {
		return nil, nil, err
	}
	wb, err := OpenReader(fi, charset)
	return wb, fi, err
}

// OpenReader opens a xls file from reader
func OpenReader(reader io.ReadSeeker, charset string) (wb *WorkBook, err error) {
	ole, err := ole2.Open(reader, charset)
	if err != nil {
		return nil, err
	}

	dir, err := ole.ListDir()
	if err != nil {
		return nil, err
	}

	var book *ole2.File
	var root *ole2.File
	for _, file := range dir {
		name := file.Name()
		if name == "Workbook" && book == nil {
			book = file
		}
		if name == "Book" {
			book = file
			// break
		}
		if name == "Root Entry" {
			root = file
		}
	}
	if book == nil {
		return wb, nil
	}
	return newWorkBookFromOle2(ole.OpenFile(book, root))

}
