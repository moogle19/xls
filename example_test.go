package xls

import (
	"fmt"
)

func ExampleOpen() {
	if xlFile, err := Open("Table.xls", "utf-8"); err == nil {
		fmt.Println(xlFile.Author)
	}
}

func ExampleWorkBookNumberSheets() {
	if xlFile, err := Open("Table.xls", "utf-8"); err == nil {
		for i := 0; i < xlFile.NumSheets(); i++ {
			sheet := xlFile.GetSheet(i)
			fmt.Println(sheet.Name)
		}
	}
}

//Output: read the content of first two cols in each row
func ExampleWorkBookGetSheet() {
	if xlFile, err := Open("Table.xls", "utf-8"); err == nil {
		if sheet1 := xlFile.GetSheet(0); sheet1 != nil {
			fmt.Print("Total Lines ", sheet1.MaxRow, sheet1.Name)
			for i := 0; i <= (int(sheet1.MaxRow)); i++ {
				row1 := sheet1.Row(i)
				fmt.Print("\n", row1.Col(0), ",", row1.Col(1))
			}
		}
	}
}
