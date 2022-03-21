package main

import (
	"fmt"
	"os"
	"time"

	"github.com/tobgu/qframe"
	"github.com/tobgu/qframe/config/csv"
	"github.com/tobgu/qframe/types"
	"github.com/xuri/excelize/v2"
)

func main() {
	file, err := os.Open("vue.csv")
	if err != nil {
		fmt.Println(err)
		return
	}
	delimiter := byte('|')
	t0 := time.Now()
	out := qframe.ReadCSV(file, csv.Delimiter(delimiter))
	fmt.Println(time.Since(t0))
	qf := out.Filter(qframe.Filter{Column: "ISS_INST", Arg: 1001, Comparator: "="})
	f := excelize.NewFile()
	named := qf.ColumnTypeMap()
	f.NewSheet("Dataset")
	c := 'A'
	last := ""
	for _, col := range qf.ColumnNames() {
		err := f.SetCellValue("Dataset", fmt.Sprintf("%c1", c), col)
		if err != nil {
			panic(err)
		}
		switch named[col] {
		case types.String:
			view := qf.MustStringView(col)
			for i := 0; i < view.Len(); i++ {
				value := view.ItemAt(i)
				err := f.SetCellValue("Dataset", fmt.Sprintf("%c%d", c, i+2), *value)
				if err != nil {
					panic(err)
				}
				last = fmt.Sprintf("%c%d", c, i+2)
			}
		case types.Int:
			view := qf.MustIntView(col)
			for i := 0; i < view.Len(); i++ {
				value := view.ItemAt(i)
				err := f.SetCellValue("Dataset", fmt.Sprintf("%c%d", c, i+2), value)
				if err != nil {
					panic(err)
				}
				last = fmt.Sprintf("%c%d", c, i+2)
			}
		}
		c++
	}
	err = f.AddTable("Dataset", "A1", last, `{
		"table_name": "table",
		"table_style": "TableStyleLight14",
		"show_first_column": true,
		"show_last_column": true,
		"show_row_stripes": false,
		"show_column_stripes": true
	}`)
	if err != nil {
		fmt.Println(err)
		return
	}
	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}
