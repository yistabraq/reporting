package main

import (
	"encoding/csv"
	"fmt"
	"os"
	"strconv"
	"time"

	"github.com/tobgu/qframe"
	fcsv "github.com/tobgu/qframe/config/csv"
	"github.com/tobgu/qframe/types"
	"github.com/xuri/excelize/v2"
)

func main() {
	file, err := os.Open("vue.csv")
	if err != nil {
		check(err)
	}
	defer file.Close()
	t0 := time.Now()
	qf := qframe.ReadCSV(file, fcsv.Delimiter(byte('|')))
	fmt.Println(time.Since(t0))
	//fmt.Println(qf.Filter(qframe.Filter{Column: "ACQ_INST",Comparator: =}))
	for _, v := range []int{1008} {
		To_Excel(qf, v)
	}

}
func check(err error) {
	if err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
}

func get_response() map[int]string {
	f, err := os.Open("response.csv")
	if err != nil {
		check(err)
	}

	// remember to close the file at the end of the program
	defer f.Close()
	csvReader := csv.NewReader(f)
	data, err := csvReader.ReadAll()
	if err != nil {
		check(err)
	}
	output := make(map[int]string)
	for _, v := range data {
		key, err := strconv.Atoi(v[0])
		if err != nil {
			continue
		}
		output[key] = v[1]
	}
	return output
}

func Get_desc(input int, dict map[int]string) string {
	if val, ok := dict[input]; ok {
		return val
	}
	return "Approuved"
}
func get_action(value int) string {
	if value == -1 {
		return "Approuved"
	}
	return "Rejected"
}

func To_Excel(qf qframe.QFrame, inst int) {
	qf = qf.Filter(qframe.Filter{Column: "ACQ_INST", Arg: inst, Comparator: "="})
	if qf.Err != nil {
		check(qf.Err)
	}
	outfile := fmt.Sprintf("template/Dashboard.xlsx")
	f, err := excelize.OpenFile(outfile, excelize.Options{})
	check(err)
	f.DeleteSheet("Dataset")
	f.NewSheet("Dataset")
	named := qf.ColumnTypeMap()
	c := 'A'
	last := ""
	response := get_response()
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
			if col == "RESP" {
				c++
				err := f.SetCellValue("Dataset", fmt.Sprintf("%c1", c), "ACTION")
				if err != nil {
					panic(err)
				}
				for i := 0; i < view.Len(); i++ {
					value := view.ItemAt(i)
					err := f.SetCellValue("Dataset", fmt.Sprintf("%c%d", c, i+2), get_action(value))
					if err != nil {
						panic(err)
					}
					last = fmt.Sprintf("%c%d", c, i+2)
				}
				c++
				err = f.SetCellValue("Dataset", fmt.Sprintf("%c1", c), "REJECTED_TYPE")
				if err != nil {
					panic(err)
				}
				for i := 0; i < view.Len(); i++ {
					value := view.ItemAt(i)
					err := f.SetCellValue("Dataset", fmt.Sprintf("%c%d", c, i+2), Get_desc(value, response))
					if err != nil {
						panic(err)
					}
					last = fmt.Sprintf("%c%d", c, i+2)
				}
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
		check(err)
	}
	if err := f.SaveAs(outfile); err != nil {
		fmt.Println(err)
	}
}
