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
	file, err := os.Open("iss_trans.csv")
	if err != nil {
		check(err)
	}
	defer file.Close()
	t0 := time.Now()
	qf := qframe.ReadCSV(file, fcsv.Delimiter(byte('|')))
	fmt.Println("Read data duration : ", time.Since(t0))
	t0 = time.Now()
	qf = qf.Apply(qframe.Instruction{Fn: update, SrcCol1: "REVERSAL", SrcCol2: "ACQ_AMOUNT", DstCol: "ACQ_AMOUNT"})
	qf = qf.Apply(qframe.Instruction{Fn: update, SrcCol1: "REVERSAL", SrcCol2: "ISS_AMOUNT", DstCol: "ISS_AMOUNT"})
	fmt.Println("Update reversal duration : ", time.Since(t0))
	for _, v := range []int{1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008, 1010, 1011, 1012, 1013} {
		t0 = time.Now()
		To_Excel(qf, v)
		fmt.Printf("Duration write template %d: %v\n", v, time.Since(t0))
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

func To_Excel(frame qframe.QFrame, inst int) {
	qf := frame.Filter(qframe.Filter{Column: "ACQ_INST", Arg: inst, Comparator: "="})
	if qf.Err != nil {
		check(qf.Err)
	}
	//qf = qf.GroupBy(groupby.Co)
	qfi := frame.Filter(qframe.Filter{Column: "ISS_INST", Arg: inst, Comparator: "="})
	if qfi.Err != nil {
		check(qf.Err)
	}
	outfile := fmt.Sprintf("template/template_%d.xlsx", inst)
	f, err := excelize.OpenFile(outfile, excelize.Options{})
	check(err)
	f.DeleteSheet("Dataset")
	f.NewSheet("Dataset")
	named := qf.ColumnTypeMap()
	a := 'A'
	last := ""
	response := get_response()
	for _, col := range qf.ColumnNames() {
		err := f.SetCellValue("Dataset", fmt.Sprintf("%c1", a), col)
		if err != nil {
			panic(err)
		}
		switch named[col] {
		case types.String:
			view := qf.MustStringView(col)
			for i := 0; i < view.Len(); i++ {
				value := view.ItemAt(i)
				err := f.SetCellValue("Dataset", fmt.Sprintf("%c%d", a, i+2), *value)
				if err != nil {
					panic(err)
				}
			}
		case types.Int:
			view := qf.MustIntView(col)
			for i := 0; i < view.Len(); i++ {
				value := view.ItemAt(i)
				err := f.SetCellValue("Dataset", fmt.Sprintf("%c%d", a, i+2), value)
				if err != nil {
					panic(err)
				}
			}
			if col == "RESP" {
				a++
				err := f.SetCellValue("Dataset", fmt.Sprintf("%c1", a), "ACTION")
				if err != nil {
					panic(err)
				}
				for i := 0; i < view.Len(); i++ {
					value := view.ItemAt(i)
					err := f.SetCellValue("Dataset", fmt.Sprintf("%c%d", a, i+2), get_action(value))
					if err != nil {
						panic(err)
					}
				}
				a++
				err = f.SetCellValue("Dataset", fmt.Sprintf("%c1", a), "REJECTED_TYPE")
				if err != nil {
					panic(err)
				}
				for i := 0; i < view.Len(); i++ {
					value := view.ItemAt(i)
					err := f.SetCellValue("Dataset", fmt.Sprintf("%c%d", a, i+2), Get_desc(value, response))
					if err != nil {
						panic(err)
					}
				}
			}
		}
		a++
	}
	last = fmt.Sprintf("%c%d", a-1, qf.Len()+1)
	fmt.Println(last)
	err = f.AddTable("Dataset", "A1", last, `{
		"table_name": "acq_trans",
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
	os.Exit(0)
}
func update(a, b int) int {
	if a == 1 {
		return -b
	}
	return b
}
