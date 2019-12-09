package excelize

import (
	"fmt"
	"testing"
	"time"
)

func TestFillSheetCells(t *testing.T) {

	type Row struct {
		A string `cell:"A" title:"TITLE-A"`
		B string `cell:"B"`
		C int    `cell:"C" title:"TITLE-C"`
		D string `cell:"-" title:"DDDDD"`
		E string
		F time.Time `cell:"D" title:"time" time_format:"-"`
		G time.Time `cell:"E" title:"time" time_format:"2006/01/02 15:04:05"`
	}

	var sheets Sheets

	sheets = append(sheets,
		&Sheet{
			Name: "data",
			Rows: []interface{}{
				&Row{A: "a1", B: "b1", C: 1, D: "d1", F: time.Now(), G: time.Now()},
				&Row{A: "a2", B: "b2", C: 2, D: "d2"},
				&Row{A: "a3", B: "b3", C: 2, D: "d3"},
				&Row{A: "a4", B: "b4", C: 2, D: "d4"},
				&Row{A: "a5", B: "b5", C: 2, D: "d5", F: time.Now()},
			},
			SheetMaxRows: 2,
			InitTitle:    true,
		},
		&Sheet{
			Name: "atad",
			Rows: []interface{}{
				&Row{A: "a1", B: "b1", C: 1, D: "d1"},
				&Row{A: "a2", B: "b2", C: 2, D: "d2"},
				&Row{A: "a3", B: "b3", C: 2, D: "d3"},
				&Row{A: "a4", B: "b4", C: 2, D: "d4"},
				&Row{A: "a5", B: "b5", C: 2, D: "d5"},
			},
			SheetMaxRows: 3,
			InitTitle:    true,
		})

	err := SaveToXlsx(sheets, "./test/fillcells.xlsx")
	if err != nil {

		t.Error(err)
	}
}

func TestPrintInfo(t *testing.T) {

	PrintInfo("./test/test.xlsx1")

	f, err := OpenFile("./test/test1.xlsx")
	if err != nil {

		t.Error(err)
	}

	rows, err := f.GetRows("data")
	if err != nil {

		t.Error(err)
	}

	t.Log(len(rows))

	err = f.SaveAs("./test/test2.xlsx")
	if err != nil {

		panic(err)
	}
}

func TestFillSheetCells2(t *testing.T) {

	type Row struct {
		A string `cell:"A" title:"TITLE-A"`
		B string `cell:"B"`
		C int    `cell:"C" title:"TITLE-C"`
		D string `cell:"D" title:"DDDDD"`
		E string `cell:"E" title:"Ee"`
	}

	var sheets Sheets
	sheet := &Sheet{
		Name:         "data",
		SheetMaxRows: 100000,
		InitTitle:    true,
	}

	for i := 1; i <= 2000000; i++ {

		sheet.Rows = append(sheet.Rows, &Row{
			A: fmt.Sprintf("%dioueoiug8935jd9oshg893qoq3ij", i),
			B: fmt.Sprintf("%dioueoiug8935jd9oshg893qoq3ij", i),
			C: i,
			D: fmt.Sprintf("%dioueoiug8935jd9oshg893qoq3ij", i),
			E: fmt.Sprintf("%dioueoiug8935jd9oshg893qoq3ij", i),
		})
	}

	sheets = append(sheets, sheet)

	err := SaveToXlsx(sheets, "./test/fillcells2.xlsx")
	if err != nil {

		t.Error(err)
	}
}

func TestRows_ReadStruct(t *testing.T) {

	f, err := OpenFile("./test/test22.xlsx")
	if err != nil {

		t.Error(err)
	}

	rows, err := f.Rows("data")
	if err != nil {

		t.Error(err)
	}

	type TestRow struct {
		H string    `cell:"A"`
		W string    `cell:"B"`
		F string    `cell:"C"`
		S string    `cell:"D"`
		T time.Time `cell:"E" time_format:"20060102150405"`
		E string    `cell:"F"`
	}

	for rows.Next() {

		row := &TestRow{}

		err = rows.ReadStruct(row)
		if err != nil {

			t.Error(err)
		}

		t.Log(row)
	}
}
