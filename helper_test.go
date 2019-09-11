package excelize

import (
	"testing"
)

func TestFillSheetCells(t *testing.T) {

	type Row struct {
		A string `cell:"A"`
		B string `cell:"B"`
		C int    `cell:"C"`
	}

	var sheets Sheets

	sheets = append(sheets,
		Sheet{
			Name: "data",
			Rows: []interface{}{
				&Row{A: "a1", B: "b1", C: 1},
				&Row{A: "a2", B: "b2", C: 2},
			},
		},
		Sheet{
			Name: "data2",
			Rows: []interface{}{
				&Row{A: "aa11", B: "bb11", C: 1},
				&Row{A: "aa22", B: "bb22", C: 2},
			},
		})

	file := NewFile()
	err := file.FillSheetCells(sheets)
	if err != nil {

		t.Error(err)
	}

	err = file.SaveAs("./test/fillcells.xlsx")
	if err != nil {

		t.Error(err)
	}
}
