package excelize

import (
	"testing"
)

func TestFillSheetCells(t *testing.T) {

	type Row struct {
		A string `cell:"A" title:"TITLE-A"`
		B string `cell:"B"`
		C int    `cell:"C" title:"TITLE-C"`
		D string `cell:"-" title:"DDDDD"`
		E string
	}

	var sheets Sheets

	sheets = append(sheets,
		Sheet{
			Name: "data",
			Rows: []interface{}{
				&Row{A: "a1", B: "b1", C: 1, D: "d1"},
				&Row{A: "a2", B: "b2", C: 2, D: "d2"},
			},
		},
		Sheet{
			Name: "data2",
			Rows: []interface{}{
				&Row{A: "aa11", B: "bb11", C: 1, D: "dd11"},
				&Row{A: "aa22", B: "bb22", C: 2, D: "dd22"},
			},
		})

	err := SaveToXlsx(sheets, "./test/fillcells.xlsx")
	if err != nil {

		t.Error(err)
	}
}
