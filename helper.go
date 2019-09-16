package excelize

import (
	"fmt"

	"github.com/laonsx/structs"
)

type (
	Sheets []Sheet
	Sheet  struct {
		Name string        `json:"name"`
		Rows []interface{} `json:"rows"`
	}
)

func (f *File) FillSheetCells(sheets Sheets) (err error) {

	delSheet1 := true

	for _, sheet := range sheets {

		_ = f.NewSheet(sheet.Name)

		for i, row := range sheet.Rows {

			s := structs.New(row)
			for _, field := range s.Fields() {

				if field.Tag("cell") == "-" {

					continue
				}

				err = f.SetCellValue(sheet.Name, fmt.Sprintf("%s%d", field.Tag("cell"), i+1), field.Value())
				if err != nil {

					return err
				}
			}
		}

		if sheet.Name == "Sheet1" {

			delSheet1 = false
		}
	}

	if delSheet1 {

		f.DeleteSheet("Sheet1")
	}

	return err
}
