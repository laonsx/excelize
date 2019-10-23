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

	for _, sheet := range sheets {

		_ = f.NewSheet(sheet.Name)

		skip := 1

		// 自动填充标题
		if len(sheet.Rows) > 0 {

			s := structs.New(sheet.Rows[0])
			for _, field := range s.Fields() {

				if skipCellField(field) {

					continue
				}

				err = f.SetCellValue(sheet.Name, fmt.Sprintf("%s%d", field.Tag("cell"), skip), getCellFieldTitle(field))
				if err != nil {

					return err
				}
			}

			skip++
		}

		// 填充数据
		for i, row := range sheet.Rows {

			s := structs.New(row)
			for _, field := range s.Fields() {

				if skipCellField(field) {

					continue
				}

				err = f.SetCellValue(sheet.Name, fmt.Sprintf("%s%d", field.Tag("cell"), i+skip), field.Value())
				if err != nil {

					return err
				}
			}
		}
	}

	return err
}

func SaveToXlsx(sheets Sheets, path string) error {

	file := NewFile()

	file.SetSheetName("Sheet1", "data")

	err := file.FillSheetCells(sheets)
	if err != nil {

		return err
	}

	err = file.SaveAs(path)
	if err != nil {

		return err
	}

	return nil
}

func skipCellField(field *structs.Field) bool {

	cellTag := field.Tag("cell")
	if cellTag == "-" || cellTag == "" {

		return true
	}

	return false
}

func getCellFieldTitle(field *structs.Field) string {

	title := field.Tag("title")
	if title == "" {

		title = field.Name()
	}

	return title
}
