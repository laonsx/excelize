package excelize

import (
	"encoding/json"
	"fmt"
	"strconv"

	"github.com/laonsx/structs"
)

const (
	FillSheet    = -1
	PrepareSheet = 0
)

type (
	Sheets []*Sheet
	Sheet  struct {
		Name         string        `json:"name"`
		Rows         []interface{} `json:"rows"`
		NextSheet    int           // 0: 准备 -1:已完成
		SheetMaxRows int
		InitTitle    bool
		sheetName    string
	}
)

func (s *Sheet) SheetName() string {

	if s.sheetName != "" {

		return s.sheetName
	}

	if s.NextSheet == PrepareSheet {

		s.sheetName = s.Name
	} else {

		s.sheetName = s.Name + strconv.Itoa(s.NextSheet)
	}

	return s.sheetName
}

func (f *File) FillSheetCells(sheet *Sheet) (err error) {

	_ = f.NewSheet(sheet.SheetName())

	// 自动填充标题
	if sheet.InitTitle && len(sheet.Rows) > 0 {

		s := structs.New(sheet.Rows[0])
		for _, field := range s.Fields() {

			if skipCellField(field) {

				continue
			}

			err = f.SetCellValue(sheet.SheetName(), fmt.Sprintf("%s%d", field.Tag("cell"), 1), getCellFieldTitle(field))
			if err != nil {

				return err
			}
		}
	}

	nextRowIndex, err := f.nextAppendRowIndex(sheet.SheetName())
	if err != nil {

		return err
	}

	preRowIndex := nextRowIndex - 1

	// 填充数据
	for i := sheet.SheetMaxRows * sheet.NextSheet; i < len(sheet.Rows); i++ {

		row := sheet.Rows[i]
		s := structs.New(row)
		for _, field := range s.Fields() {

			if skipCellField(field) {

				continue
			}

			err = f.SetCellValue(sheet.SheetName(), fmt.Sprintf("%s%d", field.Tag("cell"), nextRowIndex), field.Value())
			if err != nil {

				return err
			}
		}

		nextRowIndex++

		if nextRowIndex-preRowIndex > sheet.SheetMaxRows {

			sheet.NextSheet++
			sheet.sheetName = ""

			return nil
		}
	}

	sheet.NextSheet = FillSheet

	return err
}

func AppendToSheet(sheets Sheets, path string) (err error) {

	file, err := OpenFile(path)
	if err != nil {

		return err
	}

	for _, sheet := range sheets {

		if sheet.NextSheet == FillSheet {

			continue
		}

		err = file.FillSheetCells(sheet)
		if err != nil {

			return err
		}

		if sheet.NextSheet != FillSheet {

			err = file.SaveAs(path)
			if err != nil {

				return err
			}

			return AppendToSheet(sheets, path)
		}
	}

	err = file.SaveAs(path)
	if err != nil {

		return err
	}

	return nil
}

func SaveToXlsx(sheets Sheets, path string) (err error) {

	file := NewFile()

	for i, sheet := range sheets {

		if sheet.NextSheet == FillSheet {

			continue
		}

		if i == 0 {

			file.SetSheetName("Sheet1", sheet.Name)
		}

		err = file.FillSheetCells(sheet)
		if err != nil {

			return err
		}

		if sheet.NextSheet != FillSheet {

			err = file.SaveAs(path)
			if err != nil {

				return err
			}

			return AppendToSheet(sheets, path)
		}
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

func PrintInfo(path string) {

	file, err := OpenFile(path)
	if err != nil {

		return
	}

	datas, err := json.Marshal(file)
	if err != nil {

		return
	}

	fmt.Println(string(datas))
}
