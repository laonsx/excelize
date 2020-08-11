package excelize

import (
	"encoding/json"
	"errors"
	"fmt"
	"reflect"
	"strconv"
	"time"

	"github.com/laonsx/structs"
)

const (
	FillSheet                 = -1
	PrepareSheet              = 0
	TimeFormatTag             = "time_format" // tag:时间格式化
	CellTag                   = "cell"        // tag:列信息
	TitleTag                  = "title"       // tag:列标题
	DefaultTimeFormatTemplate = "2006-01-02 15:04:05"
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

			// todo 合并单元格情况

			err = f.SetCellValue(sheet.SheetName(), fmt.Sprintf("%s%d", field.Tag(CellTag), 1), getCellFieldTitle(field))
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

			fieldValue := field.Value()
			if v, ok := cellTimeField(field); ok {

				fieldValue = v
			}

			// todo 填充子struct

			err = f.SetCellValue(sheet.SheetName(), fmt.Sprintf("%s%d", field.Tag(CellTag), nextRowIndex), fieldValue)
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

	cellTag := field.Tag(CellTag)
	if cellTag == "-" || cellTag == "" {

		return true
	}

	return false
}

func getCellFieldTitle(field *structs.Field) string {

	title := field.Tag(TitleTag)
	if title == "" {

		title = field.Name()
	}

	return title
}

func cellTimeField(field *structs.Field) (interface{}, bool) {

	t, ok := field.Value().(time.Time)
	if !ok {

		return nil, false
	}

	timeFormat := field.Tag(TimeFormatTag)
	if timeFormat == "" || timeFormat == "-" {

		return t.Format(DefaultTimeFormatTemplate), true
	}

	return t.Format(timeFormat), true
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

var ErrNilPtr = errors.New("ErrNilPtr")
var ErrNotStructPointer = errors.New("ErrNotStructPointer")

func (rows *Rows) ReadStruct(ptr interface{}) error {

	if ptr == nil {

		return ErrNilPtr
	}

	v := reflect.ValueOf(ptr)
	if v.Kind() != reflect.Ptr {

		return ErrNotStructPointer
	}

	v = v.Elem()
	if v.Kind() != reflect.Struct {

		return ErrNotStructPointer
	}

	curRow := rows.rows[rows.curRow]
	rows.curRow++

	columns := make(map[string]string, len(curRow.C))
	d := rows.f.sharedStringsReader()

	for _, colCell := range curRow.C {

		cell, _, err := SplitCellName(colCell.R)
		if err != nil {

			return err
		}

		val, _ := colCell.getValueFrom(rows.f, d)
		columns[cell] = val
	}

	n := v.NumField()
	for i := 0; i < n; i++ {

		field := v.Type().Field(i)
		cell := field.Tag.Get(CellTag)
		if cell == "-" || cell == "" {

			continue
		}

		value, ok := columns[cell]
		if !ok {

			continue
		}

		fieldV := v.Field(i)
		if !fieldV.CanSet() {

			continue
		}

		switch field.Type.Kind() {
		case reflect.String:

			fieldV.SetString(value)
		case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:

			vInt, err := strconv.ParseInt(value, 10, 64)
			if err != nil {

				return err
			}

			fieldV.SetInt(vInt)
		case reflect.Float64:

			vFloat, err := strconv.ParseFloat(value, 64)
			if err != nil {

				return err
			}

			fieldV.SetFloat(vFloat)
		case reflect.Bool:

			if value == "" || value == "0" {

				fieldV.SetBool(false)
			} else {

				fieldV.SetBool(true)
			}
		case reflect.Ptr, reflect.Struct:

			if !fieldV.CanSet() {

				continue
			}

			var structPtr interface{}
			if field.Type.Kind() == reflect.Struct {

				structPtr = fieldV.Addr().Interface()
			} else {

				structPtr = fieldV.Interface()
			}

			_, isTime := structPtr.(*time.Time)
			if isTime {

				timeFormat := field.Tag.Get(TimeFormatTag)

				var vtime time.Time
				var err error

				if timeFormat != "" && timeFormat != "-" {

					vtime, err = time.Parse(timeFormat, value)
				} else {

					vtime, err = time.Parse(DefaultTimeFormatTemplate, value)
				}

				if err != nil {

					continue
				}

				if field.Type.Kind() == reflect.Ptr {

					fieldV.Set(reflect.ValueOf(&vtime))
				} else {

					fieldV.Set(reflect.ValueOf(vtime))
				}

			} else {

				// todo 继续读取子cell
			}
		}
	}

	value := v.Interface()
	ptr = &value

	return nil
}
