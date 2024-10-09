package excel

import (
	"os"
	"path/filepath"
	"strings"

	"github.com/mzimmerman/xlsxwriter"
)

type Rows struct {
	TableName   string       `json:"tbl,omitempty"` // TableEmpty
	ColumnTypes []ColumnType `json:"colt,omitempty"`
	Columns     []string     `json:"cols,omitempty"`
	DriverName  string       `json:"drv,omitempty"`
	Rows        []Map        `json:"rows,omitempty"`
}
type Map map[string]any

type ColumnType struct {
	Name             string `json:"nam"`
	DatabaseTypeName string `json:"dtn"`
}

const TEMP_PATH string = `./tmp/`

func ReadExcel(rows *Rows, path string) error {

	// tmp +  ชื่อไฟล์เท่านั้น
	// path := filex.TEMP_PATH + fileName
	if !strings.Contains(path, `/`) {
		path = TEMP_PATH + path
	}

	// prepare data
	dta := [][]string{}
	for k, v := range rows.Rows {
		if k == 0 {
			// header
			dta = append(dta, rows.Columns)
		}
		// detail
		dtr := []string{}
		for _, c := range rows.Columns {
			dtr = append(dtr, v[c].(string))
		}
		dta = append(dta, dtr)
	}

	// path = path/to/file.xlsx
	fo, err := os.Create(filepath.Clean(path))
	if err != nil {
		return err
	}
	defer func() {
		_ = fo.Close()
	}()
	xw, err := xlsxwriter.New(fo)
	if err != nil {
		return err
	}
	defer xw.Close()

	for _, v := range dta {
		err = xw.WriteLine(v)
		if err != nil {
			return err
		}
	}
	if err != nil {
		return err
	}

	return nil

}
