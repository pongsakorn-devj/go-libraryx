package excel

import "fmt"

func CreateFile(data *Rows, path string) error {

	shareName := fmt.Sprintf(`%v.xlsx`, path)
	pathFile := fmt.Sprintf(`%s%s`, TEMP_PATH, shareName)

	// create file
	ex := ReadExcel(data, pathFile)
	if ex != nil {
		return ex
	}

	// save file
	// ex = SaveExcel(data, pathFile)
	// if ex != nil {
	// 	return ex
	// }

	return nil

}
