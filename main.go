package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"time"
)

const (
	cellStart = 12
)

func main() {
	file, err := excelize.OpenFile("file.xlsx")

	var name string
	var year int
	var month time.Month

	fmt.Println("What is your name ?")
	_, err = fmt.Scanln(&name)

	fmt.Println("Select the year (ex: 2021)")
	_, err = fmt.Scanln(&year)

	fmt.Println("Select the month (ex: 1 for january, 12 for december)")
	_, err = fmt.Scanln(&month)

	if err != nil {
		fmt.Println(err)
	}

	// get the number of days of the current month
	t := time.Date(year, month+1, 0, 0, 0, 0, 0, time.Local)

	//> styles
	dateStyle, err := file.NewStyle(`{"fill":{"type":"pattern","color":["#E0EBF5"],"pattern":1}, "font":{"bold": true}, "border":[{"type":"left","color":"000000","style":1},{"type":"right","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1}]}`)
	if err != nil {
		fmt.Println(err)
	}

	hourStyle, err := file.NewStyle(`{"alignment":{"horizontal":"center"}, "border":[{"type":"left","color":"000000","style":1},{"type":"right","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1}]}`)
	if err != nil {
		fmt.Println(err)
	}

	WeekendStyle, err := file.NewStyle(`{"fill":{"type":"pattern","color":["#808080"],"pattern":1}, "font":{"bold": true}, "border":[{"type":"left","color":"000000","style":1},{"type":"right","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1}]}`)
	if err != nil {
		fmt.Println(err)
	}

	nameStyle, err := file.NewStyle(`{"alignment":{"horizontal":"center"}, "font":{"bold": true}}`)
	if err != nil {
		fmt.Println(err)
	}

	TotalStyle, err := file.NewStyle(`{"alignment":{"horizontal":"right"}, "fill":{"type":"pattern","color":["#cccccc"],"pattern":1}, "font":{"bold": true}, "border":[{"type":"left","color":"000000","style":1},{"type":"right","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1}]}`)
	if err != nil {
		fmt.Println(err)
	}

	TotalValueStyle, err := file.NewStyle(`{"alignment":{"horizontal":"center"}, "fill":{"type":"pattern","color":["#cccccc"],"pattern":1}, "font":{"bold": true}, "border":[{"type":"left","color":"000000","style":1},{"type":"right","color":"000000","style":1},{"type":"top","color":"000000","style":1},{"type":"bottom","color":"000000","style":1}]}`)
	if err != nil {
		fmt.Println(err)
	}
	//< style

	// Set value of a cell.
	file.SetCellValue("Feuil1", "A4", fmt.Sprintf("%s %d", month, year))
	file.SetCellValue("Feuil1", "B8", "Gaetan Rouseyrol")
	err = file.SetCellStyle("Feuil1", "B8", "B8", nameStyle)

	i := 0
	for day := 1; day <= t.Day(); day++ {
		cellValue := i + cellStart
		date := time.Date(year, month, day, 0, 0, 0, 0, time.Local)

		err = file.SetCellValue("Feuil1", fmt.Sprintf("A%d", cellValue), fmt.Sprintf("%s, %s %d, %d", date.Weekday(), date.Month(), date.Day(), date.Year()))
		err = file.SetCellStyle("Feuil1", fmt.Sprintf("A%d", cellValue), fmt.Sprintf("A%d", cellValue), dateStyle)
		err = file.SetCellStyle("Feuil1", fmt.Sprintf("B%d", cellValue), fmt.Sprintf("B%d", cellValue), hourStyle)

		if int(date.Weekday()) == 0 || int(date.Weekday()) == 6 {
			err = file.SetCellStyle("Feuil1", fmt.Sprintf("A%d", cellValue), fmt.Sprintf("B%d", cellValue), WeekendStyle)
		}

		i++
	}

	err = file.SetCellValue("Feuil1", fmt.Sprintf("A%d", cellStart+i), "Total")
	err = file.SetCellStyle("Feuil1", fmt.Sprintf("A%d", cellStart+i), fmt.Sprintf("A%d", cellStart+i), TotalStyle)
	err = file.SetCellFormula("Feuil1", fmt.Sprintf("B%d", cellStart+i), fmt.Sprintf("=SUM(B%d:B%d)", cellStart, cellStart+i-1))
	err = file.SetCellStyle("Feuil1", fmt.Sprintf("B%d", cellStart+i), fmt.Sprintf("A%d", cellStart+i), TotalValueStyle)

	// Save spreadsheet by the given path.
	if err := file.SaveAs(fmt.Sprintf("%s-%d.xlsx", month, year)); err != nil {
		fmt.Println(err)
	}
}
