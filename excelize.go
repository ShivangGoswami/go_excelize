package main

import (
	"flag"
	"fmt"
	"sort"
	"strconv"
	"strings"
	"sync"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func checkblank(str *string) bool {
	return strings.EqualFold(*str, "")
}

func CaseInsensitiveContains(s, substr string) bool {
	s, substr = strings.ToUpper(s), strings.ToUpper(substr)
	return strings.Contains(s, substr)
}

func main() {
	var wg sync.WaitGroup
	StartPtr := flag.String("start", "", "provide the starting column alphabet")
	EndPtr := flag.String("end", "", "provide the ending column alphabet")
	SheetPtr := flag.String("sheet", "", "provide the sheet")
	FilePtr := flag.String("file", "", "provide file name")
	SumPtr := flag.String("sumrow", "", "provide the sum row")
	SortPtr := flag.String("sortrow", "", "provide the sort row")
	flag.Parse()
	if checkblank(FilePtr) || checkblank(SheetPtr) || checkblank(StartPtr) || checkblank(EndPtr) || checkblank(SumPtr) || checkblank(SortPtr) {
		fmt.Println("Invalid Flags")
		return
	}
	//f, err := excelize.OpenFile("./Makerspace Spreadsheet.xlsx")
	f, err := excelize.OpenFile(*FilePtr)
	if err != nil {
		fmt.Println(err)
		return
	}
	Start, err := excelize.ColumnNameToNumber(strings.ToUpper(*StartPtr))
	if err != nil {
		fmt.Println(err)
		return
	}
	End, err := excelize.ColumnNameToNumber(strings.ToUpper(*EndPtr))
	if err != nil {
		fmt.Println(err)
		return
	}
	//rows, err := f.GetRows("gradebook-export (2)")
	rows, err := f.GetRows(*SheetPtr)
	if err != nil {
		fmt.Println(err)
		return
	}
	header := rows[0]
	rows = rows[1:]
	col, err := excelize.ColumnNameToNumber(strings.ToUpper(*SortPtr))
	if err != nil {
		return
	}
	res := sort.SliceIsSorted(rows, func(i, j int) bool {
		a, _ := strconv.Atoi(rows[i][col-1])
		b, _ := strconv.Atoi(rows[j][col-1])
		return a > b
	})
	if res {
		fmt.Println("already executed!!")
		return
	}
	sumer := make(map[string][]int)
	location := make(map[string]int)
	for ind, val := range header[Start-1 : End] {
		temp := strings.Split(val, "-")
		tmp := strings.Trim(temp[1], " ")
		sumer[tmp] = make([]int, 0)
		location[tmp] = ind + Start
	}
	for ind, val := range header[End:] {
		for key := range sumer {
			if CaseInsensitiveContains(val, key) {
				sumer[key] = append(sumer[key], ind+End)
			}
		}
	}
	for ind, val := range rows {
		if len(header) == len(val) {
			wg.Add(1)
			go func(file *excelize.File, i int, sinrow []string) {
				defer wg.Done()
				catcher := make(map[string]int)
				for key, val := range sumer {
					for _, elem := range val {
						if num, err := strconv.Atoi(sinrow[elem]); err == nil {
							catcher[key] += num
						}
					}
					col, err := excelize.ColumnNumberToName(location[key])
					if err != nil {
						fmt.Println(err)
					}
					temp := strconv.Itoa(i)
					err = file.SetCellValue(*SheetPtr, col+temp, catcher[key])
					if err != nil {
						fmt.Println(err)
					}
					sum := 0
					for i := Start; i <= End; i++ {
						str, err := excelize.ColumnNumberToName(i)
						if err != nil {
							continue
						}
						catcher, err := file.GetCellValue(*SheetPtr, str+temp)
						catcher1, _ := strconv.Atoi(catcher)
						sum += catcher1
					}
					err = file.SetCellValue(*SheetPtr, *SumPtr+temp, sum)
					//err = file.SetCellFormula(*SheetPtr, *SumPtr+temp, "=SUM("+*StartPtr+temp+":"+*EndPtr+temp+")")
					if err != nil {
						fmt.Println(err)
					}
				}
			}(f, ind+2, val)
		} else {
			continue
		}
	}
	wg.Wait()
	err = f.Save()
	if err != nil {
		fmt.Println("Unable to Process:", err)
		return
	}
	f, err = excelize.OpenFile(*FilePtr)
	if err != nil {
		fmt.Println(err)
		return
	}
	rows, err = f.GetRows(*SheetPtr)
	if err != nil {
		fmt.Println(err)
		return
	}
	rows = rows[1:]
	sort.Slice(rows, func(i, j int) bool {
		a, _ := strconv.Atoi(rows[i][col-1])
		b, _ := strconv.Atoi(rows[j][col-1])
		return a > b
	})
	for ind, val := range rows {
		wg.Add(1)
		go func(file *excelize.File, i int, sinrow []string) {
			defer wg.Done()
			for x, y := range sinrow {
				if y == "" {
					continue
				}
				col, err := excelize.ColumnNumberToName(x + 1)
				if err != nil {
					fmt.Println(err)
				}
				if a, err := strconv.Atoi(y); err != nil {
					file.SetCellValue(*SheetPtr, col+strconv.Itoa(i), y)
				} else {
					file.SetCellValue(*SheetPtr, col+strconv.Itoa(i), a)
				}
			}
		}(f, ind+2, val)
	}
	wg.Wait()
	err = f.Save()
	if err != nil {
		fmt.Println("Unable to Process:", err)
	}
}
