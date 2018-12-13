/*!
 * <将excel数据转成json>
 *
 * Copyright (c) 2018 by <yfm/ Fermin Co.>
 */

package main

import (
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/tealeg/xlsx"
)

const (
	OUT_DIR_JSON = "./json/"
)

func init() {
	CheckPath(OUT_DIR_JSON)
}

func PrintJsonFile(filename string, v ...interface{}) {
	file, err := os.OpenFile(OUT_DIR_JSON+filename+".json", os.O_CREATE|os.O_RDWR, 0660)
	if err != nil {
		log.Println("FileLog Error :", filename, " Error:", err.Error())
		return
	}

	file.Write([]byte(fmt.Sprint(v...)))
	file.Close()
}

func ExcelToJson(filename string) {
	xlFile, err := xlsx.OpenFile(filename)
	if err != nil {
		fmt.Printf("xlsx open error :", err.Error())
		return
	}

	strsplite := strings.Split(filename, ".")
	if len(strsplite) <= 0 {
		fmt.Printf("filename error :\n", filename)
		return
	}

	for _, sheet := range xlFile.Sheets {
		if len(sheet.Cols) <= 0 {
			continue
		}

		strjson := "["
		fields := make([]FieldInfo, len(sheet.Cols))
		for i, row := range sheet.Rows {
			if i > 2 {
				strjson += "{"
			}
			for index, cell := range row.Cells {
				if i == 0 {
					fields[index].FieldName = cell.String()
				} else if i == 1 {
					fields[index].FieldType = cell.String()
				} else if i > 2 {
					if fields[index].FieldType == "string" {
						strjson += fmt.Sprintf("\"%s\":\"%s\"", fields[index].FieldName, cell.String())
					} else {
						strjson += fmt.Sprintf("\"%s\":%s", fields[index].FieldName, cell.String())
					}
					if index+1 < len(row.Cells) {
						strjson += ","
					}
				}
			}
			if i > 2 {
				if i+1 < len(sheet.Rows) {
					strjson += "},\n"
				} else {
					strjson += "}"
				}
			}
		}
		strjson += "]"

		PrintJsonFile(strsplite[0], strjson)
	}
}
