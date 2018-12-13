/*!
 * <将excel数据转成xml>
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
	OUT_DIR = "./xml/"
)

func init() {
	CheckPath(OUT_DIR)
}

func PrintXmlFile(filename string, v ...interface{}) {
	file, err := os.OpenFile(OUT_DIR+filename+".xml", os.O_CREATE|os.O_RDWR, 0660)
	if err != nil {
		log.Println("FileLog Error :", filename, " Error:", err.Error())
		return
	}

	file.Write([]byte(fmt.Sprint(v...)))
	file.Close()
}

func ExcelToXml(filename string) {
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

		strxml := "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n"
		strxml += "<RECORDS>\n"
		fields := make([]FieldInfo, len(sheet.Cols))
		for i, row := range sheet.Rows {
			if i > 2 {
				strxml += "\t<RECORD>\n"
			}

			for index, cell := range row.Cells {
				if i == 0 {
					fields[index].FieldName = cell.String()
				} else if i == 1 {
					fields[index].FieldType = cell.String()
				} else if i > 2 {
					strxml += fmt.Sprintf("\t\t<%s>%s</%s>\n", fields[index].FieldName,
						cell.String(), fields[index].FieldName)
				}
			}

			if i > 2 {
				strxml += "\t</RECORD>\n"
			}
		}
		strxml += "</RECORDS>\n"
		PrintXmlFile(strsplite[0], strxml)
	}
}
