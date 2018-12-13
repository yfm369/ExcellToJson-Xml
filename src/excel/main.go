/*!
 * <将excel数据转成json>
 *
 * Copyright (c) 2018 by <yfm/ Fermin Co.>
 */

package main

import (
	"encoding/xml"
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
)

func CheckPath(path string) {
	_, err := os.Stat(path)
	if err == nil {
		return
	}

	if os.IsNotExist(err) {
		err := os.Mkdir(path, os.ModePerm)
		if err != nil {
			fmt.Println("os.Mkdir error :", err.Error())
		}
		return
	}

	fmt.Println("CheckPath error :", err.Error())
}

const (
	EXCEL_TO_JSON = 1
	EXCEL_TO_XML  = 2
	EXCEL_TO_ALL  = 3
)

type FieldInfo struct {
	FieldName string
	FieldType string
}

func ScanlXlsxFiles(ntype int) {
	files, _ := filepath.Glob("*.xlsx")
	if len(files) < 0 {
		return
	}

	if ntype == EXCEL_TO_JSON {
		for _, v := range files {
			ExcelToJson(v)
		}
	} else if ntype == EXCEL_TO_XML {
		for _, v := range files {
			ExcelToXml(v)
		}
	} else { //To Json & XML
		for _, v := range files {
			ExcelToJson(v)
			ExcelToXml(v)
		}
	}
}

func main() {
	ScanlXlsxFiles(EXCEL_TO_ALL)
	//testreadxml()
}

//测试
type Records struct {
	XMLName xml.Name `xml:"RECORDS"`
	Data    []Record `xml:"RECORD"`
}

type Record struct {
	XMLName     xml.Name `xml:"RECORD"`
	Id          int      `xml:"id"`
	Name        string   `xml:"name"`
	Photoid     int      `xml:"photoid"`
	Sceneid     int      `xml:"sceneid"`
	Desc        string   `xml:"desc"`
	Price       int      `xml:"price"`
	Maxfarmer   int      `xml:"maxfarmer"`
	Herbtime    int      `xml:"herbtime"`
	Inspiretype string   `xml:"inspiretype"`
	Inspirecd   int      `xml:"inspirecd"`
	Product     string   `xml:"product"`
	Addproduct  string   `xml:"addproduct"`
	Addlimit    int      `xml:"addlimit"`
	Backrate    float32  `xml:"backrate"`
	Quickprice  int      `xml:"quickprice"`
}

func testreadxml() {
	file, err := os.Open("herb.xml")
	if err != nil {
		fmt.Printf("error: %v", err)
		return
	}
	defer file.Close()

	data, err := ioutil.ReadAll(file)
	if err != nil {
		fmt.Printf("error: %v", err)
		return
	}

	fmt.Printf("data :", string(data))

	v := Records{}
	err = xml.Unmarshal(data, &v)
	if err != nil {
		fmt.Printf("error: %v", err)
		return
	}

	fmt.Println(v)
}
