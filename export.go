package main

import (
	"database/sql"
	"fmt"
	_ "github.com/go-sql-driver/mysql"
	"github.com/tealeg/xlsx"
	"flag"
	"os"
	"log"
	"bufio"
	"strings"
	"reflect"
)

var db *sql.DB
var file *xlsx.File
var sheet *xlsx.Sheet
var row *xlsx.Row
var cell *xlsx.Cell
var err error

type Conf struct {
	DB_DRIVER        string
	DB_NAME          string
	DB_USER          string
	DB_PASSWORD      string
	DESTINATION_FILE string
}

var config = new(Conf)
//main function
func main() {
	getConfig()

	var brand int
	//Get brand id from command line
	flag.IntVar(&brand, "brand", 0, "Get brand from command line default 0")
	flag.Parse()

	//connect to database
	db, err := sql.Open(config.DB_DRIVER, config.DB_USER+":"+config.DB_PASSWORD+"@/"+config.DB_NAME)
	if err != nil {
		panic(err.Error())
	}

	defer db.Close()

	//create new xlsx file
	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Sheet1")
	setHeader()

	sth, err := db.Prepare("SELECT `CategoryID`, `Article`, `Title`,`Brand` , `Price`, `Active`, `Description` FROM Catalog_Items where `Brand`=?")
	if err != nil {
		//panic(err.Error())
	}
	rows, err := sth.Query(brand)
	if err != nil {
		//panic(err.Error())
	}
	//fill xlsx file need data
	for rows.Next() {
		var category int
		var article string
		var title string
		var brand int
		var price string
		var active string
		var description string
		err = rows.Scan(&category, &article, &title, &brand, &price, &active, &description)
		if err != nil {
			//panic(err.Error())
		}

		row = sheet.AddRow()
		cell = row.AddCell()
		cell.Value = getCategory(db, category)
		cell = row.AddCell()
		cell.Value = article
		cell = row.AddCell()
		cell.Value = title
		cell = row.AddCell()
		cell.Value = getBrand(db, brand)
		cell = row.AddCell()
		cell.Value = price
		cell = row.AddCell()
		cell.Value = active
		cell = row.AddCell()
		cell.Value = description

	}
	if err != nil {
		fmt.Printf(err.Error())
	}
	//save xlsx file
	err = file.Save("excel/export.xlsx")
	if err != nil {
		fmt.Printf(err.Error())
	}
	os.Exit(0)
}

//getCategory return category title by id
func getCategory(db *sql.DB, category int) string {
	var title string

	rows, err := db.Prepare("select `Title` from `Catalog_Categories` where `Id`=?")
	if err != nil {
		//panic(err.Error())
	}
	defer rows.Close()

	err = rows.QueryRow(category).Scan(&title)
	if err != nil {
		////panic(err.Error())
	}
	return title
}

//getBrand return brand title by id
func getBrand(db *sql.DB, brand int) string {
	var title string
	rows, err := db.Prepare("select `Title` from `Brands_Items` where `Id`=? ")
	if err != nil {
		//panic(err.Error())
	}
	defer rows.Close()
	err = rows.QueryRow(brand).Scan(&title)
	if err != nil {
		//panic(err.Error())
	}

	return title
}

//setHeader add header rows with name rows
func setHeader() {
	names := [7]string{
		"Категория (Разделить ;)",
		"Артикул",
		"Наименование",
		"Производитель",
		"Цена",
		"Активный (0 - Нет, 1 - Да)",
		"Описание",
	}
	row = sheet.AddRow()
	for i := 0; i < len(names); i++ {
		cell = row.AddCell()
		cell.Value = names[i]
	}
}

// getConfig get all setting from .env file and set to need variables
func getConfig() {
	_, err := existsFile(".env")
	if err != nil {
		log.Fatal(err)
	}
	file, err := os.Open(".env")
	if err != nil {
		log.Fatal(err)
	}
	defer file.Close()

	//use reflection for assing struct values
	s := reflect.ValueOf(config).Elem()
	typeOfT := s.Type()

	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		param := strings.Split(scanner.Text(), "=")
		//assing struct values from .env file
		for i := 0; i < s.NumField(); i++ {
			if typeOfT.Field(i).Name == param[0] {
				s.Field(i).SetString(param[1])
			}
		}
	}
	if err := scanner.Err(); err != nil {
		log.Fatal(err)
	}
}

//check exists file
func existsFile(path string) (bool, error) {
	_, err := os.Stat(path)
	if err == nil {
		return true, nil
	}
	if os.IsNotExist(err) {
		return false, nil
	}
	return true, err
}
