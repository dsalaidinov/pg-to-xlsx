package main

import (
	"database/sql"
	"flag"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"time"

	_ "github.com/lib/pq"
	"github.com/tealeg/xlsx"
)

var (
	host string
	user string
	pass string
	dbname string
	sqlf string
	outf string
)

func init() {
	flag.StringVar(&host, "h", "", "PostgreSQL host")
	flag.StringVar(&user, "u", "", "User ID")
	flag.StringVar(&pass, "p", "", "Password")
	flag.StringVar(&dbname, "d", "", "Database name")
	flag.StringVar(&sqlf, "s", "", "SQL Query filename")
	flag.StringVar(&outf, "o", "", "Output filename")

	if len(os.Args) < 5 {
		flag.Usage()
		os.Exit(1)
	}

	flag.Parse()
}

func main() {
	dsn := fmt.Sprintf("host=%s user=%s password=%s dbname=%s sslmode=disable", host, user, pass, dbname)
	conn, err := sql.Open("postgres", dsn)
	if err != nil {
		log.Fatalf("error opening connection to server, %s\n", err)
	}
	defer conn.Close()

	err = conn.Ping()
	if err != nil {
		log.Fatalf("error verifying connection to server, %s\n", err)
		return
	}

	query, err := ioutil.ReadFile(sqlf)
	if err != nil {
		log.Fatalf("error opening SQL Query file %s, %s\n", sqlf, err)
	}

	rows, err := conn.Query(string(query))
	if err != nil {
		log.Fatalf("error running query, %s\n", err)
	}
	defer rows.Close()

	err = generateXLSXFromRows(rows, outf)
	if err != nil {
		log.Fatal(err)
	}
}

func generateXLSXFromRows(rows *sql.Rows, outf string) error {
	var err error

	colNames, err := rows.Columns()
	if err != nil {
		return fmt.Errorf("error fetching column names, %s\n", err)
	}
	length := len(colNames)

	pointers := make([]interface{}, length)
	container := make([]interface{}, length)
	for i := range pointers {
		pointers[i] = &container[i]
	}

	xfile := xlsx.NewFile()
	xsheet, err := xfile.AddSheet("Sheet1")
	if err != nil {
		return fmt.Errorf("error adding sheet to xlsx file, %s\n", err)
	}

	xrow := xsheet.AddRow()
	xrow.WriteSlice(&colNames, -1)

	for rows.Next() {
		err = rows.Scan(pointers...)
		if err != nil {
			return fmt.Errorf("error scanning sql row, %s\n", err)
		}

		xrow = xsheet.AddRow()

		for _, v := range container {
			xcell := xrow.AddCell()
			switch v := v.(type) {
			case string:
				xcell.SetString(v)
			case []byte:
				xcell.SetString(string(v))
			case int64:
				xcell.SetInt64(v)
			case float64:
				xcell.SetFloat(v)
			case bool:
				xcell.SetBool(v)
			case time.Time:
				xcell.SetDateTime(v)
			default:
				xcell.SetValue(v)
			}
		}
	}

	err = xfile.Save(outf)
	if err != nil {
		return fmt.Errorf("error writing to output file %s, %s\n", outf, err)
	}

	return nil
}
