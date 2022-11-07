package main

import (
	"database/sql"
	"encoding/json"
	"fmt"
	_ "github.com/mattn/go-sqlite3"
	"github.com/xuri/excelize/v2"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"strconv"
)

var DB *sql.DB

func main() {
	log.Println("Init DB")
	db, err := sql.Open("sqlite3", "test.db")
	if err != nil {
		log.Fatalln(err)
	}
	var version string
	err = db.QueryRow("SELECT SQLITE_VERSION()").Scan(&version)

	if err != nil {
		log.Fatal(err)
	}
	createUserTable(db)
	DB = db
	mux := http.NewServeMux()
	fileServer := http.FileServer(http.Dir("./exl"))
	mux.Handle("/file/", http.StripPrefix("/file", fileServer))
	mux.HandleFunc("/save", saveHandler)
	mux.HandleFunc("/", getAllPageHandler)
	port := os.Getenv("PORT")
	if port == "" {
		port = "9000" // Default port if not specified
	}
	log.Println("run localhost:" + port)
	err = http.ListenAndServe(":"+port, mux)
	log.Println(err)
}

func exl(users []UserInfo) error {
	f := excelize.NewFile()
	sheet := "Sheet1"
	for i, user := range users {
		n := strconv.Itoa(i + 1)
		err := f.SetCellValue(sheet, "A"+n, user.FirstName)
		if err != nil {
			return err
		}
		err = f.SetCellValue(sheet, "B"+n, user.LastName)
		if err != nil {
			return err
		}
		err = f.SetCellValue(sheet, "C"+n, user.Email)
		if err != nil {
			return err
		}
		err = f.SetCellValue(sheet, "D"+n, user.Phone)
		if err != nil {
			return err
		}
		err = f.SetCellValue(sheet, "E"+n, user.Instagram)
		if err != nil {
			return err
		}
	}
	if err := f.SaveAs("exl/simple.xlsx"); err != nil {
		return err
	}
	return nil
}

func createUserTable(db *sql.DB) {
	usersTable := `
		CREATE TABLE if not exists users (
        id integer NOT NULL PRIMARY KEY AUTOINCREMENT,
        "firstName" TEXT,
        "lastName" TEXT,
        "email" TEXT,
        "phone" TEXT,
        "instagram" TEXT
        );`
	query, err := db.Prepare(usersTable)
	if err != nil {
		log.Fatal(err)
	}
	_, err = query.Exec()
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("User table created successfully!")
}

func createUser(db *sql.DB, info UserInfo) error {
	query := "INSERT into users (firstName, lastName, email, phone, instagram) values ($1, $2, $3, $4, $5)"
	_, err := db.Exec(query, info.FirstName, info.LastName, info.Email, info.Phone, info.Instagram)
	return err
}
func getAllUsers(db *sql.DB) ([]UserInfo, error) {
	log.Println("getAllUsers")
	query := "select firstName, lastName, email, phone, instagram from users"
	rows, err := db.Query(query)
	if err != nil {
		return nil, err
	}
	var userInfos []UserInfo
	for rows.Next() {
		var tp UserInfo
		if err := rows.Scan(&tp.FirstName, &tp.LastName, &tp.Email, &tp.Phone, &tp.Instagram); err != nil {
			return nil, err
		}
		userInfos = append(userInfos, tp)
	} //end of for loop
	return userInfos, nil
}

type UserInfo struct {
	FirstName string `json:"firstName"`
	LastName  string `json:"lastName"`
	Email     string `json:"email"`
	Phone     string `json:"phone"`
	Instagram string `json:"instagram"`
}

func saveHandler(w http.ResponseWriter, r *http.Request) {
	log.Println("saveS")
	w.Header().Set("Content-Type", "text/html; charset=ascii")
	w.Header().Set("Access-Control-Allow-Origin", "*")
	w.Header().Set("Access-Control-Allow-Headers", "Content-Type,access-control-allow-origin, access-control-allow-headers")
	body, err := ioutil.ReadAll(r.Body) // response body is []byte
	defer r.Body.Close()

	if err != nil {
		w.WriteHeader(http.StatusBadRequest)
		return
	}

	var result UserInfo
	if err := json.Unmarshal(body, &result); err != nil { // Parse []byte to the go struct pointer
		w.WriteHeader(http.StatusBadRequest)
		return
	}
	err = createUser(DB, result)
	if err != nil {
		w.WriteHeader(http.StatusInternalServerError)
		return
	}

}
func getAllPageHandler(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "text/html; charset=ascii")
	w.Header().Set("Access-Control-Allow-Origin", "*")
	w.Header().Set("Access-Control-Allow-Headers", "Content-Type,access-control-allow-origin, access-control-allow-headers")
	users, err := getAllUsers(DB)
	if err != nil {
		w.WriteHeader(http.StatusInternalServerError)
		return
	}
	err = exl(users)
	if err != nil {
		log.Println(err)
		w.WriteHeader(http.StatusInternalServerError)
	}
	w.Write([]byte(`
<!DOCTYPE HTML>
<html>
 <head>
   <meta charset="utf-8">
  <title>Тег А</title>
 </head>
 <body>
  <p><a href="/file">link</a></p>
</body>
</html>
`))
}
