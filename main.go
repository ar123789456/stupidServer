package main

import (
	"context"
	"encoding/json"
	"github.com/xuri/excelize/v2"
	"go.mongodb.org/mongo-driver/bson"
	"go.mongodb.org/mongo-driver/mongo"
	"go.mongodb.org/mongo-driver/mongo/options"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"strconv"
	"time"
)

var DB *mongo.Client

const uri = "mongodb+srv://user:user@cluster0.stwzkiv.mongodb.net/?retryWrites=true&w=majority"

func main() {
	log.Println("Init DB")
	serverAPIOptions := options.ServerAPI(options.ServerAPIVersion1)
	clientOptions := options.Client().
		ApplyURI("mongodb+srv://user:user@cluster0.stwzkiv.mongodb.net/?retryWrites=true&w=majority").
		SetServerAPIOptions(serverAPIOptions)
	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	client, err := mongo.Connect(ctx, clientOptions)
	if err != nil {
		log.Fatal(err)
	}
	DB = client
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
func createUser(info UserInfo) error {
	_, err := DB.Database("main").Collection("users").InsertOne(context.TODO(), &info)
	return err
}
func getAllUsers() ([]UserInfo, error) {
	log.Println("getAllUsers")
	var userInfos []UserInfo
	cur, err := DB.Database("main").Collection("users").Find(context.TODO(), bson.D{})
	if err != nil {
		return nil, err
	}

	for cur.Next(context.TODO()) {
		var userinfo UserInfo
		err = cur.Decode(&userinfo)
		if err != nil {
			return nil, err
		}
		userInfos = append(userInfos, userinfo)
	}
	return userInfos, nil
}

type UserInfo struct {
	FirstName string `json:"firstName"`
	LastName  string `json:"lastName"`
	Email     string `json:"email"`
	Phone     string `json:"phone"`
	Instagram string `json:"instagram"`
}

func addCorsHeader(res http.ResponseWriter) {
	headers := res.Header()
	headers.Add("Access-Control-Allow-Origin", "*")
	headers.Add("Vary", "Origin")
	headers.Add("Vary", "Access-Control-Request-Method")
	headers.Add("Vary", "Access-Control-Request-Headers")
	headers.Add("Access-Control-Allow-Headers", "Content-Type, Origin, Accept, token")
	headers.Add("Access-Control-Allow-Methods", "GET, POST,OPTIONS")
}

func saveHandler(w http.ResponseWriter, r *http.Request) {
	log.Println("saveS")
	addCorsHeader(w)
	if r.Method == http.MethodOptions {
		return
	}
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
	err = createUser(result)
	if err != nil {
		w.WriteHeader(http.StatusInternalServerError)
		return
	}

}
func getAllPageHandler(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "text/html; charset=ascii")
	w.Header().Set("Access-Control-Allow-Origin", "*")
	w.Header().Set("Access-Control-Allow-Headers", "Content-Type,access-control-allow-origin, access-control-allow-headers")
	users, err := getAllUsers()
	if err != nil {
		log.Println(err)
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
<style>
	a {
font-size: 20px;
}
</style>
 </head>
 <body>
  <p><a style="font-size: 50px;" href="/file">Download</a></p>
</body>
</html>
`))
}
