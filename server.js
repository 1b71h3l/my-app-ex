const express = require('express')
const app = express()
const bodyparser = require('body-parser')
const fs = require('fs');
const readXlsxFile = require('read-excel-file/node');
const mysql = require('mysql')
const multer = require('multer')
const path = require('path')
require('dotenv').config();


//use express static folder
app.use(express.static("./public"))


// body-parser middleware use
app.use(bodyparser.json())
app.use(bodyparser.urlencoded({ extended: true }))


// Database connection
const db = mysql.createPool({
    connectionLimit : 100,
    port            : process.env.DB_port,
    host            : process.env.DB_HOST,
    user            : process.env.DB_USER,
    password        : process.env.DB_PASS,
    database        : process.env.DB_NAME
});


db.getConnection( (err,connection) => {
    if (err) {
        return console.error('error: ' + err.message);
    }
    console.log('Connected to the MySQL server as ID :' + connection.threadId);
})


// Multer Upload Storage
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, __dirname + "/uploads/")  //__basedir +
    },
    filename: (req, file, cb) => {
        cb(null, file.fieldname + "-" + Date.now() + "-" + file.originalname)
    }
});
const upload = multer({ storage: storage });



//! Routes start
//route for Home page
app.get('/', (req, res) => {
    res.sendFile( __dirname + "/index.html");  //__dirname +
});


// -> Express Upload RestAPIs
app.post('/uploadfile', upload.single("uploadfile"), (req, res) => {
    importExcelData2MySQL( __dirname + "/uploads/" + req.file.filename);  //__basedir +
    console.log(res);
    res.end();
});


// -> Import Excel Data to MySQL database
function importExcelData2MySQL(filePath) {
    // File path.
    readXlsxFile(filePath).then((rows) => {
        // `rows` is an array of rows
        // each row being an array of cells.     
        console.log(rows);
        /**
        [ [ 'Id', 'Name', 'Address', 'Age' ],
        [ 1, 'john Smith', 'London', 25 ],
        [ 2, 'Ahman Johnson', 'New York', 26 ]
        */
        // Remove Header ROW
        rows.shift();
        // Open the MySQL connection
        db.getConnection( (error,connection) => {
            if (error) {
                console.error(error);
            } else {
                let query = 'INSERT INTO customer (id, address, name, age) VALUES ?';
                connection.query(query, [rows], (error, response) => {
                    connection.release();
                    console.log(error || response);
                    /**
                    OkPacket {
                    fieldCount: 0,
                    affectedRows: 5,
                    insertId: 0,
                    serverStatus: 2,
                    warningCount: 0,
                    message: '&Records: 5  Duplicates: 0  Warnings: 0',
                    protocol41: true,
                    changedRows: 0 } 
                    */
                    
                    console.log('finished');
                });
            }
        });
    })
}

// Create a Server
let server = app.listen(8080, function () {
    let host = server.address().address
    let port = server.address().port
    console.log("App listening at http://%s:%s", host, port)
})