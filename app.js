const express = require('express');
const app = express();
const bodyParser = require('body-parser');
const multer = require('multer');
const excelToJson = require("convert-excel-to-json");
const mongo = require('mongodb').MongoClient;
const url = "mongodb://localhost:27017";
const assert = require('assert');
const path = require('path');
const hbs = require('express-handlebars');
let tableName='';

app.engine('hbs', hbs({extname: 'hbs', defaultLayout: 'layout', layoutsDir: __dirname + '/views/layouts/'}));
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'hbs');

app.use(bodyParser.json());
const storage = multer.diskStorage({ //multers disk storage settings
    destination: function (req, file, cb) {
        cb(null, './uploads/')
    },
    filename: function (req, file, cb) {
        let datetimestamp = Date.now();
        cb(null, file.fieldname + '-' + datetimestamp + ".xlsx");
    }
});
const saveToDb = (objects,name) => { 
    mongo.connect(url, function(err, client) {
        const db = client.db('mydb');
        assert.equal(null, err);
        db.collection(name).insertMany(objects, function(err, result) {
          assert.equal(null, err);
          console.log('Items inserted');
          client.close();
        });
    });   
}
app.get('/info', function(req, res) {
    if(tableName) {
        var resultArray = [];
        mongo.connect(url, function(err, client) {
        assert.equal(null, err);
        const db = client.db('mydb');
        var cursor = db.collection(tableName).find();
        cursor.forEach(function(doc, err) {
            assert.equal(null, err);
            resultArray.push(doc);
        }, function() {
            client.close();
            res.render('index', {items: resultArray});
        });
        });
    }
    else {
        res.render('error', {message:"Sorry no table for this excel file"} )
    }
});
const upload = multer({ //multer settings
                storage: storage,
                fileFilter : function(req, file, callback) { //file filter
                    let nameParticles = file.originalname.split('.');
                    let extension = nameParticles[nameParticles.length-1];
                    if (extension !== "xlsx") {
                        return callback(new Error('Wrong extension type'));
                    }
                    callback(null, true);
                }
            }).single('file');
/** API path that will upload the files */
app.post('/upload', function(req, res) {
    upload(req,res,function(err){
        if(err){
             res.render("error", { message : err});
             return;
        }
        if(!req.file){
            res.render("error", { message : "No file passed"});
            return;
        }
        try {
            const result = excelToJson({
                sourceFile: req.file.path,
                header:{
                    rows: 6 
                },
                range: 'B6:L81',
                columnToKey: {
                    '*': '{{columnHeader}}'
                },
                sheets: ['Sheet1']
            });

            tableName = req.file.originalname+Date.now();
            res.render('json', {items: result.Sheet1});

            saveToDb(result.Sheet1,tableName);
        } catch (e){
            res.render("error", { message : "Corupted excel file"});
        }
    });
});

app.get('/',function(req,res){
    res.render("form");
});
app.listen('3000', function(){
    console.log('running on 3000...');
});