const express = require('express');
const app = express();
const bodyParser = require('body-parser');
const multer = require('multer');
const excelToJson = require("convert-excel-to-json");
const mongo = require('mongodb').MongoClient;
const url = "mongodb://localhost:27017";
const assert = require('assert');

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
    // let objects = JSON.parse(json);
    // let objects = json; 
    mongo.connect(url, function(err, client) {
        const db = client.db('mydb');
        assert.equal(null, err);
        // db.collection(name).insertMany(objects, function(err, result) {
        db.collection("user-data").insertMany(objects, function(err, result) {
          assert.equal(null, err);
          console.log('Item inserted');
          client.close();
        });
    });   
}
app.get('/get', function(req, res) {
    var resultArray = [];
    mongo.connect(url, function(err, client) {
      assert.equal(null, err);
      const db = client.db('mydb');
      var cursor = db.collection('user-data').find();
      cursor.forEach(function(doc, err) {
        assert.equal(null, err);
        resultArray.push(doc);
      }, function() {
        client.close();
        res.json({data:resultArray});
      });
    });
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
        console.log(req);
        if(err){
             res.json({error_code:1,err_desc:err});
             return;
        }
        if(!req.file){
            res.json({error_code:1,err_desc:"No file passed"});
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
                }
            });
            res.json({error_code:0,err_desc:null, data: result});
            saveToDb(result.Sheet1,req.file.originalname);
            console.log(result);
        } catch (e){
            res.json({error_code:1,err_desc:"Corupted excel file"});
        }
    });
});

app.get('/',function(req,res){
    res.sendFile(__dirname + "/index.html");
});
app.listen('3000', function(){
    console.log('running on 3000...');
});