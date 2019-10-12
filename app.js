const express = require('express');
const app = express();
const bodyParser = require('body-parser');
const multer = require('multer');
const excelToJson = require("convert-excel-to-json");
app.use(bodyParser.json());
const storage = multer.diskStorage({ //multers disk storage settings
    destination: function (req, file, cb) {
        cb(null, './uploads/')
    },
    filename: function (req, file, cb) {
        let datetimestamp = Date.now();
        cb(null, file.fieldname + '-' + datetimestamp + '.' + file.originalname.split('.')[file.originalname.split('.').length -1])
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