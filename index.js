const express = require('express')
const bodyParser = require('body-parser')
const multer = require('multer')
const excelToJson = require('convert-excel-to-json')

const app = express()

var storage = multer.diskStorage({ //multers disk storage settings
  destination: function (req, file, cb) {
    cb(null, './uploads/')
  },
  filename: function (req, file, cb) {
    var datetimestamp = Date.now();
    cb(null, file.fieldname + '-' + datetimestamp + '.' + file.originalname.split('.')[file.originalname.split('.').length -1])
  }
});

var upload = multer({ //multer settings
  storage: storage,
  fileFilter : function(req, file, callback) { //file filter
    if (['xls', 'xlsx'].indexOf(file.originalname.split('.')[file.originalname.split('.').length-1]) === -1) {
      return callback(new Error('Wrong extension type'));
    }
    callback(null, true);
  }
}).single('file');


const sheets = '工作表1'

/** API path that will upload the files */
app.post('/upload/better', function(req, res) {
  upload(req,res,function(err){
    if(err){
      res.json({error_code:1,err_desc:err});
      return;
    }
    if(!req.file){
      res.json({error_code:1,err_desc:"没有上传文件"});
      return;
    }
    try {
      const result = excelToJson({
        sourceFile: req.file.path,
        header: {
          rows: 1
        },
        sheets: [sheets],
        columnToKey: {
          A: 'title',
          B: 'description',
          C: 'level'
        }
      })

      res.json(result[sheets]);
    } catch (e){
      res.json({error_code:1,err_desc:"文件损坏"});
    }
  })

});



/** API path that will upload the files */
app.post('/upload/life', function(req, res) {
  upload(req,res,function(err){
    if(err){
      res.json({error_code:1,err_desc:err});
      return;
    }
    if(!req.file){
      res.json({error_code:1,err_desc:"没有上传文件"});
      return;
    }
    try {
      const result = excelToJson({
        sourceFile: req.file.path,
        header: {
          rows: 1
        },
        sheets: [sheets],
        columnToKey: {
          A: 'year',
          B: 'status',
          C: 'title',
          D: 'description',
          E: 'level',
        }
      })

      let rawData = result[sheets]
      let data = {}
      // console.log(rawData)
      rawData.forEach(d => {
        let { year, status, title, description, level } = d
        year = String(year)
        data[year] = data[year] || {}
        data[year].year = year
        data[year].status = status || ''
        data[year].contents = data[year].contents || []
        data[year].contents.push({
          title, description, level
        })
      })

      let re = Object.values(data)

      res.json(re);
    } catch (e){
      res.json({error_code:1,err_desc:"文件损坏"});
    }
  })

});

app.get('/',function(req,res){
  res.sendFile(__dirname + "/index.html");
});

app.listen('3000', function(){
  console.log('running on 3000...');
});
