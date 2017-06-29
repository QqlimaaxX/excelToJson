var express = require('express');
var fileUpload = require('express-fileupload');
var Excel = require('exceljs');
var filename="file.xlsx";
var port = process.env.port || 3000;

var app = express();
app.use(fileUpload());
app.set('view engine','ejs')

app.get("/",function(req,res){
	res.render("index");
});

app.post('/upload', function(req, res) {
	var sampleFile;
	sampleFile = req.files.sampleFile;
	sampleFile.mv('./file.xlsx', function(err) {
		if (err) {
			console.log("Error");
			res.status(500).send(err);
		}
		else {
			createObjfromxlsx();
			setTimeout(function(){
				res.redirect('/json');
			},1000);
		}
	});
});

app.get("/json",function(req,res){
	res.json({data:json});
});

app.listen(port,function(){
	console.log("Server is Live at - "+port);
});

function createObjfromxlsx(){
  var workbook = new Excel.Workbook();
  var obj=[];
  workbook.xlsx.readFile(filename)
      .then(function() {
        //for each sheet
          workbook.eachSheet(function(worksheet,sheetId){
              var title = worksheet.getCell('D1').value;
              
              //get the cell with first name at B# and iterate until its blank
              var i=1;
              while(worksheet.getCell('B'+i).value!="Name"){
                i++;
              }
              //remove all unnecessary stuff
              worksheet.spliceRows(0,i);
              var users = [];
              worksheet.eachRow(function(row,rowNumber){
                var name  = row.getCell(2).value;
                var number  = row.getCell(4).value;
                var dob   = row.getCell(6).value;
                if(number == "-"){
                  number = null;
                }
                if(dob == "-"){
                  dob = null;
                }
                users.push({
                  'name' : name,
                  'number': number,
                  'dob': dob
                });
              });
              obj.push({
                'title':title,
                'users':users
              });
          });
          // console.log(obj);
         json= obj;
      });
}