// calling express packages
var express = require('express');
var app = express(); 
var lodash = require('lodash');
var pkg =require("xlsx") ;

// function to read the  data from excel sheet
let masterData =[];
let readdataFromMaster = ()=>{
    const workbookmaster = pkg.readFile("master-user-list.xlsx");
     masterDatasheet= workbookmaster.SheetNames;
    for(var i=0;i<masterDatasheet.length;i++)
    {
        const Data = pkg.utils.sheet_to_json(workbookmaster.Sheets[masterDatasheet[i]]).map(val=>val.PersonNumber);
        masterData.push(...Data)
    }
}
readdataFromMaster();
// console.log(masterData)
let sheetNameList,xlData =[],workBook;
let readExcel =()=>
{
    
     workBook = pkg.readFile("edata.xlsx");
    sheetNameList = workBook.SheetNames;
    for(var i=0;i<sheetNameList.length;i++)
    {
        xlData[i] = pkg.utils.sheet_to_json(workBook.Sheets[sheetNameList[i]]).map(val=>val.PersonNumber);
        // console.log(xlData)
    }
}
 readExcel();

 //function to compare data
 var notFilledList =[];
 var comparelists = ()=>{
            for(var i=0;i<sheetNameList.length;i++)
            {
                notFilledList[i] = lodash.difference(masterData,xlData[i]);
            }
        }
comparelists();
// function to create table   


 let r1=[]; 
 let r2='';
 let createTable =()=>{
    for (let j = 0; j < notFilledList[i].length; j++)
    {    
       const userName = notFilledList[i][j];
       r2 += `<tr>
        <td>${j+1}</td>
        <td>${userName}</td>
        </tr>`
    }        
    r1[i] = 
        `<h2>Emplyes who are not submmited on ${sheetNameList[i]}</h2>
         <table style="width:50%">
         <tr>
        <th>No.</th>
        <th>Oracle ID</th>
        <th>Email</th>
        </tr>
        ${r2}
        </table>` 
 }
      for(var i=1;i<notFilledList.length;i++)
      {
        if(notFilledList[i].length != 0){
            createTable();
            r2 =''
        }
        else{
            r1 +=`<h2>Every one submitted their timesheet on ${sheetNameList[i]}</h2>` 
        }    
    }

//  loading expess function to listen on port 8081
   app.get('/', function (req, res) {
     res.send(`<html>
<head>
<style>
table, th, td {
  border: 1px solid black;
  border-collapse: collapse;
}
</style>
<title>employee Detials</title>
</head>
<body>
${r1}
</body>
</html>`);
   })

    var server = app.listen(8081, function () {
        var host = server.address().address
        var port = server.address().port   
        console.log("Example app listening at http://%s:%s", host, port)
    })




// let array= [notFilledList]
// console.log(array)
// let wb = pkg.utils.book_new();
// let wsName ='sheet1'; 
// console.log(wsName)

// for(var i=0;i<sheetNameList.length;i++){

//    let  ws = pkg.utils.aoa_to_sheet(array);
//     console.log(ws)
//     wsName[i] += pkg.utils.book_append_sheet(wb, ws,wsName[i]);
// console.log(ws)
// }
//  pkg.writeFile(wb, 'out4.xlsx');


    // for(var i=0;i<notFilledList.length;i++ ){
    //     console.log("Employes who are not submmited on",sheetNameList[i],notFilledList[i]);
    
    //     }











