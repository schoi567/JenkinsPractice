const XLSX = require('xlsx');
 
 

// Excel Files
var workbook = XLSX.readFile("Book1.xlsx",{type: 'binary', cellDates: true, dateNF: 'mm/dd/yyyy'});
let worksheet = workbook.Sheets[workbook.SheetNames[0]];
const dataExcel = XLSX.utils.sheet_to_json(worksheet);  

for (var i=0; i < dataExcel.length-1; i++) {
    var a= new Date(dataExcel[i].t).getDate(); 
    var a1= new Date(dataExcel[i].t).getFullYear();
    var a2= new Date(dataExcel[i].t).getMonth()+1; 
    var quote  = "\"";  
    var c= "{t: new Date("+quote+ a1.toString()+"-"+a2.toString()+"-"+ 
    a.toString()+quote+")" +"," + "  y: " + dataExcel[i].y +"},";  
    console.log(c);
}; 

for (var i= dataExcel.length-1; i < dataExcel.length; i++) {
    var a= new Date(dataExcel[i].t).getDate(); 
    var a1= new Date(dataExcel[i].t).getFullYear();
    var a2= new Date(dataExcel[i].t).getMonth()+1; 
    var quote  = "\"";  
    var c= "{t: new Date("+quote+ a1.toString()+"-"+a2.toString()+"-"+ 
    a.toString()+quote+")" +"," + "  y: " + dataExcel[i].y +"}";  
    console.log(c);
}; 




 
 
 

 
//  console.log(c);node ./FromExcel.js

//  console.log(dataExcel); 
 
//  first go to D:\Graph JS1\PATH\ChartforEveryDept\Months
//  node ./FromExcel.js