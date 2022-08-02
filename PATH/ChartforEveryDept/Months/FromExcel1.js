const XLSX = require('xlsx');
 
const Janurary = ["C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-04-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-05-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-06-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-07-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-10-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-11-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-12-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-13-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-18-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-19-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-20-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-21-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-24-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-25-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-26-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-27-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-28-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-31-22.xlsx",
];
const Feburary = ["C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-01-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-02-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-03-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-04-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-07-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-08-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-09-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-10-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-11-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-14-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-15-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-16-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-17-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-18-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-21-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-22-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-23-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-24-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-25-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-28-22.xlsx"];

const March = ["C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-01-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-02-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-03-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-04-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-07-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-08-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-09-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-10-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-11-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-14-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-15-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-16-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-17-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-18-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-21-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-22-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-23-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-24-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-25-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-28-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-29-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-30-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-31-22.xlsx"
];
const April = [
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-01-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-04-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-05-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-06-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-07-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-08-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-11-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-12-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-13-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-14-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-18-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-19-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-20-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-21-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-22-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-25-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-26-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-27-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-28-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-29-22.xlsx"];

const May = [
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-02-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-03-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-04-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-05-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-06-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-09-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-10-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-11-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-12-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-13-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-16-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-17-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-18-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-19-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-20-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-23-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-24-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-25-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-26-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-27-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-31-22.xlsx"
];

const June = [
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-1-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-2-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-3-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-6-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-7-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-8-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-9-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-10-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-13-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-14-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-15-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-16-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-17-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-20-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-21-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-22-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-23-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-24-2022.xlsx",

];






const Total= ["C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-04-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-05-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-06-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-07-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-10-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-11-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-12-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-13-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-18-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-19-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-20-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-21-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-24-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-25-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-26-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-27-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-28-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-31-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-01-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-02-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-03-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-04-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-07-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-08-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-09-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-10-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-11-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-14-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-15-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-16-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-17-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-18-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-21-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-22-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-23-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-24-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-25-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 02-28-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-01-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-02-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-03-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-04-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-07-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-08-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-09-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-10-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-11-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-14-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-15-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-16-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-17-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-18-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-21-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-22-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-23-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-24-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-25-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-28-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-29-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-30-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 03-31-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-01-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-04-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-05-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-06-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-07-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-08-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-11-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-12-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-13-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-14-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-18-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-19-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-20-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-21-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-22-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-25-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-26-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-27-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-28-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 04-29-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-02-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-03-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-04-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-05-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-06-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-09-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-10-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-11-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-12-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-13-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-16-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-17-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-18-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-19-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-20-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-23-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-24-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-25-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-26-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-27-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 05-31-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-1-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-2-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-3-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-6-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-7-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-8-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-9-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-10-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-13-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-14-22.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-15-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-16-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-17-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-20-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-21-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-22-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-23-2022.xlsx",
"C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 06-24-2022.xlsx",  ];







 /* function formatAsPercent(num) {return new Intl.NumberFormat('default', {style: 'percent', minimumFractionDigits: 2,
maximumFractionDigits: 2, }).format(num);} */    
 /*

for (const element of Janurary) {
  var workbook = XLSX.readFile(element);
  let worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
  var address_of_cell = 'K7';
  var desired_cell = worksheet[address_of_cell];
  var desired_value = desired_cell.v;
let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
let element2 = element1.replace(".xlsx", "");  
var quote  = "\"";  
var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
console.log(c);
 
} 
*/ 
 /*

for (const element of Janurary) {
  var workbook = XLSX.readFile(element);
  let worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
  var address_of_cell = 'K6';
  var desired_cell = worksheet[address_of_cell];
  var desired_value = desired_cell.v;    
let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
let element2 = element1.replace(".xlsx", "");  
var quote  = "\"";  
var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
console.log(c);
//  console.log(element2+"  "+formatAsPercent(desired_value));
}

*/ 

 /*


for (const element of Janurary) {
  var workbook = XLSX.readFile(element);
  let worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
  var address_of_cell = 'K8';
  var desired_cell = worksheet[address_of_cell];
  var desired_value = desired_cell.v;    
let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
let element2 = element1.replace(".xlsx", "");  
var quote  = "\"";  
var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
console.log(c);
//  console.log(element2+"  "+formatAsPercent(desired_value));
}
*/  

 /*
for (const element of Janurary) {
  var workbook = XLSX.readFile(element);
  let worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
  var address_of_cell = 'K9';
  var desired_cell = worksheet[address_of_cell];
  var desired_value = desired_cell.v;    
let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
let element2 = element1.replace(".xlsx", "");  
var quote  = "\"";  
var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
console.log(c);

}

*/  


 /*
for (const element of Janurary) {
  var workbook = XLSX.readFile(element);
  let worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
  var address_of_cell = 'K10';
  var desired_cell = worksheet[address_of_cell];
  var desired_value = desired_cell.v;    
let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
let element2 = element1.replace(".xlsx", "");  
var quote  = "\"";  
var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
console.log(c);

}
*/  
 /*
for (const element of Janurary) {
  var workbook = XLSX.readFile(element);
  let worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
  var address_of_cell = 'K11';
  var desired_cell = worksheet[address_of_cell];
  var desired_value = desired_cell.v;    
let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
let element2 = element1.replace(".xlsx", "");  
var quote  = "\"";  
var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
console.log(c);

}

*/  

/*
for (const element of Janurary) {
  var workbook = XLSX.readFile(element);
  let worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
  var address_of_cell = 'J6';
  var desired_cell = worksheet[address_of_cell];
  var desired_value = desired_cell.v;    
let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
let element2 = element1.replace(".xlsx", "");  
var quote  = "\"";  
var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value +"},";  
console.log(c);

}


// Please take out last comma. 














/*

var workbook = XLSX.readFile("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-04-22.xlsx");
let worksheet = workbook.Sheets[workbook.SheetNames[0]];
const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
var address_of_cell = 'K7';
var desired_cell = worksheet[address_of_cell];
var desired_value = desired_cell.v;
function formatAsPercent(num) {
    return new Intl.NumberFormat('default', {
      style: 'percent',
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(num);}
  console.log(formatAsPercent(desired_value));


var workbook1 = XLSX.readFile("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount 01-05-22.xlsx");
let worksheet1 = workbook1.Sheets[workbook1.SheetNames[0]];
const dataExcel1 = XLSX.utils.sheet_to_json(worksheet1);  
var address_of_cell1 = 'K7';
var desired_cell1 = worksheet1[address_of_cell1];
var desired_value1 = desired_cell1.v;
function formatAsPercent(num) {
    return new Intl.NumberFormat('default', {
      style: 'percent',
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(num);}
  console.log(formatAsPercent(desired_value1));

  */ 

/*
  for (const element of Janurary) {
    var workbook = XLSX.readFile(element);
    let worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
    var address_of_cell = 'K12';
    var desired_cell = worksheet[address_of_cell];
    var desired_value = desired_cell.v;    
  let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
  let element2 = element1.replace(".xlsx", "");  
  var quote  = "\"";  
  var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
  console.log(c);
  //  console.log(element2+"  "+formatAsPercent(desired_value));
  }


  */ 

/*
  for (const element of Janurary) {
    var workbook = XLSX.readFile(element);
    let worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
    var address_of_cell = 'J7';
    var desired_cell = worksheet[address_of_cell];
    var desired_value = desired_cell.v;    
  let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
  let element2 = element1.replace(".xlsx", "");  
  var quote  = "\"";  
  var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
  console.log(c);
  }  */ 
/*
  for (const element of Janurary) {
    var workbook = XLSX.readFile(element);
    let worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
    var address_of_cell = 'J8';
    var desired_cell = worksheet[address_of_cell];
    var desired_value = desired_cell.v;    
  let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
  let element2 = element1.replace(".xlsx", "");  
  var quote  = "\"";  
  var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
  console.log(c);
  }

 */ 
/*
  for (const element of Janurary) {
    var workbook = XLSX.readFile(element);
    let worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
    var address_of_cell = 'J12';
    var desired_cell = worksheet[address_of_cell];
    var desired_value = desired_cell.v;    
  let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
  let element2 = element1.replace(".xlsx", "");  
  var quote  = "\"";  
  var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
  console.log(c);
  };

 */ 
/*
  const Production = [
    {t: new Date("01-04-22"),  y: 9.210526315789473},
    {t: new Date("01-05-22"),  y: 15.789473684210526},
    {t: new Date("01-06-22"),  y: 13.157894736842104},
    {t: new Date("01-07-22"),  y: 14.473684210526317},
    {t: new Date("01-10-22"),  y: 18.421052631578945},
    {t: new Date("01-11-22"),  y: 11.842105263157894},
    {t: new Date("01-12-22"),  y: 9.210526315789473},
    {t: new Date("01-13-22"),  y: 6.578947368421052},
    {t: new Date("01-18-22"),  y: 10.526315789473683},
    {t: new Date("01-19-22"),  y: 10.526315789473683},
    {t: new Date("01-20-22"),  y: 8.571428571428571},
    {t: new Date("01-21-22"),  y: 15.714285714285714},
    {t: new Date("01-24-22"),  y: 8.571428571428571},
    {t: new Date("01-25-22"),  y: 2.857142857142857},
    {t: new Date("01-26-22"),  y: 8.571428571428571},
    {t: new Date("01-27-22"),  y: 11.428571428571429},
    {t: new Date("01-28-22"),  y: 11.428571428571429},
    {t: new Date("01-31-22"),  y: 15.714285714285714}
    ]
    var a = 0; 
    for (i=0; i < Production.length; i++)
    {  a+=Production[i].y; }
   a = a/Production.length;
console.log(a);
   
 */ 
/*
console.log();
console.log("Materials");
for (const element of June) {
  var workbook = XLSX.readFile(element);
  let worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
  var address_of_cell = 'K6';
  var desired_cell = worksheet[address_of_cell];
  var desired_value = desired_cell.v;    
let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
let element2 = element1.replace(".xlsx", "");  
var quote  = "\"";  
var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
console.log(c);
};


console.log();
console.log("Injection");

for (const element of June) {
  var workbook = XLSX.readFile(element);
  let worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
  var address_of_cell = 'K7';
  var desired_cell = worksheet[address_of_cell];
  var desired_value = desired_cell.v;    
let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
let element2 = element1.replace(".xlsx", "");  
var quote  = "\"";  
var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
console.log(c);
};

console.log();
console.log("Assembly Wiley Road");

for (const element of June) {
  var workbook = XLSX.readFile(element);
  let worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
  var address_of_cell = 'K8';
  var desired_cell = worksheet[address_of_cell];
  var desired_value = desired_cell.v;    
let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
let element2 = element1.replace(".xlsx", "");  
var quote  = "\"";  
var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
console.log(c);
};


console.log();
console.log("Redding/Jane");

for (const element of June) {
  var workbook = XLSX.readFile(element);
  let worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
  var address_of_cell = 'K9';
  var desired_cell = worksheet[address_of_cell];
  var desired_value = desired_cell.v;    
let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
let element2 = element1.replace(".xlsx", "");  
var quote  = "\"";  
var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
console.log(c);
};


console.log();
console.log("Quality");

for (const element of June) {
  var workbook = XLSX.readFile(element);
  let worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
  var address_of_cell = 'K10';
  var desired_cell = worksheet[address_of_cell];
  var desired_value = desired_cell.v;    
let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
let element2 = element1.replace(".xlsx", "");  
var quote  = "\"";  
var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
console.log(c);
};


console.log();
console.log("Shipping");

for (const element of June) {
  var workbook = XLSX.readFile(element);
  let worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
  var address_of_cell = 'K11';
  var desired_cell = worksheet[address_of_cell];
  var desired_value = desired_cell.v;    
let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
let element2 = element1.replace(".xlsx", "");  
var quote  = "\"";  
var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
console.log(c);
};


console.log();
console.log("Total for SEG");

for (const element of June) {
  var workbook = XLSX.readFile(element);
  let worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
  var address_of_cell = 'K12';
  var desired_cell = worksheet[address_of_cell];
  var desired_value = desired_cell.v;    
let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
let element2 = element1.replace(".xlsx", "");  
var quote  = "\"";  
var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
console.log(c);
};


console.log();
console.log("Total for Staffing");

for (const element of June) {
  var workbook = XLSX.readFile(element);
  let worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
  var address_of_cell = 'J12';
  var desired_cell = worksheet[address_of_cell];
  var desired_value = desired_cell.v;    
let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
let element2 = element1.replace(".xlsx", "");  
var quote  = "\"";  
var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
console.log(c);
};

 */ 


 /*
console.log('Total for Staffing');
for (const element of Total) {
  var workbook = XLSX.readFile(element);
  let worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
  var address_of_cell = 'J12';
  var desired_cell = worksheet[address_of_cell];
  var desired_value = desired_cell.v;    
let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
let element2 = element1.replace(".xlsx", "");  
var quote  = "\"";  
var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
console.log(c);
};
console.log('Total for SEG');

for (const element of Total) {
  var workbook = XLSX.readFile(element);
  let worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
  var address_of_cell = 'K12';
  var desired_cell = worksheet[address_of_cell];
  var desired_value = desired_cell.v;    
let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
let element2 = element1.replace(".xlsx", "");  
var quote  = "\"";  
var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
console.log(c);
};
 
console.log('Total for All');

for (const element of Total) {
  var workbook = XLSX.readFile(element);
  let worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const dataExcel = XLSX.utils.sheet_to_json(worksheet);  
  var address_of_cell = 'L12';
  var desired_cell = worksheet[address_of_cell];
  var desired_value = desired_cell.v;    
let element1 = element.replace("C:/Users/SOOMIN-SAC-217/Graph JS/PATH/ExcelSheetShow/ExcelSheet/Total/Daily Headcount ", "");  
let element2 = element1.replace(".xlsx", "");  
var quote  = "\"";  
var c= "{t: new Date("+quote+ element2+quote+")" +"," + "  y: " + desired_value*100 +"},";  
console.log(c);
};

*/ 

const Total1 = [{t: new Date("01-04-22"),  y: 22.282608695652172},
{t: new Date("01-05-22"),  y: 23.097826086956523},
{t: new Date("01-06-22"),  y: 25.271739130434785},
{t: new Date("01-07-22"),  y: 28.804347826086957},
{t: new Date("01-10-22"),  y: 30.706521739130434},
{t: new Date("01-11-22"),  y: 26.902173913043477},
{t: new Date("01-12-22"),  y: 21.195652173913043},
{t: new Date("01-13-22"),  y: 25.271739130434785},
{t: new Date("01-18-22"),  y: 19.294117647058822},
{t: new Date("01-19-22"),  y: 19.52941176470588},
{t: new Date("01-20-22"),  y: 12.995594713656388},
{t: new Date("01-21-22"),  y: 14.096916299559473},
{t: new Date("01-24-22"),  y: 15.418502202643172},
{t: new Date("01-25-22"),  y: 12.995594713656388},
{t: new Date("01-26-22"),  y: 13.43612334801762},
{t: new Date("01-27-22"),  y: 14.537444933920703},
{t: new Date("01-28-22"),  y: 15.418502202643172},
{t: new Date("01-31-22"),  y: 17.62114537444934},
{t: new Date("02-01-22"),  y: 13.43612334801762},
{t: new Date("02-02-22"),  y: 11.013215859030836},
{t: new Date("02-03-22"),  y: 11.899791231732777},
{t: new Date("02-04-22"),  y: 14.40501043841336},
{t: new Date("02-07-22"),  y: 13.987473903966595},
{t: new Date("02-08-22"),  y: 16.075156576200417},
{t: new Date("02-09-22"),  y: 11.482254697286013},
{t: new Date("02-10-22"),  y: 14.40501043841336},
{t: new Date("02-11-22"),  y: 7.306889352818372},
{t: new Date("02-14-22"),  y: 23.173277661795407},
{t: new Date("02-15-22"),  y: 16.49269311064718},
{t: new Date("02-16-22"),  y: 13.987473903966595},
{t: new Date("02-17-22"),  y: 12.526096033402922},
{t: new Date("02-18-22"),  y: 15.24008350730689},
{t: new Date("02-21-22"),  y: 19.206680584551147},
{t: new Date("02-22-22"),  y: 14.19624217118998},
{t: new Date("02-23-22"),  y: 12.734864300626306},
{t: new Date("02-24-22"),  y: 17.954070981210858},
{t: new Date("02-25-22"),  y: 18.37160751565762},
{t: new Date("02-28-22"),  y: 17.745302713987474},
{t: new Date("03-01-22"),  y: 15.031315240083506},
{t: new Date("03-02-22"),  y: 11.899791231732777},
{t: new Date("03-03-22"),  y: 13.569937369519833},
{t: new Date("03-04-22"),  y: 13.987473903966595},
{t: new Date("03-07-22"),  y: 11.064718162839249},
{t: new Date("03-08-22"),  y: 10.647181628392484},
{t: new Date("03-09-22"),  y: 9.394572025052192},
{t: new Date("03-10-22"),  y: 10.438413361169102},
{t: new Date("03-11-22"),  y: 12.734864300626306},
{t: new Date("03-14-22"),  y: 16.2839248434238},
{t: new Date("03-15-22"),  y: 10.647181628392484},
{t: new Date("03-16-22"),  y: 12.943632567849686},
{t: new Date("03-17-22"),  y: 8.350730688935283},
{t: new Date("03-18-22"),  y: 12.31732776617954},
{t: new Date("03-21-22"),  y: 13.361169102296449},
{t: new Date("03-22-22"),  y: 11.273486430062631},
{t: new Date("03-23-22"),  y: 12.10855949895616},
{t: new Date("03-24-22"),  y: 9.394572025052192},
{t: new Date("03-25-22"),  y: 8.977035490605429},
{t: new Date("03-28-22"),  y: 12.526096033402922},
{t: new Date("03-29-22"),  y: 8.350730688935283},
{t: new Date("03-30-22"),  y: 13.569937369519833},
{t: new Date("03-31-22"),  y: 16.2839248434238},
{t: new Date("04-01-22"),  y: 16.075156576200417},
{t: new Date("04-04-22"),  y: 12.526096033402922},
{t: new Date("04-05-22"),  y: 10.855949895615867},
{t: new Date("04-06-22"),  y: 10.438413361169102},
{t: new Date("04-07-22"),  y: 10.22964509394572},
{t: new Date("04-08-22"),  y: 10.438413361169102},
{t: new Date("04-11-22"),  y: 12.31732776617954},
{t: new Date("04-12-22"),  y: 5.643340857787811},
{t: new Date("04-13-22"),  y: 14.446952595936793},
{t: new Date("04-14-22"),  y: 13.544018058690746},
{t: new Date("04-18-22"),  y: 12.18961625282167},
{t: new Date("04-19-22"),  y: 11.286681715575622},
{t: new Date("04-20-22"),  y: 9.932279909706546},
{t: new Date("04-21-22"),  y: 9.255079006772009},
{t: new Date("04-22-22"),  y: 8.577878103837472},
{t: new Date("04-25-22"),  y: 12.641083521444695},
{t: new Date("04-26-22"),  y: 10.835214446952596},
{t: new Date("04-27-22"),  y: 6.772009029345373},
{t: new Date("04-28-22"),  y: 7.44920993227991},
{t: new Date("04-29-22"),  y: 8.35214446952596},
{t: new Date("05-02-2022"),  y: 11.963882618510159},
{t: new Date("05-03-2022"),  y: 10.609480812641085},
{t: new Date("05-04-2022"),  y: 12.415349887133182},
{t: new Date("05-05-2022"),  y: 10.383747178329571},
{t: new Date("05-06-2022"),  y: 12.18961625282167},
{t: new Date("05-09-2022"),  y: 15.575620767494355},
{t: new Date("05-10-2022"),  y: 10.383747178329571},
{t: new Date("05-11-22"),  y: 10.15801354401806},
{t: new Date("05-12-22"),  y: 8.577878103837472},
{t: new Date("05-13-22"),  y: 8.577878103837472},
{t: new Date("05-16-22"),  y: 13.769751693002258},
{t: new Date("05-17-22"),  y: 11.060948081264108},
{t: new Date("05-18-22"),  y: 10.15801354401806},
{t: new Date("05-19-22"),  y: 12.18961625282167},
{t: new Date("05-20-22"),  y: 11.286681715575622},
{t: new Date("05-23-22"),  y: 15.349887133182843},
{t: new Date("05-24-22"),  y: 8.126410835214447},
{t: new Date("05-25-22"),  y: 10.383747178329571},
{t: new Date("05-26-22"),  y: 10.15801354401806},
{t: new Date("05-27-22"),  y: 9.029345372460497},
{t: new Date("05-31-22"),  y: 11.738148984198645},
{t: new Date("06-1-22"),  y: 9.706546275395034},
{t: new Date("06-2-22"),  y: 11.738148984198645},
{t: new Date("06-3-22"),  y: 10.383747178329571},
{t: new Date("06-6-22"),  y: 13.99548532731377},
{t: new Date("06-7-22"),  y: 12.641083521444695},
{t: new Date("06-8-22"),  y: 10.835214446952596},
{t: new Date("06-9-22"),  y: 10.609480812641085},
{t: new Date("06-10-22"),  y: 10.15801354401806},
{t: new Date("06-13-22"),  y: 11.738148984198645},
{t: new Date("06-14-22"),  y: 9.932279909706546},
{t: new Date("06-15-2022"),  y: 12.44131455399061},
{t: new Date("06-16-2022"),  y: 7.981220657276995},
{t: new Date("06-17-2022"),  y: 9.15492957746479},
{t: new Date("06-20-2022"),  y: 13.615023474178404},
{t: new Date("06-21-2022"),  y: 13.145539906103288},
{t: new Date("06-22-2022"),  y: 12.44131455399061},
{t: new Date("06-23-2022"),  y: 10.328638497652582},
{t: new Date("06-24-2022"),  y: 11.737089201877934}  ];


var a = 0; 
for (i=0; i < Total1.length; i++)
{  a+=Total1[i].y; }
a = a/Total1.length;
console.log(a);


//  console.log(c);

//  console.log(dataExcel); 
 

//  node ./FromExcel1.js

 
