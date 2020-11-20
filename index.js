// Require library
var excel = require("excel4node");

const testData = [
    {
        fecha: "2020-11-20",
        nroBoleta: 400,
        nombre: "Pepe Argento",
        bruto: 1112
    },
    {
        fecha: "2020-11-20",
        nroBoleta: 410,
        nombre: "Pepe Argento2",
        bruto: 132
    },
    {
        fecha: "2020-11-20",
        nroBoleta: 405,
        nombre: "Pepe Argento3",
        bruto: 1322
    },
    {
        fecha: "2020-11-20",
        nroBoleta: 406,
        nombre: "Pepe Argento4",
        bruto: 142
    },
    {
        fecha: "2020-11-20",
        nroBoleta: 401,
        nombre: "Pepe Argento5",
        bruto: 1325
    }
]


const createExcelFile = () => {
  // Create a new instance of a Workbook class
  var workbook = new excel.Workbook();

  // Add Worksheets to the workbook
  var worksheet = workbook.addWorksheet("Sheet 1");
  var worksheet2 = workbook.addWorksheet("Sheet 2");

  // Create a reusable style
  var style = workbook.createStyle({
    font: {
      color: "#000000",
      size: 12,
    },
    numberFormat: "$#,##0.00; ($#,##0.00); -",
  });

for (let i = 0; i < testData.length; i++) {
    let column = 1;
    const e = testData[i];
    for(var propertyName in e) {
        worksheet.cell(i + 1, column).string(e[propertyName].toString()).style(style);
        column ++;
     }    
}
  workbook.write(`${Date.now()}-${new Date().toDateString()}.xlsx`);
};

createExcelFile();
