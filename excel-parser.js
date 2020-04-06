
const Excel = require('exceljs')
const path = require('path')
const mapFile = require('./map.json')
var os = require("os");
const filename = path.join(__dirname, '/files/table.xlsx')
var workbook = new Excel.Workbook()

function getCombinedData(data, element){
    var tags = element.split('.')
    var combinedData = ''
    tags.forEach(function(tag){
        combinedData += data[tag] + ', '
    })
    return combinedData
}

function getValue (data, mapArr) {
    mapArr.forEach(function(element){
        if(element.includes('.')){
            data = getCombinedData(data, element)
        } else {
            if (data != null){
                if (Array.isArray(data[element])){
                    data = data[element].length != 0 ? data[element][0] : null
                } else {
                    data = data[element] 
                }
            }
        }
    })
    return data
}

function parseToExcel (json, innIndex) {
    workbook.xlsx.readFile(filename)
    .then(function() {
      let workSheet = workbook.getWorksheet(1)
      let row = workSheet.lastRow
      Object.keys(mapFile).forEach(function(element){
          var value = ''
          //ячейки 29, 82, 88, 94 требуют специальной обработки
          if (element == "29"){
            let founders = json[mapFile[element][0]]
            for (let [key, founderType] of Object.entries(founders)) {
                if (founderType.length > 0){
                    founderType.forEach(function(founder){
                        value += founder["FullName"] + ': ' + founder["DepositAsFounder"] + "(номинальная стоимость доли в руб.)" + os.EOL
                    })
                }
              }
          } else if(element == "82") {
            let nacesArr = json[mapFile[element][0]][mapFile[element][1]]
            nacesArr.forEach(function(nace){
                value += nace[mapFile[element][2]] + ", "
            })
          } else if(element == "88") {
            let licences = json[mapFile[element][0]]
            for (let [key, licence] of Object.entries(licences)) {
                if(Array.isArray(licence)) {
                    licence.forEach(function(data){
                        value += data.RegNumber + ': ' + data.DateIncome + ';' + os.EOL 
                    })
                }
              }
          } else if(element == "94") {
            let changesArr = json[mapFile[element][0]][mapFile[element][1]]
            if (changesArr.length != 0){
                let arrLength = changesArr.length
                value = changesArr[arrLength-1][mapFile[element][2]]
            }
          } else {
            value = getValue(json, mapFile[element])
          }
        row.getCell(parseInt(element)).value = value
      })
      
    })
    .then(function() {
       return workbook.xlsx.writeFile(filename)
    })
    .catch((err) => {
        console.log(err)
    });
}






module.exports = {
    parseToExcel
  };