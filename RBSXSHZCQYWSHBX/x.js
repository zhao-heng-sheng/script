"use strict"
console.log("1231")

const fs = require("fs")
const xlsx = require("node-xlsx")
var rateXlsx = xlsx.parse("./111.xlsx")
console.log(rateXlsx);

/**
 *
 * @param {Array} rateArr  表格数据
 * @param {number} min  开始位置(开始表格行，不是下标)
 * @param {number} max  结束位置(结束表格行，不是下标)
 * @param {object} factor 保费相关信息
 */
let pushDate = function (rateArr, min, max, factor) {
    let data = []
    min = min - 1
    for (let i = min; i < max; i++) {
        //行循环
        for (let j = 1; j < rateArr[i].length; j++) {
            //列循环
            if (rateArr[i][j] && rateArr[i][j] != "-") {
                data.push(
                    [
                        !rateArr[3][j]?chargeCodeFormat(rateArr[2][j-1]):chargeCodeFormat(rateArr[2][j]),
                        !rateArr[1][j]?,
                        genderFormat(rateArr[3][j]),
                        '终身',
                        rateArr[1][0].indexOf('疾病关爱')===-1?'0':'1',
                        rateArr[1][0].indexOf('重大疾病')===-1?'0':'1',
                        rateArr[1][0].indexOf('身故或全残')===-1?'0':'1',
                        ':',
                        String(rateArr[i][j]).replace(',','')
                    ]
                )
            }
        }
    }
    return data
}
let chargeCodeFormat = function (code) {
    if(!isNaN(parseInt(code))) return String(parseInt(code))
    if (code.indexOf("趸交")!=-1 ) return "1"
    if (code === "5年期交") return "5"
    if (code === "10年期交") return "10"
    if (code === "15年期交") return "15"
    if (code === "20年期交") return "20"
    if (code === "30年期交") return "30"
}
let genderFormat = function(gender){
    if(gender.indexOf('男')!=-1) return '1'
    return '2'
}
//最后导出的数据
let data = [["交费年期-chargeCode", "保险期间-termCode","被保人年龄-insuredAge","性别-gender","保障计划-planCode", ":", "费率/保费"]]
// for(let i =0;i<rateXlsx.length-1;i++){
    data.push(...pushDate(rateXlsx, 5, 57))
// }
console.log(data)
var buffer = xlsx.build([{ name: "table1", data: data }])
var filePath = "./tex1111.xlsx"
fs.writeFileSync(filePath, buffer, { flag: "w" })
