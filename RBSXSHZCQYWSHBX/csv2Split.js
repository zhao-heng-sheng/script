"use strict"
const fs = require("fs")
const xlsx = require("node-xlsx")
// 基本责任
var rateXlsx = xlsx.parse("./03华贵附加麦芽糖失能收入损失保险（互联网专属）费率表.xlsx")[0].data

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
    let termCode = ""
    let termPayment = ""
    let chargeCodeArr = []
    let genderArr = []
    for (let i = min; i < max; i++) {
        if (rateArr[i][factor.jMin-1] && isNaN(rateArr[i][factor.jMin-1])) {
            if (rateArr[i][factor.jMin-1].indexOf("保险期间：") != -1) {
                termCode = termCodeFormat(rateArr[i][factor.jMin-1].split("    ")[1].split("保险期间：")[1])
            }
            if (rateArr[i][factor.jMin-1].indexOf("给付期限：") != -1) {
                termPayment = parseInt(rateArr[i][factor.jMin-1].split("：")[1])
            }
            if (rateArr[i][factor.jMin-1].indexOf("交费方式") != -1) {
                chargeCodeArr = rateArr[i]
            }
            if (rateArr[i][factor.jMin-1].indexOf("性别") != -1) {
                genderArr = rateArr[i]
            }
        }

        //行循环
        for (let j = factor.jMin || 1; j < (factor.jMax || rateArr[i].length); j++) {
            //列循环
            if (rateArr[i][j] && rateArr[i][j] != "-" && !isNaN(rateArr[i][factor.jMin-1])) {
                data.push([!chargeCodeArr[j] ? chargeCodeFormat(chargeCodeArr[j - 1]) : chargeCodeFormat(chargeCodeArr[j]), rateArr[i][factor.jMin-1], genderFormat(genderArr[j]), termCode, termPayment, ":", String(rateArr[i][j]).replace(",", "")])
            }
        }
    }
    return data
}
let chargeCodeFormat = function (code) {
    if (!isNaN(parseInt(code))) return String(parseInt(code))
    if (code.indexOf("一次交清") != -1) return "1"
    if (code.indexOf("至") != -1) return parseInt(code.split("至")[1]) + "y"
}
let termCodeFormat = (code) => {
    if (code.indexOf("至被保险人") != -1) {
        return parseInt(code.split("至被保险人")[1]) + "y"
    } else {
        return parseInt(code) + ""
    }
}
let genderFormat = function (gender) {
    if (gender.indexOf("男") != -1) return "1"
    if (gender.indexOf("女") != -1) return "2"
    return ""
}
//最后导出的数据
let data = [["交费期间-chargeCode", "被保人年龄-insuredAge", "性别-gender", "保险期间-termCode", "给付期限-termPayment", ":", "费率/保费"]]
data.push(
    ...pushDate(rateXlsx, 1, rateXlsx.length, {
        jMin: 1,
        jMax: 17,
    })
)
data.push(
    ...pushDate(rateXlsx, 1, rateXlsx.length, {
        jMin: 18,
    })
)
var buffer = xlsx.build([{ name: "table1", data: data }])
var filePath = "./tex1111.xlsx"
fs.writeFileSync(filePath, buffer, { flag: "w" })
