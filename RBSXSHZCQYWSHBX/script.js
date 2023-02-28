"use strict"
const fs = require("fs")
// 引入node-xlsx
const xlsx = require("node-xlsx")
// 读取excel文件
let rateXlsx = xlsx.parse("./人保寿险守护者长期意外伤害保险（互联网专属）费率表(1).xlsx")[0].data

let getChargeCode = (str) => {
    return parseInt(str) + ""
}
let getTermCode = (str) => {
    str = str.split("保")[str.split("保").length - 1]
    if (isNaN(parseInt(str))) return parseInt(str.slice(1)) + "y"
    return parseInt(str) + ""
}
let getGenderCode = (str) => {
    return str === "男性" ? "1" : "2"
}
let getPlanCode = (str) => {
    let planMap = {
        "基本部分（每万元基本保险金额）": "A",
        "基本+可选1（每万元基本保险金额）": "B",
        "基本+可选2（每万元基本保险金额）": "C",
        "基本+可选1+可选2（每万元基本保险金额）": "D",
    }
    return planMap[str]
}
/**
 *
 * @param {number} row  所需数据在第几行
 * @param {number} col  当前遍历到第几列
 * @returns {number}
 */
let getIndex = (row, col) => {
    for (; col >= 0; col--) {
        //不为null即为合并行的数据
        if (rateXlsx[row][col]) return col
    }
}
/**
 *
 * @param {Array} rateArr  表格数据
 * @param {number} min  开始位置
 * @param {number} max  结束位置
 * @returns {array}
 */
let pushDate = function (rateArr, min, max) {
    min = min - 1
    let data = []
    for (let i = min; i < max; i++) {
        //行循环
        for (let j = 1; j < rateArr[i].length; j++) {
            //列循环
            if (rateArr[i][j] && rateArr[i][j] != "-") {
                data.push([getChargeCode(rateArr[3][j]), getTermCode(rateArr[1][getIndex(1, j)]), rateArr[i][0], getGenderCode(rateArr[2][getIndex(2, j)]), getPlanCode(rateArr[0][getIndex(0, j)]), ":", String(rateArr[i][j]).replace(",", "")])
            }
        }
    }
    return data
}

let buildData = pushDate(rateXlsx, 5, 57)
buildData.unshift(["交费年期-chargeCode", "保险期间-termCode", "被保人年龄-insuredAge", "性别-gender", "保障计划-planCode", ":", "费率/保费"])

var buffer = xlsx.build([{ name: "table1", data: buildData }])
var filePath = "./newXlsx.xlsx"
fs.writeFileSync(filePath, buffer, { flag: "w" })