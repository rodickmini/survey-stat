const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// 设置文件路径
const inputFilePath = path.join(__dirname, '../data/input/survey_collection.xls'); // 替换为你的Excel文件名
const outputFilePath = path.join(__dirname, '../data/output/survey_results_summary.xlsx');

// 读取Excel文件
const workbook = xlsx.readFile(inputFilePath);

// 假设数据在第一个工作表
const sheetName = workbook.SheetNames[1];
const worksheet = workbook.Sheets[sheetName];

// 将工作表转换为JSON数组，每一行是一个对象
const jsonData = xlsx.utils.sheet_to_json(worksheet);

//console.log(jsonData);

// 获取所有问题的列名
const questions = Object.keys(jsonData[0]);

// 初始化结果对象
const results = {};

// 定义所有可能的选项
const options = ['A', 'B', 'C', 'D', 'E'];

// 统计每道题每个选项的数量
questions.forEach((question) => {
    results[question] = {};
    options.forEach((option) => {
        // 过滤出当前选项的数量
        const count = jsonData.filter(row => row[question] === option).length;
        results[question][option] = count;
    });
});

// 打印结果
console.log('统计结果：');
console.log(results);

// 将结果转换为适合写入Excel的格式
const summaryData = [];

Object.keys(results).forEach((question) => {
    const row = { 题目: question };
    options.forEach((option) => {
        row[`选项${option}`] = results[question][option];
    });
    summaryData.push(row);
});

// 创建新的工作表
const summaryWorksheet = xlsx.utils.json_to_sheet(summaryData);

// 创建新的工作簿并添加工作表
const summaryWorkbook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(summaryWorkbook, summaryWorksheet, '统计结果');

// 保存到新的Excel文件
xlsx.writeFile(summaryWorkbook, outputFilePath);

console.log(`统计结果已保存到 ${outputFilePath}`);
