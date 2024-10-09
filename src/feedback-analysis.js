const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 输入和输出文件路径
const inputFilePath = path.join(__dirname, '../data/input/feedback.xlsx'); // 请确保input.xlsx在同一目录下
const outputFilePath = path.join(__dirname, '../data/output/feedback_sorted.xlsx');

// 读取Excel文件
function readExcel(filePath) {
    if (!fs.existsSync(filePath)) {
        console.error(`文件不存在: ${filePath}`);
        process.exit(1);
    }

    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0]; // 读取第一个工作表
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // 按行读取，header:1表示按数组返回

    // 假设数据在第一列
    const columnData = data.map(row => row[0]).filter(cell => typeof cell === 'string');

    return columnData;
}

// 处理数据：去重并按字符串长度从长到短排序
function processData(data) {
    // 去重
    const uniqueData = Array.from(new Set(data));

    // 按长度排序（从长到短）
    uniqueData.sort((a, b) => b.length - a.length);

    return uniqueData;
}

// 写入Excel文件
function writeExcel(data, filePath) {
    const worksheetData = data.map(item => [item]); // 转换为二维数组，每个元素为一行
    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    XLSX.writeFile(workbook, filePath);
    console.log(`处理后的数据已写入: ${filePath}`);
}

// 主函数
function main() {
    const inputData = readExcel(inputFilePath);
    const processedData = processData(inputData);
    writeExcel(processedData, outputFilePath);
}

main();
