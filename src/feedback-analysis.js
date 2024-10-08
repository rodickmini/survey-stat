const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// 设置文件路径
const inputFilePath = path.join(__dirname, '../data/input/feedback.xlsx'); // 替换为你的Excel文件名
const outputFilePath = path.join(__dirname, '../data/output/categorized_feedback.xlsx'); // 输出的文件名

// 读取Excel文件
const workbook = xlsx.readFile(inputFilePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// 将工作表转换为JSON数组，每一行是一个对象
const jsonData = xlsx.utils.sheet_to_json(worksheet);

// 获取反馈数据列
const feedbackColumn = '第九题：您对IM产品还有哪些其他的建议或意见？';

// 定义分类
const categories = {
    '功能改进': ['增加', '改进', '功能', '增强', '清除缓存'],
    'bug反馈': ['bug', '问题', '卡顿', '崩溃'],
    '无意见': ['无', '没有', '暂时没有', '暂无'],
    '其他': []
};

// 分类函数
function categorizeFeedback(feedback) {
    for (let category in categories) {
        const keywords = categories[category];
        if (keywords.some(keyword => feedback.includes(keyword))) {
            return category;
        }
    }
    return '其他';
}

// 初始化分类结果
const categorizedFeedback = {
    '功能改进': [],
    'bug反馈': [],
    '无意见': [],
    '其他': []
};

// 处理反馈数据
jsonData.forEach(row => {
    const feedback = row[feedbackColumn] ? String(row[feedbackColumn]) : '';  // 确保反馈是字符串
    const category = categorizeFeedback(feedback);
    categorizedFeedback[category].push(feedback);
});

// 准备写入Excel的格式
const outputData = [];

// 将分类结果转换为适合写入Excel的格式
for (let category in categorizedFeedback) {
    categorizedFeedback[category].forEach(feedback => {
        outputData.push({ '分类': category, '反馈': feedback });
    });
}

// 将分类后的反馈数据转换为Excel工作表
const newWorksheet = xlsx.utils.json_to_sheet(outputData);

// 创建新的工作簿
const newWorkbook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, '分类反馈');

// 保存到新的Excel文件
xlsx.writeFile(newWorkbook, outputFilePath);

console.log(`分类反馈结果已保存到 ${outputFilePath}`);
