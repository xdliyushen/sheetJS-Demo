let xlsx = require('xlsx');
let fs = require('fs');

// let filePath = 'graduate.json';
let filePath = 'testData.json';

fs.readFile(filePath, 'utf8', (err, json) => {
    if (err) console.log(err);

    // 原始json数据
    let data = JSON.parse(json).rows;
    // 空的excel表格工作区
    let workbook = xlsx.utils.book_new();
    // 数据表
    let excelData = [];

    // 填写表名
    excelData.push(['课题名称', '课题简介', '指导老师姓名', '指导老师单位', '审题学院', '题目难度', '是否重点扶持', '课题来源', '课题性质', '邮箱', '手机号码'])

    // 需要提取的键名
    let dataKeys = [
        'KTMC',
        'KTJJ',
        'XM',
        'SZDWDM_DISPLAY',
        'MXXY_DISPLAY',
        'YL5_DISPLAY',
        'YL4_DISPLAY',
        'TMLY_DISPLAY',
        'KTXZ_DISPLAY',
        'YX',
        'SJHM'
    ]

    // 数据获取的时候是按照页来获取的, 所以这里用页做区分
    for (let page of data) {
        let rows = page.datas.cxkbmjsktdxz.rows;

        // 提取数据
        for (let row of rows) {
            let arr = [];
            for (let key of dataKeys) {
                arr.push(row[key]);
            }

            excelData.push(arr);
        }
    }

    // 生成表格
    let worksheet = xlsx.utils.aoa_to_sheet(excelData);
    xlsx.utils.book_append_sheet(workbook, worksheet, '毕业设计题目')
    xlsx.writeFile(workbook, 'graduate.xlsx');
})