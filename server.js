const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const port = 3000;

// 创建日志目录和日志函数
const logDir = './logs';
if (!fs.existsSync(logDir)) {
    fs.mkdirSync(logDir);
}

// 日志函数
function logMessage(message) {
    const timestamp = new Date().toISOString();
    const logEntry = `[${timestamp}] ${message}`;
    console.log(logEntry); // 同时在控制台输出
    
    // 写入日志文件
    const logFilePath = path.join(logDir, `server_${new Date().toISOString().split('T')[0]}.log`);
    fs.appendFileSync(logFilePath, logEntry + '\n');
}

// 设置文件上传
const upload = multer({ dest: 'uploads/' });

// 设置静态文件服务，指向当前目录
app.use(express.static('.'));

// 主页路由 - 直接返回 index.html 文件
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// 新增规范化表格页面路由
app.get('/normalize', (req, res) => {
  res.sendFile(path.join(__dirname, 'normalize_table.html'));
});

// 文件上传和处理API
app.post('/api/upload', upload.single('excelFile'), (req, res) => {
  try {
    if (!req.file) {
      logMessage('No file uploaded.');
      return res.status(400).json({ error: 'No file uploaded.' });
    }

    // 获取分页约束值，默认为33
    const maxConstraint = req.body.maxConstraint ? parseInt(req.body.maxConstraint) : 33;
    
    logMessage('Processing file: ' + req.file.path); // 添加调试日志
    
    // 读取上传的Excel文件
    const workbook = XLSX.readFile(req.file.path);
    logMessage('Available sheets: ' + JSON.stringify(workbook.SheetNames)); // 添加调试日志
    
    const sheetName = workbook.SheetNames[0];
    logMessage('Using sheet: ' + sheetName); // 添加调试日志
    
    const worksheet = workbook.Sheets[sheetName];
    
    // 转换为JSON格式
    let data = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
    
    logMessage('Data length: ' + data.length); // 添加调试日志
    logMessage('First row: ' + JSON.stringify(data[0])); // 添加调试日志
    
    // 检查数据是否为空 - 修正这里的检查逻辑
    if (!data || data.length === 0) {
      logMessage('Data is empty'); // 添加调试日志
      fs.unlinkSync(req.file.path);
      return res.status(400).json({ error: '文件中没有数据' });
    }
    
    // 获取原始列名
    const originalColumns = Object.keys(data[0] || {});
    logMessage('Original columns: ' + JSON.stringify(originalColumns)); // 添加调试日志
    
    // 清理日期字段
    const cleanedData = cleanDateFields(data);
    const cleanedColumns = Object.keys(cleanedData[0] || {});
    
    // 查找必要的列
    const finalLimitTimeCol = findColumn(cleanedData, '最终涨停时间');
    const continuousLimitDaysCol = findColumn(cleanedData, '连续涨停天数(天)');
    const limitReasonCol = findColumn(cleanedData, '涨停原因');
    const limitReasonCategoryCol = findColumn(cleanedData, '涨停原因类别');
    
    logMessage('Found columns: ' + JSON.stringify({ finalLimitTimeCol, continuousLimitDaysCol, limitReasonCol, limitReasonCategoryCol })); // 添加调试日志
    
    if (!finalLimitTimeCol || !continuousLimitDaysCol || !limitReasonCol || !limitReasonCategoryCol) {
      logMessage('Missing required columns'); // 添加调试日志
      // 清理上传的文件
      fs.unlinkSync(req.file.path);
      return res.status(400).json({ error: '缺少必要列，请检查文件格式。' });
    }
    
    // 重命名列
    const renamedData = cleanedData.map(row => {
      const newRow = { ...row };
      
      // 移除原字段以避免重复
      if (finalLimitTimeCol && newRow.hasOwnProperty(finalLimitTimeCol)) delete newRow[finalLimitTimeCol];
      if (continuousLimitDaysCol && newRow.hasOwnProperty(continuousLimitDaysCol)) delete newRow[continuousLimitDaysCol];
      if (limitReasonCol && newRow.hasOwnProperty(limitReasonCol)) delete newRow[limitReasonCol];
      if (limitReasonCategoryCol && newRow.hasOwnProperty(limitReasonCategoryCol)) delete newRow[limitReasonCategoryCol];
      
      // 添加标准化后的字段
      if (finalLimitTimeCol) newRow['最终涨停时间'] = row[finalLimitTimeCol];
      if (continuousLimitDaysCol) newRow['连续涨停天数(天)'] = row[continuousLimitDaysCol];
      if (limitReasonCol) newRow['涨停原因'] = row[limitReasonCol];
      if (limitReasonCategoryCol) newRow['涨停原因类别'] = row[limitReasonCategoryCol];
      
      // 移除'涨停原因揭秘'字段（支持模糊匹配）
      Object.keys(newRow).forEach(key => {
        if (key.includes('涨停原因揭秘')) {
          delete newRow[key];
        }
      });
      
      return newRow;
    });
    
    // 处理涨停原因类别字段
    const processedData = processReasonCategoryField(renamedData);
    
    // 按规则排序
    const sortedData = sortData(processedData);
    
    // 分页处理，使用传入的约束值
    const { pages, crossPageInfo } = splitIntoPagesByCategoryPriority(sortedData, '涨停原因', maxConstraint);
    
    // 准备响应数据
    const result = {
      originalColumns,
      cleanedColumns,
      finalLimitTimeCol,
      continuousLimitDaysCol,
      limitReasonCol,
      limitReasonCategoryCol,
      recordCount: data.length,
      pages: pages.map((page, index) => ({
        pageNumber: index + 1,
        recordCount: page.length,
        data: page.slice(0, Math.min(5, page.length)) // 只显示前5条记录用于预览
      })),
      categoryStats: getCategoryStats(sortedData),
      maxConstraint: maxConstraint, // 将约束值传递给前端
      crossPageInfo: crossPageInfo  // 添加跨页信息
    };
    
    // 清理上传的文件
    fs.unlinkSync(req.file.path);
    
    logMessage('Successfully processed file'); // 添加调试日志
    
    // 返回JSON格式的结果
    res.json(result);
  } catch (error) {
    logMessage('Error in /api/upload: ' + error.message);
    // 确保即使出错也清理上传的文件
    if (req.file) {
      try {
        fs.unlinkSync(req.file.path);
      } catch (e) {
        logMessage('Failed to clean up uploaded file: ' + e.message);
      }
    }
    res.status(500).json({ error: '处理文件时发生错误: ' + error.message });
  }
});

// 表格合并API
app.post('/api/merge', upload.fields([{ name: 'mainFile', maxCount: 1 }, { name: 'subFile', maxCount: 1 }]), (req, res) => {
  try {
    if (!req.files || !req.files.mainFile || !req.files.subFile) {
      logMessage('Need to upload both main table and sub table files.');
      return res.status(400).json({ error: '需要上传主表和子表文件。' });
    }

    const mainFile = req.files.mainFile[0];
    const subFile = req.files.subFile[0];

    logMessage('Merging files: main=' + mainFile.path + ', sub=' + subFile.path); // 添加日志

    // 读取主表文件
    const mainWorkbook = XLSX.readFile(mainFile.path);
    const mainSheetName = mainWorkbook.SheetNames[0]; // 只取第一个工作表
    const mainWorksheet = mainWorkbook.Sheets[mainSheetName];
    let mainData = XLSX.utils.sheet_to_json(mainWorksheet, { defval: '' });

    // 检查主表数据是否为空
    if (!mainData || mainData.length === 0) {
      logMessage('Main table has no data');
      fs.unlinkSync(mainFile.path);
      fs.unlinkSync(subFile.path);
      return res.status(400).json({ error: '主表中没有数据' });
    }

    // 读取子表文件 - 读取第二个工作表（索引为1）
    const subWorkbook = XLSX.readFile(subFile.path);
    if (subWorkbook.SheetNames.length < 2) {
      // 如果子表只有一个工作表，则使用第一个
      logMessage('Sub table must have at least two sheets');
      fs.unlinkSync(mainFile.path);
      fs.unlinkSync(subFile.path);
      return res.status(400).json({ error: '子表必须至少有两个工作表，合并数据应在第二个工作表中' });
    }
    
    const subSheetName = subWorkbook.SheetNames[1]; // 使用第二个工作表（索引为1）
    const subWorksheet = subWorkbook.Sheets[subSheetName];
    let subData = XLSX.utils.sheet_to_json(subWorksheet, { defval: '' });

    // 检查子表数据是否为空
    if (!subData || subData.length === 0) {
      logMessage('Second sheet of sub table has no data');
      fs.unlinkSync(mainFile.path);
      fs.unlinkSync(subFile.path);
      return res.status(400).json({ error: '子表第二个工作表中没有数据' });
    }

    // 规范化子表数据 - 只处理code、gtime、value三个字段
    const codeCol = findColumn(subData, 'code') || findColumnByPattern(subData, ['代码', 'code', 'id', '股票代码', 'stock']);
    const gtimeCol = findColumn(subData, 'gtime') || findColumnByPattern(subData, ['时间', 'gtime', 'date', 'datetime', 'gmt']);
    const valueCol = findColumn(subData, 'value') || findColumnByPattern(subData, ['值', 'value', '数值', '金额', 'price']);

    // 如果仍然找不到关键列，则尝试使用位置来识别列（第一列是code，第二列是gtime，第三列是value）
    let actualCodeCol, actualGtimeCol, actualValueCol;
    if (codeCol && gtimeCol && valueCol) {
      actualCodeCol = codeCol;
      actualGtimeCol = gtimeCol;
      actualValueCol = valueCol;
    } else {
      const columns = Object.keys(subData[0] || {});
      if (columns.length >= 3) {
        actualCodeCol = columns[0];
        actualGtimeCol = columns[1];
        actualValueCol = columns[2];
        logMessage('Using position to identify columns: ' + actualCodeCol + ', ' + actualGtimeCol + ', ' + actualValueCol);
      } else {
        // 清理上传的文件
        logMessage('Cannot find required columns in sub table');
        fs.unlinkSync(mainFile.path);
        fs.unlinkSync(subFile.path);
        return res.status(400).json({ error: '子表第二个工作表中未找到必需的列: code, gtime, value 或无法推断的列' });
      }
    }

    // 创建子表的映射，只使用三个关键字段
    const subTableMap = {};
    for (let i = 0; i < subData.length; i++) {
      let code = subData[i][actualCodeCol];
      // 提取前6位作为键
      let key = code ? String(code).substring(0, 6) : '';
      // 如果code包含点号，如"000657.SZ"，则提取前面的部分
      if (typeof code === 'string' && code.includes('.')) {
        key = code.split('.')[0];
      }
      subTableMap[key] = {
        gtime: subData[i][actualGtimeCol],
        value: subData[i][actualValueCol]
      };
    }

    // 合并数据
    const mergedData = [];
    for (let i = 0; i < mainData.length; i++) {
      const mainRow = mainData[i];
      const newRow = { ...mainRow };

      // 获取主表的股票代码前6位
      let stockCode = mainRow['股票代码'] || mainRow['证券代码'] || mainRow['代码'];
      if (stockCode) {
        // 提取前6位
        let key = String(stockCode).substring(0, 6);
        // 如果在SZ或SH后有.，则可能是"000657.SZ"格式，需要特殊处理
        if (typeof stockCode === 'string' && stockCode.includes('.')) {
          key = stockCode.split('.')[0];
        }

        // 如果在子表中找到匹配项，则合并数据
        if (subTableMap[key]) {
          newRow['子表涨停时间'] = subTableMap[key].gtime || '';
          newRow['子表value'] = subTableMap[key].value || '';  // 使用更新后的字段名
        }
      }

      mergedData.push(newRow);
    }

    // 获取分页约束值，默认为33
    const maxConstraint = req.body.maxConstraint ? parseInt(req.body.maxConstraint) : 33;
    
    // 对合并后的数据进行分页处理
    const { pages, crossPageInfo } = splitIntoPagesByCategoryPriority(mergedData, '涨停原因', maxConstraint);
    
    // 准备响应数据
    const result = {
      mainRecordCount: mainData.length,
      subRecordCount: subData.length,
      mergedRecordCount: mergedData.length,
      mergedDataPreview: mergedData.slice(0, Math.min(5, mergedData.length)), // 只显示前5条合并记录用于预览
      pages: pages.map((page, index) => ({
        pageNumber: index + 1,
        recordCount: page.length,
        data: page.slice(0, Math.min(5, page.length)) // 只显示前5条记录用于预览
      })),
      categoryStats: getCategoryStats(mergedData), // 使用合并后数据计算统计
      maxConstraint: maxConstraint, // 将约束值传递给前端
      crossPageInfo: crossPageInfo  // 添加跨页信息
    };

    // 清理上传的文件
    fs.unlinkSync(mainFile.path);
    fs.unlinkSync(subFile.path);

    logMessage('Successfully merged files');
    
    // 返回JSON格式的结果
    res.json(result);
  } catch (error) {
    logMessage('Error in /api/merge: ' + error.message);
    // 确保即使出错也清理上传的文件
    if (req.files && req.files.mainFile) {
      try {
        fs.unlinkSync(req.files.mainFile[0].path);
      } catch (e) {
        logMessage('Failed to clean up main file: ' + e.message);
      }
    }
    if (req.files && req.files.subFile) {
      try {
        fs.unlinkSync(req.files.subFile[0].path);
      } catch (e) {
        logMessage('Failed to clean up sub file: ' + e.message);
      }
    }
    res.status(500).json({ error: '合并文件时发生错误: ' + error.message });
  }
});

// 规范化表格API
app.post('/api/normalize', upload.single('excelFile'), (req, res) => {
  try {
    if (!req.file) {
      logMessage('No file uploaded for normalization.');
      return res.status(400).json({ error: 'No file uploaded.' });
    }

    logMessage('Normalizing file: ' + req.file.path);

    // 读取上传的Excel文件
    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // 转换为JSON格式
    let data = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
    
    // 获取原始列名
    const originalColumns = Object.keys(data[0] || {});
    
    // 查找关键列
    const codeCol = findColumn(data, 'code');
    const gtimeCol = findColumn(data, 'gtime');
    const valueCol = findColumn(data, 'value');
    
    if (!codeCol || !gtimeCol || !valueCol) {
      // 清理上传的文件
      fs.unlinkSync(req.file.path);
      logMessage('Required columns not found in normalization: code, gtime, value');
      return res.status(400).json({ error: '未找到必需的列: code, gtime, value' });
    }
    
    // 规范化数据
    const normalizedData = normalizeTable(data, codeCol, gtimeCol, valueCol);
    
    // 准备响应数据
    const result = {
      originalColumns,
      recordCount: data.length,
      normalizedRecordCount: normalizedData.length,
      columnCount: Object.keys(normalizedData[0] || {}).length,
      codeColumn: codeCol,
      gtimeColumn: gtimeCol,
      valueColumn: valueCol,
      originalDataPreview: data.slice(0, Math.min(5, data.length)), // 只显示前5条原始记录用于预览
      normalizedDataPreview: normalizedData.slice(0, Math.min(5, normalizedData.length)) // 只显示前5条规范化记录用于预览
    };
    
    // 清理上传的文件
    fs.unlinkSync(req.file.path);
    
    logMessage('Successfully normalized file');
    
    // 返回JSON格式的结果
    res.json(result);
  } catch (error) {
    logMessage('Error in /api/normalize: ' + error.message);
    // 确保即使出错也清理上传的文件
    if (req.file) {
      try {
        fs.unlinkSync(req.file.path);
      } catch (e) {
        logMessage('Failed to clean up uploaded file: ' + e.message);
      }
    }
    res.status(500).json({ error: '处理文件时发生错误: ' + error.message });
  }
});

// 导出规范化数据的API
app.post('/api/export-normalized', upload.single('excelFile'), (req, res) => {
  try {
    if (!req.file) {
      logMessage('No file uploaded for export normalization.');
      return res.status(400).json({ error: 'No file uploaded.' });
    }

    logMessage('Exporting normalized file: ' + req.file.path);

    // 读取上传的Excel文件
    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // 转换为JSON格式
    let data = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
    
    // 查找关键列
    const codeCol = findColumn(data, 'code');
    const gtimeCol = findColumn(data, 'gtime');
    const valueCol = findColumn(data, 'value');
    
    if (!codeCol || !gtimeCol || !valueCol) {
      // 清理上传的文件
      fs.unlinkSync(req.file.path);
      logMessage('Required columns not found for export normalization: code, gtime, value');
      return res.status(400).json({ error: '未找到必需的列: code, gtime, value' });
    }
    
    // 规范化数据
    const normalizedData = normalizeTable(data, codeCol, gtimeCol, valueCol);
    
    // 创建工作簿和工作表
    const worksheetOut = XLSX.utils.json_to_sheet(normalizedData);
    const workbookOut = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbookOut, worksheetOut, '规范化数据');
    
    // 生成二进制数据
    const buffer = XLSX.write(workbookOut, { bookType: 'xlsx', type: 'buffer' });
    
    // 清理上传的文件
    fs.unlinkSync(req.file.path);
    
    logMessage('Successfully exported normalized file');
    
    // 设置响应头并发送文件
    res.setHeader('Content-Disposition', 'attachment; filename="normalized_table.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
  } catch (error) {
    logMessage('Error in /api/export-normalized: ' + error.message);
    // 确保即使出错也清理上传的文件
    if (req.file) {
      try {
        fs.unlinkSync(req.file.path);
      } catch (e) {
        logMessage('Failed to clean up uploaded file: ' + e.message);
      }
    }
    res.status(500).json({ error: '处理文件时发生错误: ' + error.message });
  }
});

// 规范化表格数据
function normalizeTable(data, codeCol, gtimeCol, valueCol) {
  if (!data || data.length === 0) {
    return [];
  }
  
  // 创建规范化数据
  const normalizedResult = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const normalizedRow = {};
    
    // 复制关键数据到标准化列
    normalizedRow['股票代码'] = row[codeCol];
    normalizedRow['涨停时间'] = row[gtimeCol];
    normalizedRow['涨停原因'] = row[valueCol];
    
    // 添加其他可能的列
    for (const key in row) {
      if (row.hasOwnProperty(key)) {
        // 跳过已经处理过的列
        if (key !== codeCol && key !== gtimeCol && key !== valueCol) {
          normalizedRow[key] = row[key];
        }
      }
    }
    
    normalizedResult.push(normalizedRow);
  }
  
  return normalizedResult;
}

// 清理包含日期的字段名
function cleanDateFields(data) {
  if (data.length === 0) return data;
  
  const firstRow = data[0];
  const newHeaders = {};
  
  // 创建新的列名映射
  Object.keys(firstRow).forEach(col => {
    // 匹配类似 "字段名YYYY.MM.DD" 的模式
    const cleanedCol = col.replace(/\d{4}\.\d{2}\.\d{2}$/, '');
    newHeaders[col] = cleanedCol;
  });
  
  // 应用新的列名到所有行
  return data.map(row => {
    const newRow = {};
    Object.keys(row).forEach(oldKey => {
      const newKey = newHeaders[oldKey];
      newRow[newKey] = row[oldKey];
    });
    return newRow;
  });
}

// 查找列名，支持模糊匹配
function findColumn(data, columnName) {
  if (data.length === 0) return null;
  
  const firstRow = data[0];
  const columns = Object.keys(firstRow);
  
  // 精确匹配
  if (columns.includes(columnName)) {
    return columnName;
  }
  
  // 模糊匹配（去除空格后匹配）
  for (const col of columns) {
    if (col.trim().toLowerCase().includes(columnName.toLowerCase())) {
      return col;
    }
  }
  
  // 部分匹配
  for (const col of columns) {
    if (col.toLowerCase().includes(columnName.toLowerCase())) {
      return col;
    }
  }
  
  return null;
}

// 根据模式查找列
function findColumnByPattern(data, patterns) {
  if (data.length === 0) return null;

  const firstRow = data[0];
  const columns = Object.keys(firstRow);

  for (let i = 0; i < patterns.length; i++) {
    const pattern = patterns[i];
    for (let j = 0; j < columns.length; j++) {
      const col = columns[j];
      if (col.toLowerCase().includes(pattern.toLowerCase())) {
        return col;
      }
    }
  }

  return null;
}

// 计算文本长度，中文字符算2个长度，英文、数字和其他字符算1个长度
function calculateChineseLength(text) {
  if (text === null || text === undefined) {
    return 0;
  }
  
  text = String(text);
  // 统计中文字符数量
  const chineseChars = text.match(/[\u4e00-\u9fff]/g) || [];
  const chineseCount = chineseChars.length;
  
  // 总长度 = 中文字符数*2 + 其他字符数(英文、数字、标点等)
  const totalLength = chineseCount * 2 + (text.length - chineseCount);
  return totalLength;
}

// 标准化涨停原因类别字段
function normalizeReasonCategory(reasonCategory) {
  if (reasonCategory === null || reasonCategory === undefined) {
    return "";
  }
  // 去除首尾空格
  reasonCategory = String(reasonCategory).trim();
  // 去除多余的空格
  reasonCategory = reasonCategory.replace(/\s+/g, ' ');
  return reasonCategory;
}

// 确保涨停原因类别字段总长度不超过指定字符数
function trimReasonCategoryField(reasonCategory, maxLength = 36) {
  if (calculateChineseLength(reasonCategory) <= maxLength) {
    // 即使长度满足要求，也要检查末尾是否是"+"并移除
    let result = reasonCategory;
    while (result.endsWith('+')) {
      result = result.slice(0, -1);
    }
    return result;
  }
  
  // 从后向前逐步截断直到满足长度要求
  for (let i = reasonCategory.length; i > 0; i--) {
    let truncated = reasonCategory.slice(0, i);
    
    // 如果截断后长度满足要求
    if (calculateChineseLength(truncated) <= maxLength) {
      // 找到最后一个"+"的位置
      const lastPlusIndex = truncated.lastIndexOf('+');
      
      // 如果存在"+"且不在末尾，则在最后一个"+"处截断
      if (lastPlusIndex !== -1 && lastPlusIndex < truncated.length - 1) {
        truncated = truncated.slice(0, lastPlusIndex);
      }
      
      // 移除末尾的"+"字符
      while (truncated.endsWith('+')) {
        truncated = truncated.slice(0, -1);
      }
      
      return truncated;
    }
  }
  
  // 如果单个字符就超长了，返回空字符串
  return "";
}

// 处理涨停原因类别字段
function processReasonCategoryField(data) {
  return data.map(row => {
    const newRow = { ...row };
    try {
      const normalized = normalizeReasonCategory(newRow['涨停原因类别']);
      newRow['涨停原因类别'] = trimReasonCategoryField(normalized);
    } catch (error) {
      logMessage("Error processing reason category field: " + error.message);
      newRow['涨停原因类别'] = "";
    }
    return newRow;
  });
}

// 排序数据
function sortData(data) {
  return data.sort((a, b) => {
    // 首先按连续涨停天数(天)降序排序（天数多的在前）
    const daysDiff = b['连续涨停天数(天)'] - a['连续涨停天数(天)'];
    if (daysDiff !== 0) {
      return daysDiff;
    }
    
    // 然后按最终涨停时间升序排序（时间早的在前）
    if (a['最终涨停时间'] < b['最终涨停时间']) return -1;
    if (a['最终涨停时间'] > b['最终涨停时间']) return 1;
    return 0;
  });
}

// 根据分类数和条目数的关系进行分页
function splitIntoPagesByCategoryPriority(data, categoryCol, maxConstraint = 33) {
  // 统计各类别的出现次数
  const categoryCounts = {};
  data.forEach(row => {
    const category = row[categoryCol];
    categoryCounts[category] = (categoryCounts[category] || 0) + 1;
  });
  
  // 按出现次数降序排列，但将"其他概念"放在最后
  const sortedCategories = Object.entries(categoryCounts)
    .filter(([cat]) => cat !== "其他概念")
    .sort((a, b) => b[1] - a[1])
    .map(([cat]) => cat);
    
  if (categoryCounts["其他概念"]) {
    sortedCategories.push("其他概念");
  }
  
  // 按照类别优先级重新排列数据
  const reorderedData = [];
  const otherConceptData = [];
  
  sortedCategories.forEach(cat => {
    const catData = data.filter(row => row[categoryCol] === cat);
    if (cat === "其他概念") {
      otherConceptData.push(...catData);
    } else {
      reorderedData.push(...catData);
    }
  });
  
  reorderedData.push(...otherConceptData);
  
  // 按约束条件进行分页，并检测跨页情况
  const pages = [];
  const crossPageInfo = []; // 存储跨页信息
  
  // 跟踪每个类别出现在哪些页面中
  const categoryPageMap = {};
  
  let i = 0;
  
  while (i < reorderedData.length) {
    let j = i;
    let categoryCount = 0;
    let itemCount = 0;
    const categories = new Set();
    
    while (j < reorderedData.length) {
      const currentCategory = reorderedData[j][categoryCol];
      if (!categories.has(currentCategory)) {
        categories.add(currentCategory);
        categoryCount++;
      }
      
      itemCount++;
      
      // 检查是否满足约束条件
      if (categoryCount * 2 + itemCount > maxConstraint) {
        // 如果加入这条记录会超出限制，则不包含这条记录
        break;
      }
      
      j++;
    }
    
    // 如果没有满足条件的记录（可能第一条就不满足），至少保留一条
    if (j === i) {
      j = i + 1;
    }
    
    // 记录当前页面的索引（从1开始）
    const currentPageIndex = pages.length + 1;
    
    // 添加调试信息
    logMessage('Processing page ' + currentPageIndex + ', data range: ' + i + ' to ' + (j-1));
    
    // 显示当前页面包含的所有类别
    const currentPageCategories = {};
    for (let k = i; k < j; k++) {
      const category = reorderedData[k][categoryCol];
      if (!currentPageCategories.hasOwnProperty(category)) {
        currentPageCategories[category] = 0;
      }
      currentPageCategories[category]++;
    }
    logMessage('Page ' + currentPageIndex + ' categories: ' + JSON.stringify(currentPageCategories));
    
    // 先检查当前页面中的类别是否已在之前的页面中出现
    for (let k = i; k < j; k++) {
      const category = reorderedData[k][categoryCol];
      
      // 如果这个类别之前出现过，记录跨页信息
      if (categoryPageMap.hasOwnProperty(category)) {
        // 检查是否是相邻页面，避免重复记录
        if (!categoryPageMap[category].includes(currentPageIndex)) {
          crossPageInfo.push({
            category: category,
            fromPage: Math.max(...categoryPageMap[category]),
            toPage: currentPageIndex
          });
          
          // 调试输出
          logMessage('Cross-page detected: ' + category + ' from page ' + Math.max(...categoryPageMap[category]) + ' to page ' + currentPageIndex);
        }
      }
    }
    
    // 然后更新类别页面映射
    for (let k = i; k < j; k++) {
      const category = reorderedData[k][categoryCol];
      
      if (!categoryPageMap.hasOwnProperty(category)) {
        categoryPageMap[category] = [];
      }
      categoryPageMap[category].push(currentPageIndex);
    }
    
    pages.push(reorderedData.slice(i, j));
    i = j;
  }
  
  // 将跨页信息添加到结果中
  return { pages, crossPageInfo };
}

// 获取类别统计数据
function getCategoryStats(data) {
  const categoryCounts = {};
  data.forEach(row => {
    const category = row['涨停原因'];
    categoryCounts[category] = (categoryCounts[category] || 0) + 1;
  });
  
  return Object.entries(categoryCounts)
    .sort((a, b) => b[1] - a[1])
    .map(([category, count]) => ({ category, count }));
}

// 启动服务器
app.listen(port,() => {
  logMessage(`Server running at http://0.0.0.0:${port}`);
});