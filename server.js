const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const port = 3001;

// 设置文件上传
const upload = multer({ dest: 'uploads/' });

// 设置静态文件服务，指向当前目录
app.use(express.static('.'));

// 主页路由 - 直接返回 index.html 文件
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// 文件上传和处理API
app.post('/api/upload', upload.single('excelFile'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded.' });
    }

    // 获取分页约束值，默认为33
    const maxConstraint = req.body.maxConstraint ? parseInt(req.body.maxConstraint) : 33;
    
    // 读取上传的Excel文件
    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // 转换为JSON格式
    let data = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
    
    // 获取原始列名
    const originalColumns = Object.keys(data[0] || {});
    
    // 清理日期字段
    const cleanedData = cleanDateFields(data);
    const cleanedColumns = Object.keys(cleanedData[0] || {});
    
    // 查找必要的列
    const finalLimitTimeCol = findColumn(cleanedData, '最终涨停时间');
    const continuousLimitDaysCol = findColumn(cleanedData, '连续涨停天数(天)');
    const limitReasonCol = findColumn(cleanedData, '涨停原因');
    const limitReasonCategoryCol = findColumn(cleanedData, '涨停原因类别');
    
    if (!finalLimitTimeCol || !continuousLimitDaysCol || !limitReasonCol || !limitReasonCategoryCol) {
      // 清理上传的文件
      fs.unlinkSync(req.file.path);
      return res.status(400).json({ error: '缺少必要列，请检查文件格式。' });
    }
    
    // 重命名列
    const renamedData = cleanedData.map(row => {
      const newRow = { ...row };
      
      // 移除原字段以避免重复
      delete newRow[finalLimitTimeCol];
      delete newRow[continuousLimitDaysCol];
      delete newRow[limitReasonCol];
      delete newRow[limitReasonCategoryCol];
      
      // 添加标准化后的字段
      newRow['最终涨停时间'] = row[finalLimitTimeCol];
      newRow['连续涨停天数(天)'] = row[continuousLimitDaysCol];
      newRow['涨停原因'] = row[limitReasonCol];
      newRow['涨停原因类别'] = row[limitReasonCategoryCol];
      
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
    const pages = splitIntoPagesByCategoryPriority(sortedData, '涨停原因', maxConstraint);
    
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
      maxConstraint: maxConstraint // 将约束值传递给前端
    };
    
    // 清理上传的文件
    fs.unlinkSync(req.file.path);
    
    // 返回JSON格式的结果
    res.json(result);
  } catch (error) {
    console.error(error);
    // 确保即使出错也清理上传的文件
    if (req.file) {
      try {
        fs.unlinkSync(req.file.path);
      } catch (e) {
        console.error('Failed to clean up uploaded file:', e);
      }
    }
    res.status(500).json({ error: '处理文件时发生错误: ' + error.message });
  }
});

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
    if (col.trim() === columnName) {
      return col;
    }
  }
  
  // 部分匹配
  for (const col of columns) {
    if (col.includes(columnName)) {
      return col;
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
      console.error("处理涨停原因类别字段时出错:", error);
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
  
  // 按约束条件进行分页
  const pages = [];
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
    
    pages.push(reorderedData.slice(i, j));
    i = j;
  }
  
  return pages;
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
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});