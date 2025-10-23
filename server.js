// 加载环境变量配置（必须在最开头）
require('dotenv').config();

const express = require('express');
const ExcelJS = require('exceljs');
const { execSync } = require('child_process');
const { startOfWeek, endOfWeek, format, parseISO } = require('date-fns');
const OpenAI = require('openai');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');

const app = express();
const PORT = 3000;

// 中间件
app.use(express.json());
app.use(express.static('public'));

// ==================== 邮件服务配置 ====================
/**
 * 创建邮件传输器
 */
function createMailTransporter() {
  const transporter = nodemailer.createTransporter({
    host: process.env.SMTP_HOST,
    port: parseInt(process.env.SMTP_PORT) || 587,
    secure: process.env.SMTP_SECURE === 'true', // true for 465, false for other ports
    auth: {
      user: process.env.SMTP_USER,
      pass: process.env.SMTP_PASS
    }
  });
  
  return transporter;
}

/**
 * 发送邮件
 * @param {string} to - 收件人邮箱
 * @param {string} cc - 抄送邮箱（可选）
 * @param {string} subject - 邮件主题
 * @param {string} html - 邮件内容（HTML格式）
 * @param {string} attachmentPath - 附件路径
 * @param {string} attachmentName - 附件名称
 */
async function sendEmail(to, cc, subject, html, attachmentPath, attachmentName) {
  try {
    const transporter = createMailTransporter();
    
    const mailOptions = {
      from: {
        name: process.env.MAIL_FROM_NAME || '周报生成器',
        address: process.env.MAIL_FROM_EMAIL || process.env.SMTP_USER
      },
      to: to,
      cc: cc,
      subject: subject,
      html: html,
      attachments: [
        {
          filename: attachmentName,
          path: attachmentPath,
          contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
      ]
    };
    
    console.log(`📧 正在发送邮件...`);
    console.log(`   收件人: ${to}`);
    console.log(`   抄送: ${cc || '无'}`);
    console.log(`   主题: ${subject}`);
    console.log(`   附件: ${attachmentName}`);
    
    const result = await transporter.sendMail(mailOptions);
    console.log(`✅ 邮件发送成功！消息ID: ${result.messageId}`);
    
    return { success: true, messageId: result.messageId };
  } catch (error) {
    console.error(`❌ 邮件发送失败:`, error.message);
    return { success: false, error: error.message };
  }
}

// 初始化 DeepSeek 客户端
const openai = new OpenAI({
  baseURL: 'https://api.deepseek.com',
  apiKey: process.env.DEEPSEEK_API_KEY
});

// 配置文件路径
const CONFIG_FILE = './config.json';

// 读取配置文件
function loadConfig() {
  try {
    if (fs.existsSync(CONFIG_FILE)) {
      const data = fs.readFileSync(CONFIG_FILE, 'utf8');
      return JSON.parse(data);
    }
  } catch (err) {
    console.error('❌ 读取配置文件失败:', err.message);
  }
  return {
    userName: "",
    projectPaths: [],
    lastUsed: null,
    settings: {
      autoSave: true,
      defaultDateRange: "currentWeek",
      maxProjects: 10
    }
  };
}

// 保存配置文件
function saveConfig(config) {
  try {
    config.lastUsed = new Date().toISOString();
    fs.writeFileSync(CONFIG_FILE, JSON.stringify(config, null, 2));
    return true;
  } catch (err) {
    console.error('❌ 保存配置文件失败:', err.message);
    return false;
  }
}

// ==================== 工具函数 ====================

/**
 * 获取指定日期范围
 * 默认为本周周一到周五（工作日）
 */
function getWeekRange(startDate, endDate) {
  let start, end;
  
  if (startDate && endDate) {
    // 如果提供了日期，使用提供的日期
    start = parseISO(startDate);
    end = parseISO(endDate);
  } else {
    // 默认：本周周一到周五
    const today = new Date();
    start = startOfWeek(today, { weekStartsOn: 1 }); // 周一
    
    // 周五 = 周一 + 4天
    end = new Date(start);
    end.setDate(start.getDate() + 4);
  }

  return {
    start,
    end,
    startStr: format(start, 'MM月dd日'),
    endStr: format(end, 'MM月dd日'),
    year: format(start, 'yyyy'),
    month: format(start, 'MM')
  };
}

/**
 * 从单个Git仓库获取提交记录
 */
function getGitCommitsFromRepo(projectPath, since, until) {
  try {
    if (!fs.existsSync(projectPath)) {
      console.error(`❌ 项目路径不存在: ${projectPath}`);
      return [];
    }

    const cmd = `git -C "${projectPath}" log \
      --since="${since}" --until="${until} 23:59:59" \
      --pretty=format:"COMMIT_SEP|%H|%an|%ad|%s" --date=short \
      --name-status`;

    const output = execSync(cmd, { encoding: 'utf-8' });
    
    if (!output.trim()) {
      console.log(`  ℹ️  项目 [${path.basename(projectPath)}] 在此期间无提交记录`);
      return [];
    }
    
    const lines = output.split('\n').filter(line => line.trim() !== '');

    const commits = [];
    let currentCommit = null;

    for (const line of lines) {
      if (line.startsWith('COMMIT_SEP|')) {
        if (currentCommit) commits.push(currentCommit);
        const [, hash, author, date, message] = line.split('|');
        currentCommit = {
          hash: hash.substring(0, 8), // 只保留前8位
          author,
          date,
          message: message.trim(),
          files: [],
          project: path.basename(projectPath)
        };
      } else if (currentCommit) {
        currentCommit.files.push(line.trim());
      }
    }
    if (currentCommit) commits.push(currentCommit);

    return commits;
  } catch (err) {
    console.error(`❌ 获取Git提交记录失败 (${projectPath}):`, err.message);
    return [];
  }
}

/**
 * 从多个Git仓库获取提交记录
 */
function getGitCommits(projectPaths, startDate, endDate) {
  const { start, end, startStr, endStr } = getWeekRange(startDate, endDate);
  const since = format(start, 'yyyy-MM-dd');
  const until = format(end, 'yyyy-MM-dd');

  console.log(`\n📅 查询时间范围: ${since} ~ ${until} (${startStr} ~ ${endStr})`);
  console.log(`📁 扫描项目数量: ${projectPaths.length}\n`);

  let allCommits = [];
  
  for (const projectPath of projectPaths) {
    const projectName = path.basename(projectPath);
    console.log(`🔍 正在扫描项目: ${projectName}`);
    const commits = getGitCommitsFromRepo(projectPath, since, until);
    
    if (commits.length > 0) {
      console.log(`  ✅ 找到 ${commits.length} 条提交记录\n`);
      
      // 输出每条提交的详细信息
      commits.forEach((commit, index) => {
        console.log(`  📝 提交 ${index + 1}/${commits.length}:`);
        console.log(`     提交哈希: ${commit.hash}`);
        console.log(`     提交作者: ${commit.author}`);
        console.log(`     提交日期: ${commit.date}`);
        console.log(`     提交信息: ${commit.message}`);
        if (commit.files.length > 0) {
          console.log(`     修改文件: ${commit.files.slice(0, 3).join(', ')}${commit.files.length > 3 ? '...' : ''}`);
        }
        console.log('');
      });
    }
    
    allCommits = allCommits.concat(commits);
  }

  console.log(`\n✅ 总计获取 ${allCommits.length} 条提交记录（来自 ${projectPaths.length} 个项目）\n`);
  return allCommits;
}

// ==================== 工具函数：分组和聚合 ====================
/**
 * 按项目分组提交记录
 */
function groupCommitsByProject(commits) {
  const grouped = {};
  commits.forEach(commit => {
    if (!grouped[commit.project]) {
      grouped[commit.project] = [];
    }
    grouped[commit.project].push(commit);
  });
  return grouped;
}

/**
 * 智能分析项目的所有提交，按模块聚合并生成周报条目
 */
async function analyzeProjectCommits(projectName, commits) {
  console.log(`🤖 [${projectName}] 正在分析 ${commits.length} 条提交记录...`);
  
  // 构建提交信息摘要，包含文件路径用于模块识别
  const commitSummary = commits.map((commit, index) => {
    const fileList = commit.files.slice(0, 5).join(', ');
    const moreFiles = commit.files.length > 5 ? ` 等${commit.files.length}个文件` : '';
    return `${index + 1}. [${commit.date}] ${commit.message}\n   修改文件: ${fileList}${moreFiles}`;
  }).join('\n\n');

  const prompt = `你是一个专业的技术周报生成助手。请分析以下项目的 Git 提交记录，智能识别代码模块和功能，将相关提交聚合成高质量的周报条目。

项目名称: ${projectName}
提交记录（共 ${commits.length} 条）:

${commitSummary}

分析要求:
1. **模块识别**: 根据文件路径和提交信息，识别代码模块（如：用户模块、订单模块、支付模块等）
2. **功能聚合**: 将同一模块或功能的多次提交合并为一个任务
3. **工作描述**: 用专业、简洁的语言描述工作内容，避免过于技术化的细节且 让领导看到做了很多任务 而且同事看了任务很难实现
4. **关键改动**: 总结该任务的主要改动点（2-4个要点）
5. **提交编号**: 记录该任务涉及的提交编号（用于确定实际开始和结束时间）

输出格式（必须是有效的 JSON 数组）:
[
  {
    "模块": "模块或功能名称",
    "分类": "开发新功能|修复bug|优化性能|代码重构|文档更新",
    "描述": "简洁专业的工作描述（15-40字）",
    "关键改动": ["改动点1", "改动点2", "改动点3"],
    "涉及提交编号": [1, 2, 3]
  }
]

注意事项:
- 如果多个提交属于同一功能开发，请合并为一条
- 如果提交之间完全无关，可以分成多条
- 描述要站在周报汇报的角度，突出工作价值
- 避免使用"修复了一个bug"这样的模糊描述，要具体说明修复了什么问题
- **重要**: 必须包含"涉及提交编号"字段，记录该任务对应的提交编号（从1开始）

请直接输出 JSON 数组，不要有其他内容。`;

  try {
    const startTime = Date.now();
    
    const completion = await openai.chat.completions.create({
      model: 'deepseek-chat',
      messages: [{ role: 'user', content: prompt }],
      temperature: 0.3,
      max_tokens: 2000
    });

    const result = completion.choices[0].message.content.trim();
    
    // 尝试解析 JSON
    let parsedTasks;
    try {
      const jsonContent = result.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
      parsedTasks = JSON.parse(jsonContent);
    } catch (parseError) {
      console.error(`   ❌ JSON 解析失败，原始内容:\n${result}`);
      throw parseError;
    }
    
    const duration = Date.now() - startTime;
    console.log(`   ✅ AI 分析完成 (耗时: ${duration}ms)`);
    console.log(`   📊 识别出 ${parsedTasks.length} 个任务模块\n`);
    
    // 显示识别的模块
    parsedTasks.forEach((task, index) => {
      const commitInfo = task.涉及提交编号 && task.涉及提交编号.length > 0
        ? `提交#${task.涉及提交编号.join(', #')}`
        : `${task.涉及提交数 || 0}个提交`;
      console.log(`   ${index + 1}. [${task.模块}] ${task.描述} (${commitInfo})`);
    });
    
    return parsedTasks;
  } catch (err) {
    console.error(`   ❌ DeepSeek API 调用失败:`, err.message);
    
    // 降级方案
    console.log(`   ⚠️  使用降级方案: 按日期简单分组\n`);
    return [{
      模块: '未分类',
      分类: '开发任务',
      描述: `${projectName} 项目开发工作（${commits.length}个提交）`,
      关键改动: commits.slice(0, 3).map(c => c.message),
      涉及提交数: commits.length,
      涉及提交编号: Array.from({ length: commits.length }, (_, i) => i + 1) // 所有提交编号
    }];
  }
}

/**
 * 调用DeepSeek API解析单个提交信息（旧方法，保留作为备用）
 */
async function parseCommitWithDeepSeek(commitMessage, projectName) {
  console.log(`🤖 调用 DeepSeek AI 解析: [${projectName}] ${commitMessage.substring(0, 50)}...`);
  
  const prompt = `请严格按照以下要求解析代码提交信息：
  1. 输出格式：必须是JSON字符串，无其他多余内容
  2. 字段说明：
     - 类型："任务"或"问题"（修复bug、解决异常属于"问题"；开发新功能、优化代码属于"任务"）
     - 分类：任务/问题的具体分类（例如：开发新功能、修复生产bug、优化性能、文档更新等）
     - 描述：简化为10-30字的具体工作内容（去除冗余词汇）
     - 关联ID：提取需求号/BUG号（如#123则为"123"，无则为"无"）
  
  提交信息：${commitMessage}
  示例输出：{"类型": "任务", "分类": "开发新功能", "描述": "实现用户登录页验证码功能", "关联ID": "REQ-456"}`;

  try {
    const startTime = Date.now();
    
    const completion = await openai.chat.completions.create({
      model: 'deepseek-chat',
      messages: [{ role: 'user', content: prompt }],
      temperature: 0.1,
      max_tokens: 200
    });

    const result = completion.choices[0].message.content.trim();
    const parsed = JSON.parse(result);
    
    const duration = Date.now() - startTime;
    console.log(`   ✅ AI 解析完成 (耗时: ${duration}ms) -> ${parsed.描述}`);
    
    return parsed;
  } catch (err) {
    console.error(`   ❌ DeepSeek API 调用失败（${projectName}）:`, err.message);
    const fallback = {
      类型: '任务',
      分类: '未分类',
      描述: commitMessage.substring(0, 50),
      关联ID: '无'
    };
    
    console.log(`   ⚠️  使用降级方案: ${fallback.描述}`);
    return fallback;
  }
}

/**
 * 处理提交记录为周报数据（智能模块聚合版本）
 */
async function processCommits(commits, userName) {
  const tasks = [];
  const problems = []; // 保持空白，不填充任何内容

  console.log(`\n${'='.repeat(70)}`);
  console.log(`📊 智能分析模式：按项目分组，识别模块，聚合相关提交`);
  console.log(`${'='.repeat(70)}\n`);
  console.log(`📦 总提交数: ${commits.length} 条`);
  
  // 按项目分组
  const groupedCommits = groupCommitsByProject(commits);
  const projectNames = Object.keys(groupedCommits);
  console.log(`🗂️  涉及项目: ${projectNames.length} 个 (${projectNames.join(', ')})\n`);

  let taskNumber = 1;
  
  // 逐个项目进行智能分析
  for (const [projectName, projectCommits] of Object.entries(groupedCommits)) {
    console.log(`${'─'.repeat(70)}`);
    console.log(`📁 项目: ${projectName} (${projectCommits.length} 个提交)`);
    console.log(`${'─'.repeat(70)}\n`);
    
    // 调用 AI 智能分析该项目的所有提交
    const projectTasks = await analyzeProjectCommits(projectName, projectCommits);
    
    // 将分析结果转换为周报格式
    for (const task of projectTasks) {
      // 根据涉及的提交编号计算实际的日期范围
      let startDate, endDate;
      
      if (task.涉及提交编号 && Array.isArray(task.涉及提交编号) && task.涉及提交编号.length > 0) {
        // 获取该任务涉及的所有commits
        const taskCommits = task.涉及提交编号
          .map(index => projectCommits[index - 1]) // 编号从1开始，数组从0开始
          .filter(commit => commit !== undefined);
        
        if (taskCommits.length > 0) {
          // 从这些commits中提取日期并排序
          const taskDates = taskCommits.map(c => c.date).sort();
          startDate = taskDates[0];
          endDate = taskDates[taskDates.length - 1];
          
          console.log(`   📅 任务[${task.模块}] 时间范围: ${startDate} ~ ${endDate} (基于${taskCommits.length}个提交)`);
        } else {
          // 如果提交编号无效，使用整个项目的日期范围
          const dates = projectCommits.map(c => c.date).sort();
          startDate = dates[0];
          endDate = dates[dates.length - 1];
          console.log(`   ⚠️  任务[${task.模块}] 提交编号无效，使用项目整体时间范围`);
        }
      } else {
        // 如果没有提交编号信息，使用整个项目的日期范围
        const dates = projectCommits.map(c => c.date).sort();
        startDate = dates[0];
        endDate = dates[dates.length - 1];
        console.log(`   ⚠️  任务[${task.模块}] 缺少提交编号，使用项目整体时间范围`);
      }
      
      // 构建详细的事项说明
      const taskDescription = task.关键改动 && task.关键改动.length > 0
        ? `${task.描述}\n关键改动:\n${task.关键改动.map(item => `• ${item}`).join('\n')}`
        : task.描述;
      
      tasks.push({
        序号: taskNumber++,
        重点需求或任务: `[${projectName}] ${task.模块}`,
        事项说明: taskDescription,
        启动日期: startDate,
        预计完成日期: endDate,
        负责人: userName,
        协同人或部门: '无',
        完成进度: '100%',
        备注: ''
      });
    }
    
    console.log('');
  }

  console.log(`${'='.repeat(70)}`);
  console.log(`✅ 分析完成！`);
  console.log(`   📝 原始提交: ${commits.length} 条`);
  console.log(`   📊 生成任务: ${tasks.length} 条`);
  console.log(`   🎯 聚合率: ${((1 - tasks.length / commits.length) * 100).toFixed(1)}%`);
  console.log(`${'='.repeat(70)}\n`);

  return { tasks, problems };
}

/**
 * 智能计算单元格行高
 * @param {string} text - 单元格文本内容
 * @param {number} columnWidth - 列宽（字符数）
 * @param {number} fontSize - 字体大小
 * @returns {number} 建议的行高
 */
function calculateRowHeight(text, columnWidth = 40, fontSize = 11) {
  if (!text) return 20; // 默认行高
  
  const textStr = String(text);
  const lines = textStr.split('\n'); // 按换行符分割
  let totalLines = 0;
  
  for (const line of lines) {
    if (line.trim() === '') {
      totalLines += 1; // 空行也算一行
    } else {
      // 计算该行在指定列宽下会占用多少行
      // 中文字符按2个字符计算，英文和数字按1个字符计算
      const chineseChars = (line.match(/[\u4e00-\u9fa5]/g) || []).length;
      const otherChars = line.length - chineseChars;
      const effectiveLength = chineseChars * 2 + otherChars;
      
      const wrappedLines = Math.ceil(effectiveLength / columnWidth);
      totalLines += Math.max(1, wrappedLines);
    }
  }
  
  // 根据行数计算高度：基础高度 + (行数 × 行高系数)
  const baseHeight = 20;
  const lineHeightFactor = fontSize * 1.5; // 行高系数
  const calculatedHeight = baseHeight + (totalLines * lineHeightFactor);
  
  // 限制最小和最大高度
  return Math.max(30, Math.min(calculatedHeight, 300));
}

/**
 * 生成Excel周报
 */
async function generateExcel(userName, tasks, problems, startDate, endDate, outputPath) {
  const templatePath = './周报模版.xlsx';
  
  if (!fs.existsSync(templatePath)) {
    throw new Error(`❌ 模板文件不存在：${templatePath}`);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);
  const worksheet = workbook.getWorksheet(1);

  // 更新周报标题（合并C到F列）
  const { year, startStr, endStr } = getWeekRange(startDate, endDate);
  const title = `${userName} ${year}年${startStr}-${endStr}工作周报`;
  worksheet.getCell('C1').value = title;

  // 设置"事项说明"列（C列）的宽度
  worksheet.getColumn(3).width = 60; // 设置为60个字符宽度，可根据需要调整

  // 填充重点任务表格 (从A4开始，动态扩展)
  const taskStartRow = 4;
  const taskTemplateRows = 4; // 模板预留的任务行数（4-7行）
  
  tasks.forEach((task, index) => {
    const rowNum = taskStartRow + index;
    
    // 如果超过模板预留的行数，需要插入新行
    if (index >= taskTemplateRows) {
      // 复制上一行的样式作为模板
      const templateRow = worksheet.getRow(taskStartRow + taskTemplateRows - 1);
      worksheet.insertRow(rowNum, []);
      const newRow = worksheet.getRow(rowNum);
      
      // 复制样式
      for (let j = 1; j <= 9; j++) {
        const sourceCell = templateRow.getCell(j);
        const targetCell = newRow.getCell(j);
        
        // 复制样式属性
        if (sourceCell.style) {
          targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
        }
        // 确保有边框和背景
        targetCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFFFFF' }
        };
        targetCell.border = {
          top: { style: 'thin', color: { argb: 'FF000000' } },
          left: { style: 'thin', color: { argb: 'FF000000' } },
          bottom: { style: 'thin', color: { argb: 'FF000000' } },
          right: { style: 'thin', color: { argb: 'FF000000' } }
        };
        
        // 设置对齐方式（与主循环保持一致）
        if (j === 3) {
          targetCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true, indent: 1 };
        } else if (j === 1 || (j >= 4 && j <= 8)) {
          targetCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        } else {
          targetCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
        }
      }
    }
    
    const row = worksheet.getRow(rowNum);

    // 设置数据并支持换行
    row.getCell(1).value = task.序号;
    row.getCell(2).value = task.重点需求或任务;
    row.getCell(3).value = task.事项说明;
    row.getCell(4).value = task.启动日期;
    row.getCell(5).value = task.预计完成日期;
    row.getCell(6).value = task.负责人;
    row.getCell(7).value = task.协同人或部门;
    row.getCell(8).value = task.完成进度;
    row.getCell(9).value = task.备注;

    // 设置样式
    for (let j = 1; j <= 9; j++) {
      const cell = row.getCell(j);
      
      // 确保有边框和背景
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFFFF' }
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
      };
      
      // 设置不同列的对齐方式
      if (j === 3) {
        // 事项说明列：左对齐，顶部对齐
        cell.alignment = { 
          horizontal: 'left', 
          vertical: 'top', 
          wrapText: true,
          indent: 1
        };
      } else if (j === 1 || (j >= 4 && j <= 8)) {
        // 序号、启动日期、预计完成日期、负责人、协同人/部门、完成进度：居中对齐
        cell.alignment = { 
          horizontal: 'center', 
          vertical: 'middle', 
          wrapText: true 
        };
      } else {
        // 其他列（重点需求/任务、备注）：左对齐，顶部对齐
        cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
      }
    }
    
    // 根据事项说明内容智能设置行高
    const calculatedHeight = calculateRowHeight(task.事项说明, 55, 11);
    row.height = calculatedHeight;

    row.commit(); // 提交行修改
  });
  console.log(`✅ 已填充 ${tasks.length} 条重点任务`);

  // 填充日常问题表格 (从第15行开始，动态扩展)
  // 注意：由于可能插入了任务行，问题表格的起始行需要动态计算
  const problemStartRowBase = 15;
  const insertedTaskRows = Math.max(0, tasks.length - taskTemplateRows);
  const problemStartRow = problemStartRowBase + insertedTaskRows;
  const problemTemplateRows = 5; // 模板预留的问题行数（15-19行）
  
  problems.forEach((problem, index) => {
    const rowNum = problemStartRow + index;
    
    // 如果超过模板预留的行数，需要插入新行
    if (index >= problemTemplateRows) {
      // 复制上一行的样式作为模板
      const templateRow = worksheet.getRow(problemStartRow + problemTemplateRows - 1);
      worksheet.insertRow(rowNum, []);
      const newRow = worksheet.getRow(rowNum);
      
      // 复制样式
      for (let j = 1; j <= 6; j++) {
        const sourceCell = templateRow.getCell(j);
        const targetCell = newRow.getCell(j);
        
        // 复制样式属性
        if (sourceCell.style) {
          targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
        }
        // 确保有边框和背景
        targetCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFFFFF' }
        };
        targetCell.border = {
          top: { style: 'thin', color: { argb: 'FF000000' } },
          left: { style: 'thin', color: { argb: 'FF000000' } },
          bottom: { style: 'thin', color: { argb: 'FF000000' } },
          right: { style: 'thin', color: { argb: 'FF000000' } }
        };
        
        // 设置对齐方式（与主循环保持一致）
        if (j === 3 || j === 5) {
          targetCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true, indent: 1 };
        } else if (j === 1 || j === 2 || j === 4 || j === 6) {
          targetCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        } else {
          targetCell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
        }
      }
    }
    
    const row = worksheet.getRow(rowNum);
    
    // 设置数据并支持换行
    row.getCell(1).value = problem.序号;
    row.getCell(2).value = problem.问题分类;
    row.getCell(3).value = problem.具体描述;
    row.getCell(4).value = problem.提出日期;
    row.getCell(5).value = problem.解决方案;
    row.getCell(6).value = problem.解决日期;

    // 设置样式
    for (let j = 1; j <= 6; j++) {
      const cell = row.getCell(j);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFFFF' }
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
      };
      
      // 设置不同列的对齐方式
      if (j === 3 || j === 5) {
        // 具体描述、解决方案：左对齐，顶部对齐
        cell.alignment = { 
          horizontal: 'left', 
          vertical: 'top', 
          wrapText: true,
          indent: 1
        };
      } else if (j === 1 || j === 2 || j === 4 || j === 6) {
        // 序号、问题分类、提出日期、解决日期：居中对齐
        cell.alignment = { 
          horizontal: 'center', 
          vertical: 'middle', 
          wrapText: true 
        };
      } else {
        // 其他列：左对齐，顶部对齐
        cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
      }
    }
    
    // 根据具体描述和解决方案的内容智能设置行高
    const descHeight = calculateRowHeight(problem.具体描述, 55, 11);
    const solutionHeight = calculateRowHeight(problem.解决方案, 55, 11);
    row.height = Math.max(descHeight, solutionHeight);
    
    row.commit(); // 提交行修改
  });
  console.log(`✅ 已填充 ${problems.length} 条日常问题`);

  await workbook.xlsx.writeFile(outputPath);
  console.log(`🎉 周报生成成功！路径：${outputPath}`);
}

// ==================== API路由 ====================

/**
 * 获取配置API
 */
app.get('/api/config', (req, res) => {
  try {
    const config = loadConfig();
    res.json({ success: true, config });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

/**
 * 保存配置API
 */
app.post('/api/config', (req, res) => {
  try {
    const { userName, projectPaths, settings } = req.body;
    const config = {
      userName: userName || "",
      projectPaths: projectPaths || [],
      lastUsed: new Date().toISOString(),
      settings: {
        autoSave: true,
        defaultDateRange: "currentWeek",
        maxProjects: 10,
        ...settings
      }
    };
    
    if (saveConfig(config)) {
      res.json({ success: true, message: '配置保存成功' });
    } else {
      res.status(500).json({ success: false, error: '配置保存失败' });
    }
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

/**
 * 浏览目录API - 打开系统文件选择器
 */
app.get('/api/browse-directory', async (req, res) => {
  try {
    // 根据操作系统选择不同的方法打开文件选择器
    const platform = process.platform;
    let selectedPath = '';

    if (platform === 'darwin') {
      // macOS 使用 osascript (AppleScript)
      const script = `
        tell application "System Events"
          activate
          set folderPath to choose folder with prompt "请选择 Git 项目目录"
          return POSIX path of folderPath
        end tell
      `;
      
      try {
        selectedPath = execSync(`osascript -e '${script.replace(/'/g, "'\\''")}'`, { 
          encoding: 'utf-8',
          stdio: ['pipe', 'pipe', 'pipe'] // 抑制错误输出
        }).trim();
      } catch (err) {
        // 用户取消选择（-128 是用户取消的错误码）
        if (err.status === 1 || err.message.includes('-128')) {
          return res.json({ success: false, cancelled: true });
        }
        throw err;
      }
    } else if (platform === 'win32') {
      // Windows 使用 PowerShell
      const script = `
        Add-Type -AssemblyName System.Windows.Forms
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = '请选择 Git 项目目录'
        $result = $dialog.ShowDialog()
        if ($result -eq 'OK') {
          Write-Output $dialog.SelectedPath
        }
      `;
      
      try {
        selectedPath = execSync(`powershell -Command "${script}"`, {
          encoding: 'utf-8'
        }).trim();
        
        if (!selectedPath) {
          return res.json({ success: false, cancelled: true });
        }
      } catch (err) {
        throw err;
      }
    } else {
      // Linux - 尝试使用 zenity
      try {
        selectedPath = execSync('zenity --file-selection --directory --title="请选择 Git 项目目录"', {
          encoding: 'utf-8'
        }).trim();
      } catch (err) {
        if (err.status === 1) {
          return res.json({ success: false, cancelled: true });
        }
        // zenity 可能未安装
        return res.json({ 
          success: false, 
          error: 'Linux 系统需要安装 zenity：sudo apt-get install zenity' 
        });
      }
    }

    if (selectedPath) {
      res.json({ success: true, path: selectedPath });
    } else {
      res.json({ success: false, cancelled: true });
    }

  } catch (err) {
    // 忽略用户取消的错误
    if (err.message && err.message.includes('用户已取消')) {
      return res.json({ success: false, cancelled: true });
    }
    console.error('❌ 打开文件选择器失败：', err.message);
    res.json({ success: false, error: err.message });
  }
});

/**
 * 生成周报API
 */
app.post('/api/generate-report', async (req, res) => {
  try {
    const { userName, projectPaths, startDate, endDate } = req.body;

    if (!userName || !projectPaths || projectPaths.length === 0) {
      return res.status(400).json({ 
        success: false, 
        error: '请提供姓名和至少一个项目路径' 
      });
    }

    console.log(`\n${'='.repeat(60)}`);
    console.log(`🚀 开始生成周报`);
    console.log(`${'='.repeat(60)}`);
    console.log(`👤 周报负责人: ${userName}`);
    console.log(`📦 项目数量: ${projectPaths.length}`);
    console.log(`📅 日期范围: ${startDate || '本周一'} ~ ${endDate || '本周五'}\n`);

    // 1. 获取Git提交记录
    const commits = getGitCommits(projectPaths, startDate, endDate);
    
    if (commits.length === 0) {
      return res.json({
        success: true,
        message: '本周无提交记录，无需生成周报',
        tasks: 0,
        problems: 0
      });
    }

    // 2. 解析并处理提交记录
    const { tasks, problems } = await processCommits(commits, userName);

    // 3. 返回周报数据供预览
    const { startStr, endStr, year } = getWeekRange(startDate, endDate);
    const title = `${userName} ${year}年${startStr}-${endStr}工作周报`;

    res.json({
      success: true,
      message: '周报数据生成成功',
      title,
      tasks,
      problems,
      projectCount: projectPaths.length,
      dateRange: {
        start: startDate,
        end: endDate,
        startStr,
        endStr
      }
    });

  } catch (err) {
    console.error('❌ 生成周报失败：', err.message);
    res.status(500).json({ 
      success: false, 
      error: err.message 
    });
  }
});

/**
 * 下载周报API
 */
app.get('/download/:fileName', (req, res) => {
  const filePath = path.join(__dirname, 'output', req.params.fileName);
  
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: '文件不存在' });
  }

  res.download(filePath);
});

/**
 * 获取输出目录中的所有周报文件
 */
app.get('/api/reports', (req, res) => {
  const outputDir = path.join(__dirname, 'output');
  
  if (!fs.existsSync(outputDir)) {
    return res.json({ reports: [] });
  }

  const files = fs.readdirSync(outputDir)
    .filter(file => file.endsWith('.xlsx'))
    .map(file => {
      const stats = fs.statSync(path.join(outputDir, file));
      return {
        name: file,
        size: (stats.size / 1024).toFixed(2) + ' KB',
        createdAt: stats.birthtime,
        downloadUrl: `/download/${file}`
      };
    })
    .sort((a, b) => b.createdAt - a.createdAt);

  res.json({ reports: files });
});

/**
 * 生成Excel文件API
 */
app.post('/api/generate-excel', async (req, res) => {
  try {
    const { userName, title, tasks, problems, dateRange, emailConfig } = req.body;
    
    // 如果没有userName，从title中提取
    let finalUserName = userName;
    if (!finalUserName && title) {
      // 从标题中提取用户名，例如："陈毅 2025年10月20日-10月24日工作周报" -> "陈毅"
      const match = title.match(/^([^0-9\s]+)/);
      if (match) {
        finalUserName = match[1].trim();
      }
    }
    
    if (!finalUserName || !title || !tasks) {
      return res.status(400).json({
        success: false,
        error: '缺少必要参数: userName, title, tasks'
      });
    }

    // 生成Excel文件
    const { startStr, endStr } = dateRange;
    const fileName = `${finalUserName}_${startStr}-${endStr}_周报.xlsx`;
    const outputPath = path.join(__dirname, 'output', fileName);
    
    // 确保输出目录存在
    if (!fs.existsSync(path.join(__dirname, 'output'))) {
      fs.mkdirSync(path.join(__dirname, 'output'));
    }

    await generateExcel(finalUserName, tasks, problems, dateRange.start, dateRange.end, outputPath);

    // 邮件发送结果
    let emailResult = null;
    
    // 如果配置了邮件发送
    if (emailConfig && emailConfig.enabled) {
      const { to, cc, subject, content } = emailConfig;
      
      if (to && to.trim()) {
        // 构建邮件内容
        const emailSubject = subject || `${finalUserName} ${startStr}-${endStr} 工作周报`;
        const emailContent = content || `
          <div style="font-family: Arial, sans-serif; line-height: 1.6;">
            <h2 style="color: #1976d2;">📊 工作周报</h2>
            <p>您好，</p>
            <p>附件是 <strong>${finalUserName}</strong> 的 ${startStr}-${endStr} 工作周报，请查收。</p>
            <p>周报包含以下内容：</p>
            <ul>
              <li>📝 重点任务跟进：${tasks.length} 项</li>
              <li>📅 时间范围：${startStr} - ${endStr}</li>
              <li>👤 负责人：${finalUserName}</li>
            </ul>
            <p>如有疑问，请随时联系。</p>
            <hr style="margin: 20px 0; border: none; border-top: 1px solid #eee;">
            <p style="color: #666; font-size: 12px;">
              此邮件由周报生成器自动发送，请勿回复。
            </p>
          </div>
        `;
        
        emailResult = await sendEmail(to, cc, emailSubject, emailContent, outputPath, fileName);
      }
    }

    res.json({
      success: true,
      message: 'Excel文件生成成功',
      fileName,
      downloadUrl: `/download/${fileName}`,
      emailSent: emailResult ? emailResult.success : false,
      emailResult: emailResult
    });

  } catch (err) {
    console.error('❌ 生成Excel失败：', err.message);
    res.status(500).json({ 
      success: false, 
      error: err.message 
    });
  }
});

// 启动服务器
app.listen(PORT, () => {
  console.log(`\n✨ 周报生成器服务已启动！`);
  console.log(`🌐 访问地址: http://localhost:${PORT}`);
  console.log(`📝 请在浏览器中打开上述地址使用可视化界面\n`);
  
  // 检查 API Key 是否设置
  if (!process.env.DEEPSEEK_API_KEY) {
    console.log(`⚠️  警告: 未检测到 DEEPSEEK_API_KEY 环境变量`);
    console.log(`   请设置环境变量后重启服务：`);
    console.log(`   export DEEPSEEK_API_KEY="sk-your-api-key-here"\n`);
  } else {
    console.log(`✅ DeepSeek API Key 已配置\n`);
  }
});

