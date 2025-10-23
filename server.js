// 加载环境变量配置（必须在最开头）
require('dotenv').config();

const express = require('express');
const ExcelJS = require('exceljs');
const { execSync } = require('child_process');
const { startOfWeek, endOfWeek, format, parseISO } = require('date-fns');
const OpenAI = require('openai');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;

// 中间件
app.use(express.json());
app.use(express.static('public'));

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

/**
 * 调用DeepSeek API解析提交信息
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
 * 处理提交记录为周报数据
 */
async function processCommits(commits, userName) {
  const tasks = [];
  const problems = []; // 保持空白，不填充任何内容

  console.log(`\n📊 开始使用 DeepSeek AI 解析 ${commits.length} 条提交记录...\n`);
  
  for (const [index, commit] of commits.entries()) {
    console.log(`\n[${index + 1}/${commits.length}] 处理提交: ${commit.hash} (${commit.date})`);
    const parsed = await parseCommitWithDeepSeek(commit.message, commit.project);

    // 所有AI生成的内容都放到重点任务表格中
    tasks.push({
      序号: tasks.length + 1,
      重点需求或任务: parsed.分类,
      事项说明: `[${commit.project}] ${parsed.描述}`,
      启动日期: commit.date,
      预计完成日期: commit.date,
      负责人: userName,
      协同人或部门: '无',
      完成进度: '100%',
      备注: '' // 备注栏为空
    });
  }

  console.log(`\n✅ DeepSeek AI 解析完成！共处理 ${commits.length} 条提交，生成 ${tasks.length} 条任务\n`);

  return { tasks, problems };
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

  // 填充重点任务表格 (A4:I7)
  const taskStartRow = 4;
  tasks.forEach((task, index) => {
    const rowNum = taskStartRow + index;
    if (rowNum > 7) return; // 限制在4行内
    
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

    // 保持白色背景和边框，支持换行
    for (let j = 1; j <= 9; j++) {
      const cell = row.getCell(j);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFFFF' } // 白色背景
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
      };
      
      // 特别优化"事项说明"列（第3列）的换行显示
      if (j === 3) {
        cell.alignment = { 
          horizontal: 'left', 
          vertical: 'top', 
          wrapText: true,
          indent: 1
        };
        // 设置行高以适应换行内容
        row.height = Math.max(60, (task.事项说明.length / 50) * 20);
      } else {
        cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
      }
    }

    row.commit(); // 提交行修改
  });
  console.log(`✅ 已填充 ${Math.min(tasks.length, 4)} 条重点任务`);

  // 填充日常问题表格
  const problemStartRow = 15;
  problems.forEach((problem, index) => {
    const rowNum = problemStartRow + index;
    if (rowNum > 19) return; // 限制在5行内
    
    const row = worksheet.getRow(rowNum);
    
    // 设置数据并支持换行
    row.getCell(1).value = problem.序号;
    row.getCell(2).value = problem.问题分类;
    row.getCell(3).value = problem.具体描述;
    row.getCell(4).value = problem.提出日期;
    row.getCell(5).value = problem.解决方案;
    row.getCell(6).value = problem.解决日期;

    // 保持白色背景和边框，支持换行
    for (let j = 1; j <= 6; j++) {
      const cell = row.getCell(j);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFFFF' } // 白色背景
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
      };
      cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
    }
    
    row.commit(); // 提交行修改
  });
  console.log(`✅ 已填充 ${Math.min(problems.length, 5)} 条日常问题`);

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
    const { userName, title, tasks, problems, dateRange } = req.body;
    
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

    res.json({
      success: true,
      message: 'Excel文件生成成功',
      fileName,
      downloadUrl: `/download/${fileName}`
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

