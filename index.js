// 加载环境变量配置（必须在最开头）
require('dotenv').config();

const ExcelJS = require('exceljs');
const { execSync } = require('child_process');
const { startOfWeek, endOfWeek, format } = require('date-fns');
const OpenAI = require('openai');
const fs = require('fs');

// ==================== 读取配置文件 =====================
function loadConfig() {
  try {
    if (fs.existsSync('./config.json')) {
      const data = fs.readFileSync('./config.json', 'utf8');
      const configData = JSON.parse(data);
      
      // 从config.json读取配置
      const userName = configData.userName || '用户';
      const projectPaths = configData.projectPaths || [];
      
      // 生成输出文件名
      const start = startOfWeek(new Date(), { weekStartsOn: 1 });
      const end = new Date(start);
      end.setDate(start.getDate() + 4);
      const startStr = format(start, 'MM月dd日');
      const endStr = format(end, 'MM月dd日');
      
      return {
        userName,
        projectPaths,
        templatePath: './周报模版.xlsx',
        outputPath: `./${userName}_${startStr}-${endStr}_周报.xlsx`,
        deepseekApiKey: process.env.DEEPSEEK_API_KEY,
        deepseekModel: 'deepseek-chat',
        weekStartsOnMonday: true,
        templateRows: {
          titleRow: 1,
          taskStartRow: 4,
          problemStartRow: 12
        }
      };
    }
  } catch (err) {
    console.error('❌ 读取配置文件失败:', err.message);
  }
  
  // 如果配置文件不存在或读取失败，使用默认配置
  console.log('⚠️  未找到config.json，使用默认配置');
  return {
    userName: '用户',
    projectPaths: [],
    templatePath: './周报模版.xlsx',
    outputPath: './周报.xlsx',
    deepseekApiKey: process.env.DEEPSEEK_API_KEY,
    deepseekModel: 'deepseek-chat',
    weekStartsOnMonday: true,
    templateRows: {
      titleRow: 1,
      taskStartRow: 4,
      problemStartRow: 12
    }
  };
}

const config = loadConfig();

// ==================== 工具函数：日期处理 ====================
/**
 * 获取本周日期范围（周一至周日）
 * @returns {Object} { start: Date, end: Date, startStr: 字符串, endStr: 字符串 }
 */
function getThisWeekRange() {
  const today = new Date();
  const start = startOfWeek(today, { weekStartsOn: 1 }); // 周一
  
  // 周五 = 周一 + 4天
  const end = new Date(start);
  end.setDate(start.getDate() + 4);

  return {
    start,
    end,
    startStr: format(start, 'MM月dd日'),
    endStr: format(end, 'MM月dd日'),
    year: format(start, 'yyyy'),
    month: format(start, 'MM')
  };
}

// ==================== 工具函数：Git提交记录提取 ====================
/**
 * 从Git仓库获取本周提交记录
 * @returns {Array} 结构化的提交记录数组
 */
function getGitCommits(projectPath) {
  const { start, end, startStr, endStr } = getThisWeekRange();
  const since = format(start, 'yyyy-MM-dd');
  const until = format(end, 'yyyy-MM-dd');

  console.log(`📅 查询时间范围: ${since} ~ ${until} (${startStr} ~ ${endStr})`);
  console.log(`📁 扫描项目: ${projectPath}`);

  try {
    // 执行Git命令：获取指定时间范围内的提交（含文件修改记录）
    const cmd = `git -C "${projectPath}" log \
      --since="${since}" --until="${until} 23:59:59" \
      --pretty=format:"COMMIT_SEP|%H|%an|%ad|%s" --date=short \
      --name-status`;

    const output = execSync(cmd, { encoding: 'utf-8' });
    
    if (!output.trim()) {
      console.log('ℹ️  此期间无提交记录');
      return [];
    }
    
    const lines = output.split('\n').filter(line => line.trim() !== '');

    // 解析提交记录为结构化数据
    const commits = [];
    let currentCommit = null;

    for (const line of lines) {
      if (line.startsWith('COMMIT_SEP|')) {
        // 新提交的分隔行：拆分哈希、作者、日期、提交信息
        if (currentCommit) commits.push(currentCommit);
        const [, hash, author, date, message] = line.split('|');
        currentCommit = {
          hash: hash.substring(0, 8), // 只保留前8位
          author,
          date,
          message: message.trim(),
          files: [], // 存储修改的文件列表
          project: require('path').basename(projectPath)
        };
      } else if (currentCommit) {
        // 处理文件修改记录（A=新增，M=修改，D=删除）
        currentCommit.files.push(line.trim());
      }
    }
    // 添加最后一个提交
    if (currentCommit) commits.push(currentCommit);

    console.log(`✅ 找到 ${commits.length} 条提交记录\n`);
    
    // 输出每条提交的详细信息
    commits.forEach((commit, index) => {
      console.log(`📝 提交 ${index + 1}/${commits.length}:`);
      console.log(`   提交哈希: ${commit.hash}`);
      console.log(`   提交作者: ${commit.author}`);
      console.log(`   提交日期: ${commit.date}`);
      console.log(`   提交信息: ${commit.message}`);
      if (commit.files.length > 0) {
        console.log(`   修改文件: ${commit.files.slice(0, 3).join(', ')}${commit.files.length > 3 ? '...' : ''}`);
      }
      console.log('');
    });
    
    return commits;
  } catch (err) {
    console.error('❌ 获取Git提交记录失败：', err.message);
    return [];
  }
}

// ==================== 初始化 DeepSeek 客户端 ====================
const openai = new OpenAI({
  baseURL: 'https://api.deepseek.com',
  apiKey: process.env.DEEPSEEK_API_KEY
});

// ==================== 工具函数：DeepSeek AI解析 ====================
/**
 * 调用DeepSeek API解析提交信息
 * @param {string} commitMessage - Git提交信息
 * @param {string} projectName - 项目名称
 * @returns {Object} 解析后的结构化数据
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
      model: config.deepseekModel,
      messages: [{ role: 'user', content: prompt }],
      temperature: 0.1, // 低随机性，确保格式稳定
      max_tokens: 200
    });

    const result = completion.choices[0].message.content.trim();
    const parsed = JSON.parse(result);
    
    const duration = Date.now() - startTime;
    console.log(`   ✅ AI 解析完成 (耗时: ${duration}ms) -> ${parsed.描述}`);
    
    return parsed;
  } catch (err) {
    console.error(`   ❌ DeepSeek API 调用失败（${projectName}）：`, err.message);
    // 解析失败时降级处理
    const fallback = {
      类型: '任务',
      分类: '未分类',
      描述: commitMessage.substring(0, 50), // 截断过长描述
      关联ID: '无'
    };
    
    console.log(`   ⚠️  使用降级方案: ${fallback.描述}`);
    return fallback;
  }
}

// ==================== 工具函数：处理提交记录为周报数据 ====================
/**
 * 将Git提交记录转换为周报所需的任务和问题数据
 * @param {Array} commits - Git提交记录数组
 * @returns {Object} { tasks: 重点任务数组, problems: 日常问题数组 }
 */
async function processCommits(commits) {
  const tasks = []; // 重点任务跟进项
  const problems = []; // 日常工作遇到的问题（保持空白）

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
      负责人: config.userName,
      协同人或部门: '无',
      完成进度: '100%',
      备注: ``
    });
  }

  console.log(`\n✅ DeepSeek AI 解析完成！共处理 ${commits.length} 条提交，生成 ${tasks.length} 条任务\n`);
  
  return { tasks, problems };
}

// ==================== 工具函数：填充Excel模板 ====================
/**
 * 将处理后的数据填充到Excel模板并生成最终周报
 * @param {Array} tasks - 重点任务数组
 * @param {Array} problems - 日常问题数组
 */
async function generateExcel(tasks, problems) {
  // 检查模板文件是否存在
  if (!fs.existsSync(config.templatePath)) {
    throw new Error(`❌ 模板文件不存在：${config.templatePath}`);
  }

  // 读取模板
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(config.templatePath);
  const worksheet = workbook.getWorksheet(1); // 获取第一个工作表

  // 1. 更新周报标题
  const { year, startStr, endStr } = getThisWeekRange();
  const title = `${config.userName} ${year}年${startStr}-${endStr}工作周报`;
  worksheet.getCell(`A${config.templateRows.titleRow}`).value = title;
  console.log(`📝 周报标题：${title}`);

  // 2. 填充重点任务表格
  tasks.forEach((task, index) => {
    const rowNum = config.templateRows.taskStartRow + index;
    const row = worksheet.getRow(rowNum);
    row.getCell(1).value = task.序号; // A列：序号
    row.getCell(2).value = task.重点需求或任务; // B列：重点需求/任务
    row.getCell(3).value = task.事项说明; // C列：事项说明
    row.getCell(4).value = task.启动日期; // D列：启动日期
    row.getCell(5).value = task.预计完成日期; // E列：预计完成日期
    row.getCell(6).value = task.负责人; // F列：负责人
    row.getCell(7).value = task.协同人或部门; // G列：协同人/部门
    row.getCell(8).value = task.完成进度; // H列：完成进度
    row.getCell(9).value = task.备注; // I列：备注
    
    // 设置单元格样式，特别优化"事项说明"列的换行显示
    for (let j = 1; j <= 9; j++) {
      const cell = row.getCell(j);
      if (j === 3) { // 事项说明列
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
  });
  console.log(`✅ 已填充 ${tasks.length} 条重点任务`);

  // 3. 填充日常问题表格
  problems.forEach((problem, index) => {
    const rowNum = config.templateRows.problemStartRow + index;
    const row = worksheet.getRow(rowNum);
    row.getCell(1).value = problem.序号; // A列：序号
    row.getCell(2).value = problem.问题分类; // B列：问题分类
    row.getCell(3).value = problem.具体描述; // C列：具体描述
    row.getCell(4).value = problem.提出日期; // D列：提出日期
    row.getCell(5).value = problem.解决方案; // E列：解决方案
    row.getCell(6).value = problem.解决日期; // F列：解决日期
  });
  console.log(`✅ 已填充 ${problems.length} 条日常问题`);

  // 保存文件
  await workbook.xlsx.writeFile(config.outputPath);
  console.log(`🎉 周报生成成功！路径：${config.outputPath}`);
}

// ==================== 主函数 ====================
async function main() {
  try {
    console.log(`\n${'='.repeat(60)}`);
    console.log(`🚀 周报生成器 - 命令行版本`);
    console.log(`${'='.repeat(60)}\n`);
    
    // 1. 检查DeepSeek API密钥
    if (!process.env.DEEPSEEK_API_KEY) {
      console.error('❌ 错误: 未设置 DEEPSEEK_API_KEY 环境变量');
      console.error('   请先设置环境变量：');
      console.error('   export DEEPSEEK_API_KEY="sk-your-api-key-here"\n');
      process.exit(1);
    }

    // 2. 显示配置信息
    console.log(`👤 周报负责人: ${config.userName}`);
    console.log(`📦 项目数量: ${config.projectPaths.length}`);
    console.log(`📁 项目路径: ${config.projectPaths.join(', ')}`);
    console.log(`📄 输出文件: ${config.outputPath}\n`);

    // 3. 获取Git提交记录（支持多项目）
    const commits = [];
    for (const projectPath of config.projectPaths) {
      console.log(`🔍 正在扫描项目: ${projectPath}`);
      const projectCommits = getGitCommits(projectPath);
      commits.push(...projectCommits);
    }
    
    if (commits.length === 0) {
      console.log('ℹ️ 本周（周一至周五）无提交记录，无需生成周报\n');
      return;
    }

    // 3. 解析并处理提交记录
    const { tasks, problems } = await processCommits(commits);

    // 4. 生成Excel周报
    await generateExcel(tasks, problems);

  } catch (err) {
    console.error('❌ 程序执行失败：', err.message);
    process.exit(1);
  }
}

// 启动程序
main();