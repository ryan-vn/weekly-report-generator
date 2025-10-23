// 加载环境变量配置（必须在最开头）

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

// ==================== 工具函数：分组和聚合 ====================
/**
 * 按项目分组提交记录
 * @param {Array} commits - 所有提交记录
 * @returns {Object} 按项目名分组的提交记录
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

// ==================== 工具函数：DeepSeek AI解析 ====================
/**
 * 智能分析项目的所有提交，按模块聚合并生成周报条目
 * @param {string} projectName - 项目名称
 * @param {Array} commits - 该项目的所有提交记录
 * @returns {Array} 聚合后的任务列表
 */
async function analyzeProjectCommits(projectName, commits) {
  console.log(`🤖 [${projectName}] 正在分析 ${commits.length} 条提交记录...`);
  
  // 构建提交信息摘要，包含文件路径用于模块识别
  const commitSummary = commits.map((commit, index) => {
    const fileList = commit.files.slice(0, 5).join(', '); // 只取前5个文件
    const moreFiles = commit.files.length > 5 ? ` 等${commit.files.length}个文件` : '';
    return `${index + 1}. [${commit.date}] ${commit.message}\n   修改文件: ${fileList}${moreFiles}`;
  }).join('\n\n');

  const prompt = `你是一个专业的技术周报生成助手。请分析以下项目的 Git 提交记录，智能识别代码模块和功能，将相关提交聚合成高质量的周报条目。

项目名称: ${projectName}
提交记录（共 ${commits.length} 条）:

${commitSummary}

分析要求:
1. **细粒度拆分**: 根据文件路径、提交信息、功能点，尽可能细粒度地拆分任务
2. **多维度识别**: 识别代码模块、功能点、bug修复、性能优化等不同维度，每个维度都可以是独立任务
3. **最大化条目数**: 目标是生成尽可能多的任务条目，展示丰富的工作内容
4. **独立展示**: 即使是小的改动，如果属于不同的功能或模块，也要分开展示
5. **工作描述**: 用专业、简洁的语言描述工作内容，避免过于技术化的细节且 让领导看到做了很多任务 而且同事看了任务很难实现
6. **关键改动**: 总结该任务的主要改动点（2-4个要点）

拆分策略（尽可能多拆分）:
假设有10个提交涉及以下内容：
- 2个提交开发用户登录功能 → 生成1条"用户登录功能开发"任务
- 1个提交优化用户登录性能 → 生成1条"用户登录性能优化"任务
- 2个提交开发订单列表功能 → 生成1条"订单列表功能开发"任务
- 1个提交修复订单bug → 生成1条"订单模块bug修复"任务
- 2个提交开发支付接口 → 生成1条"支付接口开发"任务
- 1个提交数据库优化 → 生成1条"数据库性能优化"任务
- 1个提交UI样式调整 → 生成1条"界面UI优化"任务
最终应该输出 7 条独立的任务，而不是合并成3-4条

核心原则: 
- 不同模块 → 分开
- 不同功能点 → 分开
- 不同分类（开发/修复/优化） → 分开
- 同一模块的开发和优化 → 分开
- 同一模块的功能和bug修复 → 分开

输出格式（必须是有效的 JSON 数组）:
[
  {
    "模块": "具体的模块或功能名称",
    "分类": "开发新功能|修复bug|优化性能|代码重构|文档更新",
    "描述": "简洁专业的工作描述（15-40字）",
    "关键改动": ["改动点1", "改动点2", "改动点3"],
    "涉及提交数": 提交数量
  }
]

注意事项:
- 🎯 **最重要**: 尽可能生成更多的任务条目，不要过度聚合
- ⚠️ **关键**: 不同模块、不同功能、不同分类都要分开
- ⚠️ **关键**: 宁可拆分过细，也不要合并过多
- 💡 即使只有1-2个提交，如果是独立的功能点，也应该单独成条
- 💡 同一个文件的不同功能改动，也可以拆分成多条
- 📊 目标: 让提交数和任务数的比例尽可能接近 1:1
- 📝 描述要站在周报汇报的角度，突出工作价值
- 🚫 避免使用"修复了一个bug"这样的模糊描述，要具体说明修复了什么问题

请直接输出 JSON 数组，不要有其他内容。`;

  try {
    const startTime = Date.now();
    
    const completion = await openai.chat.completions.create({
      model: config.deepseekModel,
      messages: [{ role: 'user', content: prompt }],
      temperature: 0.3, // 适度的创造性
      max_tokens: 4000 // 增加 token 以支持生成更多任务条目
    });

    const result = completion.choices[0].message.content.trim();
    
    // 尝试解析 JSON
    let parsedTasks;
    try {
      // 移除可能的 markdown 代码块标记
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
      console.log(`   ${index + 1}. [${task.模块}] ${task.描述} (合并${task.涉及提交数}个提交)`);
    });
    
    return parsedTasks;
  } catch (err) {
    console.error(`   ❌ DeepSeek API 调用失败:`, err.message);
    
    // 降级方案：简单分组
    console.log(`   ⚠️  使用降级方案: 按日期简单分组\n`);
    return [{
      模块: '未分类',
      分类: '开发任务',
      描述: `${projectName} 项目开发工作（${commits.length}个提交）`,
      关键改动: commits.slice(0, 3).map(c => c.message),
      涉及提交数: commits.length
    }];
  }
}

/**
 * 调用DeepSeek API解析单个提交信息（旧方法，保留作为备用）
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
 * 将Git提交记录转换为周报所需的任务和问题数据（智能模块聚合版本）
 * @param {Array} commits - Git提交记录数组
 * @returns {Object} { tasks: 重点任务数组, problems: 日常问题数组 }
 */
async function processCommits(commits) {
  const tasks = []; // 重点任务跟进项
  const problems = []; // 日常工作遇到的问题（保持空白）

  console.log(`\n${'='.repeat(70)}`);
  console.log(`📊 智能分析模式：细粒度拆分，最大化任务条目数`);
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
      // 计算日期范围
      const dates = projectCommits.map(c => c.date).sort();
      const startDate = dates[0];
      const endDate = dates[dates.length - 1];
      
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
        负责人: config.userName,
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
  console.log(`   🎯 拆分比例: ${(tasks.length / commits.length).toFixed(2)}:1 (任务数:提交数)`);
  console.log(`   💡 提示: 比例越接近1:1，说明拆分越细致`);
  console.log(`${'='.repeat(70)}\n`);
  
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