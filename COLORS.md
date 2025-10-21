# 颜色规范说明

## 🎨 周报模板颜色方案

根据原始模板图片，我们实现了完整的颜色一致性：

### 📋 重点任务表格

| 元素 | 颜色 | 十六进制 | RGB | 说明 |
|------|------|----------|-----|------|
| 表头背景 | 绿色 | #66CC00 | RGB(102, 204, 0) | 重点任务表头行 |
| 表头文字 | 黑色 | #000000 | RGB(0, 0, 0) | 表头文字颜色 |
| 数据区背景 | 白色 | #FFFFFF | RGB(255, 255, 255) | 数据行背景 |
| 边框 | 灰色 | #CCCCCC | RGB(204, 204, 204) | 单元格边框 |

### 🐛 日常问题表格

| 元素 | 颜色 | 十六进制 | RGB | 说明 |
|------|------|----------|-----|------|
| 表头背景 | 绿色 | #66CC00 | RGB(102, 204, 0) | 日常问题表头行 |
| 表头文字 | 黑色 | #000000 | RGB(0, 0, 0) | 表头文字颜色 |
| 数据区背景 | 白色 | #FFFFFF | RGB(255, 255, 255) | 数据行背景 |
| 边框 | 灰色 | #CCCCCC | RGB(204, 204, 204) | 单元格边框 |

### 🏷️ 标题和说明

| 元素 | 颜色 | 十六进制 | RGB | 说明 |
|------|------|----------|-----|------|
| 主标题 | 黑色 | #000000 | RGB(0, 0, 0) | 周报标题 |
| 章节标题 | 红色 | #FF0000 | RGB(255, 0, 0) | "一、重点任务跟进项" |
| 说明文字 | 黑色 | #000000 | RGB(0, 0, 0) | 说明和备注文字 |

## 🔧 技术实现

### ExcelJS 颜色设置

```javascript
// 绿色表头背景
fill: {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FF66CC00' }
}

// 白色数据背景
fill: {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFFFFFF' }
}

// 灰色边框
border: {
  top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
  left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
  bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
  right: { style: 'thin', color: { argb: 'FFCCCCCC' } }
}

// 红色文字
font: { color: { argb: 'FFFF0000' } }
```

### ARGB 颜色格式

- `FF` = Alpha 通道（不透明度，FF = 100%）
- `66CC00` = RGB 十六进制值
- 完整格式：`FF66CC00`

## 📊 颜色验证

生成周报后，可以通过以下方式验证颜色：

```javascript
// 检查表头颜色
const headerCell = worksheet.getRow(3).getCell(1);
console.log('表头背景色:', headerCell.fill.fgColor?.argb); // 应显示 FF66CC00

// 检查数据区颜色
const dataCell = worksheet.getRow(4).getCell(1);
console.log('数据背景色:', dataCell.fill.fgColor?.argb); // 应显示 FFFFFFFF
```

## 🎯 一致性保证

1. **模板文件**：`周报模版_带颜色.xlsx` 包含所有颜色设置
2. **代码填充**：每次填充数据时重新应用颜色样式
3. **边框保持**：确保所有单元格都有统一的灰色边框
4. **背景保持**：数据行始终保持白色背景

## 🔍 故障排除

### 颜色不显示

1. 检查 Excel 版本是否支持颜色
2. 确认模板文件路径正确
3. 验证 ARGB 颜色格式

### 颜色不一致

1. 检查 `generateExcel` 函数中的颜色设置
2. 确认 `row.commit()` 被正确调用
3. 验证模板文件没有被修改

### 边框问题

1. 检查边框颜色设置
2. 确认边框样式为 'thin'
3. 验证所有四个方向都有边框设置
