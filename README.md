# BfTaskBoard

一个功能强大的Windows桌面任务管理和数据表格应用程序，支持多种数据类型、AI智能建表、多语言界面和主题切换。

## 主要功能

### 核心功能
- **多标签页管理**：支持创建、编辑、排序多个数据表格标签页
- **多种数据类型**：
  - 文本（Text）：普通文本内容
  - 单选（Single）：带彩色标记的选项
  - 图片（Image）：支持粘贴板图片（Ctrl+V）
  - 待办事项（TodoList）：可勾选的任务列表
  - 文本域（TextArea）：长文本编辑，支持自动保存
- **AI智能建表**：支持OpenAI、DeepSeek、Claude三种AI服务自动生成表格结构
- **数据导入导出**：支持Excel文件导出（.xlsx格式）

### 高级功能
- **全局搜索**：跨标签页搜索内容，支持定位跳转（Ctrl+S）
- **数据处理**：
  - **排序功能**：文本列支持升序/降序排序，点击列标题排序图标
  - **筛选功能**：单选列支持按选项过滤数据，保留原始数据
  - **列求和**：对数值列自动计算总和，右键菜单操作
  - **一键切割**：快速将长文本按分隔符切割成多行
  - **JSON编辑器**：内置JSON查看和编辑器，支持格式化和验证
- **批量操作**：
  - 批量删除选中行
  - 批量复制粘贴
  - 列数据批量填充
- **界面定制**：
  - 双语支持（中文/英文），自动检测系统语言
  - 主题切换（暗色/亮色），实时预览
  - 标签彩色标记，6种预设颜色
  - 列宽自适应调整
- **快捷键支持**：
  - Ctrl+T：新建标签页
  - Ctrl+Q：标签页管理
  - Ctrl+S：全局搜索
  - Ctrl+V：粘贴图片（图片列）
  - Delete：删除选中行

### 数据管理
- **自动保存**：实时保存数据变更
- **性能优化**：大数据量异步处理
- **数据持久化**：JSON格式本地存储
- **文件监控**：显示最后修改时间

## 使用方法

### 基本操作
1. **创建标签页**：点击顶部"+"按钮或使用Ctrl+T
2. **编辑单元格**：双击单元格进入编辑模式
3. **添加行**：点击底部"添加行"按钮
4. **管理列**：右键点击列标题进行编辑、删除等操作
5. **AI建表**：点击顶部机器人图标，输入需求自动生成表格

### 高级操作
1. **排序数据**：
   - 鼠标悬停在文本列标题，点击排序图标
   - 支持升序和降序切换
   - 排序不会修改原始数据
   
2. **筛选数据**：
   - 鼠标悬停在单选列标题，点击筛选图标
   - 选择要显示的选项
   - 支持多选筛选
   
3. **列求和**：
   - 右键点击数值列
   - 选择"列求和"
   - 自动计算并显示总和
   
4. **一键切割**：
   - 选中包含分隔符的文本单元格
   - 右键选择"一键切割"
   - 输入分隔符（如逗号、分号等）
   - 自动将文本切割成多行
   
5. **JSON编辑**：
   - 双击包含JSON的文本单元格
   - 自动识别并打开JSON编辑器
   - 支持格式化、验证和编辑

### 数据类型使用
- **文本**：直接输入文本内容
- **单选**：点击显示选项列表，支持自定义选项和颜色
- **图片**：点击上传或直接粘贴（Ctrl+V）
- **待办事项**：点击添加任务项，勾选完成状态
- **文本域**：点击打开编辑器，支持长文本编辑

## 技术要点

### 技术栈
- **框架**：.NET 6.0 + Windows Forms
- **语言**：C# 10.0
- **数据存储**：JSON (Newtonsoft.Json)
- **Excel处理**：EPPlus
- **HTTP客户端**：HttpClient（用于AI API调用）

### 架构设计
- **MVC模式**：视图与数据分离
- **服务层**：
  - `DataService`：数据持久化管理
  - `LanguageService`：多语言支持
  - `ThemeService`：主题管理
  - `AIService`：AI接口集成
- **自定义控件**：
  - `ModernDataGridView`：增强的数据表格
  - `TabButton`：自定义标签按钮
  - `TodoListControl`：待办事项控件

### 性能优化
- 异步文件保存
- 防抖动处理（文本域自动保存）
- 大数据量分页处理建议
- JSON压缩存储

## 构建方法

### 环境要求
- Windows 10/11
- .NET 6.0 SDK 或更高版本
- Visual Studio 2022（推荐）或 VS Code

### 构建步骤
1. 克隆仓库
```bash
git clone https://github.com/binfenhulian/BfTaskBoard.git
cd BfTaskBoard
```

2. 构建项目
```bash
build.bat
```
然后选择：
- 选项 1：仅构建（用于开发）
- 选项 2：发布为单个.exe文件（推荐）

或使用dotnet命令：
```bash
# 仅构建
dotnet build -c Release

# 发布为单文件
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true -o publish
```

3. 运行程序
- 开发版本：`bin\Release\net6.0-windows\BfTaskBoard.exe`
- 单文件版本：`publish\BfTaskBoard.exe`

### 单文件发布说明
使用选项2发布后，会在`publish`文件夹中生成一个独立的`BfTaskBoard.exe`文件（约50-70MB）。

**优点**：
- 无需安装.NET运行时
- 可直接复制到任何Windows 10/11电脑运行
- 包含所有依赖项
- 便于分发和部署

## 配置说明

### AI服务配置
程序支持三种AI服务提供商：
- **OpenAI**：需要OpenAI API密钥
- **DeepSeek**：需要DeepSeek API密钥
- **Claude**：需要Anthropic API密钥

API密钥在使用AI建表功能时输入，程序会自动保存。

### 数据存储位置
- 应用数据：`%APPDATA%\BfTaskBoard\appdata.json`
- 主题设置：`%APPDATA%\BfTaskBoard\theme.json`
- AI日志：`%APPDATA%\BfTaskBoard\ai_logs\`

## 开源协议

本项目采用 MIT 协议开源，详见 [LICENSE](LICENSE) 文件。

## 贡献指南

欢迎提交 Issue 和 Pull Request！

### 开发建议
1. Fork 本仓库
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 提交 Pull Request

## 更新日志

### v1.0.0 (2024-01)
- 初始版本发布
- 支持多种数据类型
- AI智能建表功能
- 多语言和主题支持
- Excel导出功能

## 联系方式

- 项目主页：[https://github.com/binfenhulian/BfTaskBoard](https://github.com/binfenhulian/BfTaskBoard)
- Bug反馈：[Issues](https://github.com/binfenhulian/BfTaskBoard/issues)

## 致谢

感谢所有为本项目做出贡献的开发者！