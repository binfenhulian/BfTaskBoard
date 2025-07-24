using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace TaskBoard.Services
{
    public class LanguageService
    {
        private static LanguageService _instance;
        private Dictionary<string, Dictionary<string, string>> _translations;
        private string _currentLanguage;
        private readonly string _languageSettingPath;

        public static LanguageService Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new LanguageService();
                }
                return _instance;
            }
        }

        public event EventHandler<string> LanguageChanged;

        private LanguageService()
        {
            var appDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "BfTaskBoard");
            Directory.CreateDirectory(appDataFolder);
            _languageSettingPath = Path.Combine(appDataFolder, "language.json");
            
            InitializeTranslations();
            LoadLanguageSetting();
        }

        private void InitializeTranslations()
        {
            _translations = new Dictionary<string, Dictionary<string, string>>
            {
                ["zh-CN"] = new Dictionary<string, string>
                {
                    // 窗口标题
                    ["AppTitle"] = "BfTaskBoard 2.0 - 现代任务管理",
                    
                    // 按钮
                    ["Search"] = "搜索",
                    ["Filter"] = "筛选",
                    ["Clear"] = "清除",
                    ["QuickCreate"] = "快速新建",
                    ["AddTab"] = "添加标签页",
                    ["DeleteTab"] = "删除标签页",
                    ["RenameTab"] = "重命名",
                    ["CloseTab"] = "关闭标签页",
                    ["AddRow"] = "添加新行",
                    ["AddColumn"] = "添加列",
                    ["EditColumn"] = "编辑列",
                    ["DeleteColumn"] = "删除列",
                    ["SumColumn"] = "列求和",
                    ["OK"] = "确定",
                    ["Cancel"] = "取消",
                    ["Save"] = "保存",
                    ["Export"] = "导出",
                    ["Restore"] = "恢复",
                    ["Delete"] = "删除",
                    ["Hide"] = "隐藏",
                    ["Show"] = "显示",
                    ["SelectAll"] = "全选/取消全选",
                    
                    // 菜单项
                    ["OpenDataFolder"] = "打开数据目录",
                    ["TabManagement"] = "标签管理",
                    ["JsonView"] = "JSON视图",
                    ["ExportExcel"] = "导出Excel",
                    ["OneClickCut"] = "一键切割",
                    ["RestoreBackup"] = "恢复备份",
                    ["Language"] = "语言",
                    
                    // 右键菜单
                    ["AddRowAbove"] = "在上方插入行",
                    ["AddRowBelow"] = "在下方插入行",
                    ["CopyRow"] = "复制行",
                    ["DeleteRow"] = "删除行",
                    
                    // 列类型
                    ["Text"] = "文本",
                    ["SingleSelect"] = "单选",
                    ["MultiSelect"] = "多选",
                    ["Image"] = "图片",
                    ["TodoList"] = "待办事项",
                    ["TextArea"] = "文本域",
                    
                    // 对话框标题
                    ["SearchDialog"] = "全局搜索",
                    ["SearchResult"] = "搜索结果",
                    ["NewTab"] = "新建标签页",
                    ["RenameTabDialog"] = "重命名标签页",
                    ["AddColumnDialog"] = "添加列",
                    ["EditColumnDialog"] = "编辑列",
                    ["TabManagementDialog"] = "标签管理",
                    ["ExportExcelDialog"] = "导出Excel - 选择要导出的标签页",
                    ["RestoreBackupDialog"] = "恢复备份",
                    ["JsonViewDialog"] = "JSON视图 - 配置文件编辑器",
                    ["SumResultDialog"] = "列求和结果",
                    
                    // 提示信息
                    ["EnterSearchKeyword"] = "请输入搜索关键词：",
                    ["NoSearchResult"] = "未找到匹配的内容",
                    ["FoundInTabs"] = "在 {0} 个标签页中找到匹配内容",
                    ["ClickToNavigate"] = "点击可快速定位",
                    ["SelectTabFirst"] = "请先选择一个标签页",
                    ["EnterTabName"] = "请输入标签页名称：",
                    ["EnterNewTabName"] = "请输入新的标签页名称：",
                    ["ConfirmDeleteTab"] = "确定要删除标签页 \"{0}\" 吗？\\n所有数据将被永久删除。",
                    ["ConfirmCloseTab"] = "确定要关闭标签页 \"{0}\" 吗？\\n数据会保留，下次可以重新打开。",
                    ["ConfirmDeleteColumn"] = "确定要删除列 \"{0}\" 吗？\\n这将删除该列的所有数据。",
                    ["ConfirmDeleteRow"] = "确定要删除这一行吗？",
                    ["AtLeastOneTab"] = "至少需要保留一个标签页",
                    ["SelectAtLeastOne"] = "请至少选择一个标签页",
                    ["SelectBackupFile"] = "请选择要恢复的备份：",
                    ["InvalidJson"] = "JSON格式无效：{0}",
                    ["LoadJsonFailed"] = "加载JSON失败：{0}",
                    ["SaveJsonFailed"] = "保存JSON失败：{0}",
                    ["ExportSuccess"] = "导出成功！\\n文件保存在：{0}",
                    ["ExportFailed"] = "导出失败：{0}",
                    ["RestoreFailed"] = "恢复失败：{0}",
                    ["CutFailed"] = "切割失败：{0}",
                    
                    // 一键切割确认
                    ["ConfirmCut"] = "一键切割将：\\n1. 备份当前所有配置\\n2. 创建全新的空白配置\\n3. 原有数据将保存在备份中\\n\\n是否继续？",
                    ["CutSuccess"] = "切割成功！\\n\\n备份已保存至：\\n{0}\\n\\n现在已切换到新的空白配置。",
                    
                    // 恢复备份确认
                    ["ConfirmRestore"] = "恢复备份将：\\n1. 备份当前配置\\n2. 恢复选中的备份：{0}\\n3. 当前数据将被替换\\n\\n是否继续？",
                    ["RestoreSuccess"] = "恢复成功！\\n\\n已从备份 {0} 恢复。\\n原配置已备份为：{1}",
                    
                    // 列求和
                    ["ColumnTypeNotSupported"] = "列 \"{0}\" 的类型不支持求和操作。\\n只有文本类型的列可以进行数值求和。",
                    ["SumResult"] = "列 \"{0}\" 求和结果：\\n\\n",
                    ["Sum"] = "总和：{0:F2}\\n",
                    ["ValidCount"] = "有效数值个数：{0}\\n",
                    ["InvalidCount"] = "非数值个数：{0}\\n",
                    ["InvalidExamples"] = "非数值示例：{0}",
                    
                    // 表格列标题
                    ["BackupFile"] = "备份文件",
                    ["BackupTime"] = "备份时间",
                    ["FileSize"] = "文件大小",
                    ["TabCount"] = "标签页数",
                    ["TabName"] = "标签名称",
                    ["Status"] = "状态",
                    ["ColumnCount"] = "列数",
                    ["RowCount"] = "行数",
                    ["Operation"] = "操作",
                    ["MoveUp"] = "上移",
                    ["MoveDown"] = "下移",
                    
                    // 其他
                    ["SortAscending"] = "升序排列",
                    ["SortDescending"] = "降序排列",
                    ["ClearSort"] = "清除排序",
                    ["Hidden"] = "隐藏",
                    ["Visible"] = "显示",
                    ["Unknown"] = "未知",
                    ["AutoSaveEnabled"] = "自动保存已启用",
                    ["Typing"] = "正在输入...",
                    ["AutoSaved"] = "已自动保存 - {0}",
                    ["ManualSaved"] = "已手动保存 - {0}",
                    ["NewTaskList"] = "新任务列表",
                    ["Task"] = "任务",
                    ["DailyPlan"] = "每日计划",
                    ["NotStarted"] = "未开始",
                    ["InProgress"] = "进行中",
                    ["Completed"] = "已完成",
                    ["ColumnName"] = "列名:",
                    ["ColumnType"] = "列类型:",
                    ["Images"] = "{0}张图片",
                    ["TodoProgress"] = "[{0}/{1}]",
                    
                    // AI建表
                    ["AICreateTable"] = "AI建表",
                    ["AIProvider"] = "AI提供商",
                    ["APIKey"] = "API密钥",
                    ["TableRequirement"] = "表格需求",
                    ["GenerateTable"] = "生成表格",
                    ["EnterRequirement"] = "请描述您需要的表格...",
                    ["GeneratingTable"] = "正在生成表格...",
                    ["TableGeneratedSuccess"] = "表格生成成功！",
                    ["TableGenerateFailed"] = "表格生成失败：{0}",
                    ["InvalidAPIKey"] = "API密钥无效",
                    ["PleaseEnterRequirement"] = "请输入表格需求"
                },
                
                ["en-US"] = new Dictionary<string, string>
                {
                    // Window title
                    ["AppTitle"] = "BfTaskBoard 2.0 - Modern Task Management",
                    
                    // Buttons
                    ["Search"] = "Search",
                    ["Filter"] = "Filter",
                    ["Clear"] = "Clear",
                    ["QuickCreate"] = "Quick Create",
                    ["AddTab"] = "Add Tab",
                    ["DeleteTab"] = "Delete Tab",
                    ["RenameTab"] = "Rename",
                    ["CloseTab"] = "Close Tab",
                    ["AddRow"] = "Add New Row",
                    ["AddColumn"] = "Add Column",
                    ["EditColumn"] = "Edit Column",
                    ["DeleteColumn"] = "Delete Column",
                    ["SumColumn"] = "Sum Column",
                    ["OK"] = "OK",
                    ["Cancel"] = "Cancel",
                    ["Save"] = "Save",
                    ["Export"] = "Export",
                    ["Restore"] = "Restore",
                    ["Delete"] = "Delete",
                    ["Hide"] = "Hide",
                    ["Show"] = "Show",
                    ["SelectAll"] = "Select All/Deselect All",
                    
                    // Menu items
                    ["OpenDataFolder"] = "Open Data Folder",
                    ["TabManagement"] = "Tab Management",
                    ["JsonView"] = "JSON View",
                    ["ExportExcel"] = "Export Excel",
                    ["OneClickCut"] = "One-Click Cut",
                    ["RestoreBackup"] = "Restore Backup",
                    ["Language"] = "Language",
                    
                    // Context menu
                    ["AddRowAbove"] = "Insert Row Above",
                    ["AddRowBelow"] = "Insert Row Below",
                    ["CopyRow"] = "Copy Row",
                    ["DeleteRow"] = "Delete Row",
                    
                    // Column types
                    ["Text"] = "Text",
                    ["SingleSelect"] = "Single Select",
                    ["MultiSelect"] = "Multi Select",
                    ["Image"] = "Image",
                    ["TodoList"] = "Todo List",
                    ["TextArea"] = "Text Area",
                    
                    // Dialog titles
                    ["SearchDialog"] = "Global Search",
                    ["SearchResult"] = "Search Results",
                    ["NewTab"] = "New Tab",
                    ["RenameTabDialog"] = "Rename Tab",
                    ["AddColumnDialog"] = "Add Column",
                    ["EditColumnDialog"] = "Edit Column",
                    ["TabManagementDialog"] = "Tab Management",
                    ["ExportExcelDialog"] = "Export Excel - Select tabs to export",
                    ["RestoreBackupDialog"] = "Restore Backup",
                    ["JsonViewDialog"] = "JSON View - Configuration Editor",
                    ["SumResultDialog"] = "Column Sum Result",
                    
                    // Prompts
                    ["EnterSearchKeyword"] = "Enter search keyword:",
                    ["NoSearchResult"] = "No matching content found",
                    ["FoundInTabs"] = "Found matches in {0} tabs",
                    ["ClickToNavigate"] = "Click to navigate",
                    ["SelectTabFirst"] = "Please select a tab first",
                    ["EnterTabName"] = "Enter tab name:",
                    ["EnterNewTabName"] = "Enter new tab name:",
                    ["ConfirmDeleteTab"] = "Are you sure you want to delete tab \"{0}\"?\\nAll data will be permanently deleted.",
                    ["ConfirmCloseTab"] = "Are you sure you want to close tab \"{0}\"?\\nData will be preserved and can be reopened later.",
                    ["ConfirmDeleteColumn"] = "Are you sure you want to delete column \"{0}\"?\\nThis will delete all data in this column.",
                    ["ConfirmDeleteRow"] = "Are you sure you want to delete this row?",
                    ["AtLeastOneTab"] = "At least one tab must be kept",
                    ["SelectAtLeastOne"] = "Please select at least one tab",
                    ["SelectBackupFile"] = "Select backup to restore:",
                    ["InvalidJson"] = "Invalid JSON format: {0}",
                    ["LoadJsonFailed"] = "Failed to load JSON: {0}",
                    ["SaveJsonFailed"] = "Failed to save JSON: {0}",
                    ["ExportSuccess"] = "Export successful!\\nFile saved at: {0}",
                    ["ExportFailed"] = "Export failed: {0}",
                    ["RestoreFailed"] = "Restore failed: {0}",
                    ["CutFailed"] = "Cut failed: {0}",
                    
                    // One-click cut confirmation
                    ["ConfirmCut"] = "One-click cut will:\\n1. Backup all current configuration\\n2. Create a new blank configuration\\n3. Original data will be saved in backup\\n\\nContinue?",
                    ["CutSuccess"] = "Cut successful!\\n\\nBackup saved to:\\n{0}\\n\\nNow switched to new blank configuration.",
                    
                    // Restore backup confirmation
                    ["ConfirmRestore"] = "Restore backup will:\\n1. Backup current configuration\\n2. Restore selected backup: {0}\\n3. Current data will be replaced\\n\\nContinue?",
                    ["RestoreSuccess"] = "Restore successful!\\n\\nRestored from backup {0}.\\nOriginal configuration backed up as: {1}",
                    
                    // Column sum
                    ["ColumnTypeNotSupported"] = "Column \"{0}\" type does not support sum operation.\\nOnly text type columns can be summed.",
                    ["SumResult"] = "Column \"{0}\" sum result:\\n\\n",
                    ["Sum"] = "Sum: {0:F2}\\n",
                    ["ValidCount"] = "Valid numbers: {0}\\n",
                    ["InvalidCount"] = "Invalid values: {0}\\n",
                    ["InvalidExamples"] = "Invalid examples: {0}",
                    
                    // Table headers
                    ["BackupFile"] = "Backup File",
                    ["BackupTime"] = "Backup Time",
                    ["FileSize"] = "File Size",
                    ["TabCount"] = "Tab Count",
                    ["TabName"] = "Tab Name",
                    ["Status"] = "Status",
                    ["ColumnCount"] = "Columns",
                    ["RowCount"] = "Rows",
                    ["Operation"] = "Operation",
                    ["MoveUp"] = "Move Up",
                    ["MoveDown"] = "Move Down",
                    
                    // Others
                    ["SortAscending"] = "Sort Ascending",
                    ["SortDescending"] = "Sort Descending",
                    ["ClearSort"] = "Clear Sort",
                    ["Hidden"] = "Hidden",
                    ["Visible"] = "Visible",
                    ["Unknown"] = "Unknown",
                    ["AutoSaveEnabled"] = "Auto-save enabled",
                    ["Typing"] = "Typing...",
                    ["AutoSaved"] = "Auto-saved - {0}",
                    ["ManualSaved"] = "Manually saved - {0}",
                    ["NewTaskList"] = "New Task List",
                    ["Task"] = "Task",
                    ["DailyPlan"] = "Daily Plan",
                    ["NotStarted"] = "Not Started",
                    ["InProgress"] = "In Progress",
                    ["Completed"] = "Completed",
                    ["ColumnName"] = "Column Name:",
                    ["ColumnType"] = "Column Type:",
                    ["Images"] = "{0} images",
                    ["TodoProgress"] = "[{0}/{1}]",
                    
                    // AI Create Table
                    ["AICreateTable"] = "AI Create Table",
                    ["AIProvider"] = "AI Provider",
                    ["APIKey"] = "API Key",
                    ["TableRequirement"] = "Table Requirement",
                    ["GenerateTable"] = "Generate Table",
                    ["EnterRequirement"] = "Please describe the table you need...",
                    ["GeneratingTable"] = "Generating table...",
                    ["TableGeneratedSuccess"] = "Table generated successfully!",
                    ["TableGenerateFailed"] = "Failed to generate table: {0}",
                    ["InvalidAPIKey"] = "Invalid API Key",
                    ["PleaseEnterRequirement"] = "Please enter table requirement"
                }
            };
        }

        private void LoadLanguageSetting()
        {
            try
            {
                if (File.Exists(_languageSettingPath))
                {
                    var json = File.ReadAllText(_languageSettingPath);
                    var settings = JsonConvert.DeserializeObject<LanguageSettings>(json);
                    _currentLanguage = settings?.Language ?? GetSystemLanguage();
                }
                else
                {
                    _currentLanguage = GetSystemLanguage();
                    SaveLanguageSetting();
                }
            }
            catch
            {
                _currentLanguage = GetSystemLanguage();
            }

            // 确保语言存在
            if (!_translations.ContainsKey(_currentLanguage))
            {
                _currentLanguage = "en-US";
            }
        }

        private void SaveLanguageSetting()
        {
            try
            {
                var settings = new LanguageSettings { Language = _currentLanguage };
                var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                File.WriteAllText(_languageSettingPath, json);
            }
            catch { }
        }

        private string GetSystemLanguage()
        {
            var culture = CultureInfo.CurrentUICulture;
            
            // 检查是否是中文
            if (culture.Name.StartsWith("zh"))
            {
                return "zh-CN";
            }
            
            // 默认英文
            return "en-US";
        }

        public string GetString(string key)
        {
            if (_translations.ContainsKey(_currentLanguage) && 
                _translations[_currentLanguage].ContainsKey(key))
            {
                return _translations[_currentLanguage][key];
            }
            
            // 如果找不到，返回英文
            if (_translations["en-US"].ContainsKey(key))
            {
                return _translations["en-US"][key];
            }
            
            // 如果都找不到，返回key本身
            return key;
        }

        public string GetString(string key, params object[] args)
        {
            var format = GetString(key);
            try
            {
                return string.Format(format, args);
            }
            catch
            {
                return format;
            }
        }

        public string CurrentLanguage => _currentLanguage;

        public string[] AvailableLanguages => _translations.Keys.ToArray();

        public void SetLanguage(string language)
        {
            if (_translations.ContainsKey(language) && language != _currentLanguage)
            {
                _currentLanguage = language;
                SaveLanguageSetting();
                LanguageChanged?.Invoke(this, language);
            }
        }

        private class LanguageSettings
        {
            public string Language { get; set; }
        }
    }

    // 扩展方法，方便使用
    public static class Lang
    {
        public static string Get(string key) => LanguageService.Instance.GetString(key);
        public static string Get(string key, params object[] args) => LanguageService.Instance.GetString(key, args);
    }
}