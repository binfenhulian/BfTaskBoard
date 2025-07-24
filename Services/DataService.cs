using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using TaskBoard.Models;

namespace TaskBoard.Services
{
    public class DataService
    {
        private static DataService _instance;
        private readonly string _dataPath;
        private readonly string _imagePath;
        private AppData _appData;
        private System.Threading.Timer _saveTimer;
        private bool _pendingSave = false;

        public event EventHandler DataChanged;
        
        public static DataService Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new DataService();
                }
                return _instance;
            }
        }

        [Obsolete("Use DataService.Instance instead")]
        public DataService()
        {
            if (_instance == null)
            {
                _instance = this;
            }
            
            var appDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "BfTaskBoard");
            Directory.CreateDirectory(appDataFolder);
            
            _dataPath = Path.Combine(appDataFolder, "data.json");
            _imagePath = Path.Combine(appDataFolder, "images");
            Directory.CreateDirectory(_imagePath);

            LoadData();
        }

        public AppData GetAppData() => _appData;

        public string GetImagePath() => _imagePath;

        public string GetDataPath() => Path.GetDirectoryName(_dataPath);

        private void LoadData()
        {
            if (File.Exists(_dataPath))
            {
                try
                {
                    var json = File.ReadAllText(_dataPath);
                    _appData = JsonConvert.DeserializeObject<AppData>(json) ?? new AppData();
                }
                catch
                {
                    _appData = new AppData();
                }
            }
            else
            {
                _appData = new AppData();
            }
        }

        public void SaveData()
        {
            _pendingSave = true;
            
            if (_saveTimer == null)
            {
                _saveTimer = new System.Threading.Timer(async _ =>
                {
                    if (_pendingSave)
                    {
                        _pendingSave = false;
                        await SaveDataAsync();
                    }
                }, null, 500, Timeout.Infinite);
            }
            else
            {
                _saveTimer.Change(500, Timeout.Infinite);
            }
        }

        private async Task SaveDataAsync()
        {
            try
            {
                // 在后台线程进行序列化，避免阻塞UI
                // 使用 Formatting.None 来压缩JSON，减少文件大小
                var formatting = _appData.Tabs.Count > 10 || 
                                _appData.Tabs.Any(t => t.Rows.Count > 500) 
                                ? Formatting.None  // 数据量大时使用压缩格式
                                : Formatting.Indented; // 数据量小时保持可读性
                                
                var json = await Task.Run(() => 
                    JsonConvert.SerializeObject(_appData, formatting));
                
                // 异步写入文件
                await File.WriteAllTextAsync(_dataPath, json);
                
                // 可选：如果文件很大，考虑写入临时文件后再重命名
                // var tempPath = _dataPath + ".tmp";
                // await File.WriteAllTextAsync(tempPath, json);
                // File.Replace(tempPath, _dataPath, null);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存数据失败: {ex.Message}");
            }
        }

        public void NotifyDataChanged(TabData modifiedTab = null)
        {
            if (modifiedTab != null)
            {
                modifiedTab.LastModified = DateTime.Now;
            }
            DataChanged?.Invoke(this, EventArgs.Empty);
            SaveData();
            
            // 性能监控
            CheckDataFileSize();
        }
        
        private DateTime _lastSizeWarning = DateTime.MinValue;
        
        private void CheckDataFileSize()
        {
            try
            {
                var fileInfo = new FileInfo(_dataPath);
                if (fileInfo.Exists)
                {
                    var sizeMB = fileInfo.Length / (1024.0 * 1024.0);
                    
                    // 如果文件大于10MB，且距离上次警告超过1小时
                    if (sizeMB > 10 && (DateTime.Now - _lastSizeWarning).TotalHours > 1)
                    {
                        _lastSizeWarning = DateTime.Now;
                        
                        var message = LanguageService.Instance.CurrentLanguage == "zh-CN" 
                            ? $"数据文件已达到 {sizeMB:F1}MB，可能会影响性能。\n建议使用\"一键切割\"功能归档旧数据。"
                            : $"Data file has reached {sizeMB:F1}MB, which may affect performance.\nConsider using \"One-Click Cut\" to archive old data.";
                            
                        System.Windows.Forms.MessageBox.Show(message, 
                            LanguageService.Instance.CurrentLanguage == "zh-CN" ? "性能提示" : "Performance Tip",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Information);
                    }
                }
            }
            catch
            {
                // 忽略检查错误
            }
        }

        public void Dispose()
        {
            if (_pendingSave)
            {
                SaveDataAsync().Wait();
            }
            _saveTimer?.Dispose();
        }
    }
}