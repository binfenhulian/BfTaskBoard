using System;
using System.Drawing;
using System.IO;
using Newtonsoft.Json;

namespace TaskBoard.Services
{
    public class ThemeService
    {
        private static ThemeService _instance;
        private Theme _currentTheme;
        private readonly string _themeSettingPath;

        public static ThemeService Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new ThemeService();
                }
                return _instance;
            }
        }

        public event EventHandler ThemeChanged;

        private ThemeService()
        {
            var appDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "BfTaskBoard");
            Directory.CreateDirectory(appDataFolder);
            _themeSettingPath = Path.Combine(appDataFolder, "theme.json");
            
            LoadThemeSetting();
        }

        private void LoadThemeSetting()
        {
            try
            {
                if (File.Exists(_themeSettingPath))
                {
                    var json = File.ReadAllText(_themeSettingPath);
                    var settings = JsonConvert.DeserializeObject<ThemeSettings>(json);
                    _currentTheme = settings?.Theme ?? Theme.Dark;
                }
                else
                {
                    _currentTheme = Theme.Dark;
                    SaveThemeSetting();
                }
            }
            catch
            {
                _currentTheme = Theme.Dark;
            }
        }

        private void SaveThemeSetting()
        {
            try
            {
                var settings = new ThemeSettings { Theme = _currentTheme };
                var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                File.WriteAllText(_themeSettingPath, json);
            }
            catch { }
        }

        public Theme CurrentTheme => _currentTheme;

        public void SetTheme(Theme theme)
        {
            if (_currentTheme != theme)
            {
                _currentTheme = theme;
                SaveThemeSetting();
                ThemeChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        public ThemeColors GetColors()
        {
            return _currentTheme == Theme.Dark ? DarkThemeColors : LightThemeColors;
        }

        private static readonly ThemeColors DarkThemeColors = new ThemeColors
        {
            // 主背景色
            FormBackground = Color.FromArgb(37, 37, 38),
            PanelBackground = Color.FromArgb(45, 45, 48),
            ControlBackground = Color.FromArgb(30, 30, 30),
            
            // 文本颜色
            Text = Color.White,
            SecondaryText = Color.FromArgb(150, 150, 150),
            DisabledText = Color.FromArgb(100, 100, 100),
            
            // 边框和分隔线
            Border = Color.FromArgb(50, 50, 50),
            GridLines = Color.FromArgb(50, 50, 50),
            
            // 按钮颜色
            ButtonBackground = Color.FromArgb(60, 60, 60),
            ButtonHover = Color.FromArgb(80, 80, 80),
            ButtonPressed = Color.FromArgb(40, 40, 40),
            PrimaryButton = Color.FromArgb(0, 122, 204),
            PrimaryButtonHover = Color.FromArgb(28, 151, 234),
            DangerButton = Color.FromArgb(200, 50, 50),
            
            // 选中和高亮
            Selection = Color.FromArgb(28, 151, 234),
            SelectionText = Color.White,
            Hover = Color.FromArgb(50, 50, 50),
            
            // 输入框
            InputBackground = Color.FromArgb(45, 45, 48),
            InputBorder = Color.FromArgb(60, 60, 60),
            InputFocusBorder = Color.FromArgb(0, 122, 204),
            
            // 标签页
            TabBackground = Color.FromArgb(45, 45, 45),
            TabActiveBackground = Color.FromArgb(30, 30, 30),
            TabHoverBackground = Color.FromArgb(50, 50, 50),
            
            // 菜单
            MenuBackground = Color.FromArgb(37, 37, 38),
            MenuItemHover = Color.FromArgb(50, 50, 50),
            
            // 网格
            GridBackground = Color.FromArgb(37, 37, 38),
            GridAlternateRow = Color.FromArgb(40, 40, 41),
            GridHeader = Color.FromArgb(45, 45, 48),
            
            // 其他
            Success = Color.FromArgb(76, 175, 80),
            Warning = Color.FromArgb(255, 152, 0),
            Error = Color.FromArgb(244, 67, 54),
            Info = Color.FromArgb(33, 150, 243)
        };

        private static readonly ThemeColors LightThemeColors = new ThemeColors
        {
            // 主背景色
            FormBackground = Color.FromArgb(250, 250, 250),
            PanelBackground = Color.White,
            ControlBackground = Color.FromArgb(245, 245, 245),
            
            // 文本颜色
            Text = Color.FromArgb(33, 33, 33),
            SecondaryText = Color.FromArgb(117, 117, 117),
            DisabledText = Color.FromArgb(189, 189, 189),
            
            // 边框和分隔线
            Border = Color.FromArgb(224, 224, 224),
            GridLines = Color.FromArgb(224, 224, 224),
            
            // 按钮颜色
            ButtonBackground = Color.FromArgb(245, 245, 245),
            ButtonHover = Color.FromArgb(229, 229, 229),
            ButtonPressed = Color.FromArgb(214, 214, 214),
            PrimaryButton = Color.FromArgb(25, 118, 210),
            PrimaryButtonHover = Color.FromArgb(33, 150, 243),
            DangerButton = Color.FromArgb(211, 47, 47),
            
            // 选中和高亮
            Selection = Color.FromArgb(33, 150, 243),
            SelectionText = Color.White,
            Hover = Color.FromArgb(245, 245, 245),
            
            // 输入框
            InputBackground = Color.White,
            InputBorder = Color.FromArgb(189, 189, 189),
            InputFocusBorder = Color.FromArgb(33, 150, 243),
            
            // 标签页
            TabBackground = Color.FromArgb(245, 245, 245),
            TabActiveBackground = Color.White,
            TabHoverBackground = Color.FromArgb(238, 238, 238),
            
            // 菜单
            MenuBackground = Color.White,
            MenuItemHover = Color.FromArgb(238, 238, 238),
            
            // 网格
            GridBackground = Color.White,
            GridAlternateRow = Color.FromArgb(250, 250, 250),
            GridHeader = Color.FromArgb(245, 245, 245),
            
            // 其他
            Success = Color.FromArgb(67, 160, 71),
            Warning = Color.FromArgb(251, 140, 0),
            Error = Color.FromArgb(229, 57, 53),
            Info = Color.FromArgb(30, 136, 229)
        };

        private class ThemeSettings
        {
            public Theme Theme { get; set; }
        }
    }

    public enum Theme
    {
        Dark,
        Light
    }

    public class ThemeColors
    {
        // 主背景色
        public Color FormBackground { get; set; }
        public Color PanelBackground { get; set; }
        public Color ControlBackground { get; set; }
        
        // 文本颜色
        public Color Text { get; set; }
        public Color SecondaryText { get; set; }
        public Color DisabledText { get; set; }
        
        // 边框和分隔线
        public Color Border { get; set; }
        public Color GridLines { get; set; }
        
        // 按钮颜色
        public Color ButtonBackground { get; set; }
        public Color ButtonHover { get; set; }
        public Color ButtonPressed { get; set; }
        public Color PrimaryButton { get; set; }
        public Color PrimaryButtonHover { get; set; }
        public Color DangerButton { get; set; }
        
        // 选中和高亮
        public Color Selection { get; set; }
        public Color SelectionText { get; set; }
        public Color Hover { get; set; }
        
        // 输入框
        public Color InputBackground { get; set; }
        public Color InputBorder { get; set; }
        public Color InputFocusBorder { get; set; }
        
        // 标签页
        public Color TabBackground { get; set; }
        public Color TabActiveBackground { get; set; }
        public Color TabHoverBackground { get; set; }
        
        // 菜单
        public Color MenuBackground { get; set; }
        public Color MenuItemHover { get; set; }
        
        // 网格
        public Color GridBackground { get; set; }
        public Color GridAlternateRow { get; set; }
        public Color GridHeader { get; set; }
        
        // 其他
        public Color Success { get; set; }
        public Color Warning { get; set; }
        public Color Error { get; set; }
        public Color Info { get; set; }
    }
}