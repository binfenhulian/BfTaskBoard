using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;
using System.Collections.Generic;
using System.IO;
using TaskBoard.Controls;
using TaskBoard.Models;
using TaskBoard.Services;
using Microsoft.Win32;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Globalization;
using System.Threading.Tasks;
using System.Diagnostics;

namespace TaskBoard
{
    public partial class MainFormNoMaterial : Form
    {
        private DataService _dataService;
        private TabControl _tabControl;
        private Panel _tabPanel;
        private Button _addTabButton;
        private ContextMenuStrip _rowContextMenu;
        private ContextMenuStrip _columnContextMenu;
        private ContextMenuStrip _tabContextMenu;
        private DataGridViewRow _selectedRow;
        private int _selectedColumnIndex;
        private TabPage _selectedTab;
        private bool _todoClickHandled = false;
        private Dictionary<string, bool> _hiddenTabs = new Dictionary<string, bool>();
        private Dictionary<string, List<string>> _columnFilters = new Dictionary<string, List<string>>();
        private int _hoveredColumnIndex = -1;

        public MainFormNoMaterial()
        {
            InitializeComponent();
            InitializeServices();
            SetupAutoStart();
            SetupKeyboardShortcuts();
            
            // 移到 Load 事件中，确保窗口句柄已创建
            this.Load += (s, e) => LoadTabs();
        }

        private void InitializeComponent()
        {
            this.Text = Lang.Get("AppTitle");
            this.Size = new Size(1400, 900);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(30, 30, 30);

            // 创建自定义标签面板
            _tabPanel = new Panel();
            _tabPanel.Height = 40;
            _tabPanel.Dock = DockStyle.Top;
            _tabPanel.BackColor = Color.FromArgb(30, 30, 30);

            _addTabButton = new Button();
            _addTabButton.Text = "+";
            _addTabButton.Size = new Size(30, 30);
            _addTabButton.Location = new Point(5, 5);
            _addTabButton.FlatStyle = FlatStyle.Flat;
            _addTabButton.FlatAppearance.BorderSize = 0;
            _addTabButton.BackColor = Color.FromArgb(45, 45, 45);
            _addTabButton.ForeColor = Color.White;
            _addTabButton.Font = new Font("Arial", 14, FontStyle.Bold);
            _addTabButton.Cursor = Cursors.Hand;
            _addTabButton.Click += OnAddTabClicked;
            _addTabButton.MouseEnter += (s, e) => _addTabButton.BackColor = Color.FromArgb(60, 60, 60);
            _addTabButton.MouseLeave += (s, e) => _addTabButton.BackColor = Color.FromArgb(45, 45, 45);

            _tabPanel.Controls.Add(_addTabButton);

            _tabControl = new TabControl();
            _tabControl.Dock = DockStyle.Fill;
            _tabControl.Appearance = TabAppearance.FlatButtons;
            _tabControl.ItemSize = new Size(0, 1);
            _tabControl.SizeMode = TabSizeMode.Fixed;
            _tabControl.TabStop = false;

            InitializeContextMenus();

            var bottomPanel = new Panel();
            bottomPanel.Height = 30;
            bottomPanel.Dock = DockStyle.Bottom;
            bottomPanel.BackColor = Color.FromArgb(45, 45, 48);

            var openDataFolderButton = new Button();
            openDataFolderButton.Text = Lang.Get("OpenDataFolder");
            openDataFolderButton.Location = new Point(10, 2);
            openDataFolderButton.Size = new Size(120, 26);
            openDataFolderButton.FlatStyle = FlatStyle.Flat;
            openDataFolderButton.BackColor = Color.FromArgb(60, 60, 60);
            openDataFolderButton.ForeColor = Color.White;
            openDataFolderButton.Click += (s, e) => OpenDataFolder();

            bottomPanel.Controls.Add(openDataFolderButton);

            var tabManagementButton = new Button();
            tabManagementButton.Text = Lang.Get("TabManagement");
            tabManagementButton.Location = new Point(140, 2);
            tabManagementButton.Size = new Size(120, 26);
            tabManagementButton.FlatStyle = FlatStyle.Flat;
            tabManagementButton.BackColor = Color.FromArgb(60, 60, 60);
            tabManagementButton.ForeColor = Color.White;
            tabManagementButton.Click += (s, e) => ShowTabManagement();
            bottomPanel.Controls.Add(tabManagementButton);

            var jsonViewButton = new Button();
            jsonViewButton.Text = Lang.Get("JsonView");
            jsonViewButton.Location = new Point(270, 2);
            jsonViewButton.Size = new Size(100, 26);
            jsonViewButton.FlatStyle = FlatStyle.Flat;
            jsonViewButton.BackColor = Color.FromArgb(60, 60, 60);
            jsonViewButton.ForeColor = Color.White;
            jsonViewButton.Click += (s, e) => ShowJsonView();
            bottomPanel.Controls.Add(jsonViewButton);

            var exportExcelButton = new Button();
            exportExcelButton.Text = Lang.Get("ExportExcel");
            exportExcelButton.Location = new Point(380, 2);
            exportExcelButton.Size = new Size(100, 26);
            exportExcelButton.FlatStyle = FlatStyle.Flat;
            exportExcelButton.BackColor = Color.FromArgb(60, 60, 60);
            exportExcelButton.ForeColor = Color.White;
            exportExcelButton.Click += (s, e) => ShowExportExcelDialog();
            bottomPanel.Controls.Add(exportExcelButton);

            var cutButton = new Button();
            cutButton.Text = Lang.Get("OneClickCut");
            cutButton.Location = new Point(490, 2);
            cutButton.Size = new Size(100, 26);
            cutButton.FlatStyle = FlatStyle.Flat;
            cutButton.BackColor = Color.FromArgb(60, 60, 60);
            cutButton.ForeColor = Color.White;
            cutButton.Click += (s, e) => CutAndBackupConfig();
            bottomPanel.Controls.Add(cutButton);

            var restoreButton = new Button();
            restoreButton.Text = Lang.Get("RestoreBackup");
            restoreButton.Location = new Point(590, 2);
            restoreButton.Size = new Size(100, 26);
            restoreButton.FlatStyle = FlatStyle.Flat;
            restoreButton.BackColor = Color.FromArgb(60, 60, 60);
            restoreButton.ForeColor = Color.White;
            restoreButton.Click += (s, e) => ShowRestoreBackupDialog();
            bottomPanel.Controls.Add(restoreButton);

            var searchButton = new Button();
            searchButton.Text = "🔍 " + Lang.Get("Search");
            searchButton.Location = new Point(700, 2);
            searchButton.Size = new Size(100, 26);
            searchButton.FlatStyle = FlatStyle.Flat;
            searchButton.BackColor = Color.FromArgb(60, 60, 60);
            searchButton.ForeColor = Color.White;
            searchButton.Click += (s, e) => ShowSearchDialog();
            bottomPanel.Controls.Add(searchButton);
            
            var aiButton = new Button();
            aiButton.Text = "🤖 " + Lang.Get("AICreateTable");
            aiButton.Location = new Point(810, 2);
            aiButton.Size = new Size(110, 26);
            aiButton.FlatStyle = FlatStyle.Flat;
            aiButton.BackColor = Color.FromArgb(60, 60, 60);
            aiButton.ForeColor = Color.White;
            aiButton.Click += (s, e) => ShowAICreateTableDialog();
            bottomPanel.Controls.Add(aiButton);

            var languageButton = new Button();
            languageButton.Text = Lang.Get("Language") + ": " + (LanguageService.Instance.CurrentLanguage == "zh-CN" ? "中文" : "English");
            languageButton.Location = new Point(930, 2);
            languageButton.Size = new Size(120, 26);
            languageButton.FlatStyle = FlatStyle.Flat;
            languageButton.BackColor = Color.FromArgb(60, 60, 60);
            languageButton.ForeColor = Color.White;
            languageButton.Click += (s, e) => ShowLanguageMenu(languageButton);
            bottomPanel.Controls.Add(languageButton);
            
            // 订阅语言变化事件
            LanguageService.Instance.LanguageChanged += (s, lang) =>
            {
                ApplyLanguage();
                languageButton.Text = Lang.Get("Language") + ": " + (lang == "zh-CN" ? "中文" : "English");
            };
            
            // 主题切换按钮
            var themeButton = new Button();
            themeButton.Text = "🎨 " + (ThemeService.Instance.CurrentTheme == Theme.Dark ? "Dark" : "Light");
            themeButton.Location = new Point(1060, 2);
            themeButton.Size = new Size(100, 26);
            themeButton.FlatStyle = FlatStyle.Flat;
            themeButton.BackColor = Color.FromArgb(60, 60, 60);
            themeButton.ForeColor = Color.White;
            themeButton.Click += (s, e) =>
            {
                var newTheme = ThemeService.Instance.CurrentTheme == Theme.Dark ? Theme.Light : Theme.Dark;
                ThemeService.Instance.SetTheme(newTheme);
                themeButton.Text = "🎨 " + (newTheme == Theme.Dark ? "Dark" : "Light");
            };
            bottomPanel.Controls.Add(themeButton);
            
            // 订阅主题变化事件
            ThemeService.Instance.ThemeChanged += (s, e) =>
            {
                ApplyTheme();
            };

            this.Controls.Add(_tabControl);
            this.Controls.Add(_tabPanel);
            this.Controls.Add(bottomPanel);
            
            // 初始应用主题
            ApplyTheme();
        }

        private void UpdateTabButtons()
        {
            if (!this.IsHandleCreated || !_tabPanel.IsHandleCreated)
            {
                return;
            }

            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(UpdateTabButtons));
                return;
            }

            // 移除旧的标签按钮
            for (int i = _tabPanel.Controls.Count - 1; i >= 0; i--)
            {
                if (_tabPanel.Controls[i] != _addTabButton)
                {
                    _tabPanel.Controls[i].Dispose();
                }
            }

            int x = 5;
            for (int i = 0; i < _tabControl.TabCount; i++)
            {
                var tabData = (TabData)_tabControl.TabPages[i].Tag;
                var tabButton = new TabButton(_tabControl.TabPages[i].Text, i, tabData.Id, tabData.DotColor);
                tabButton.Location = new Point(x, 5);
                tabButton.IsActive = (_tabControl.SelectedIndex == i);

                tabButton.TabClick += TabButton_MouseClick;
                tabButton.TabDoubleClick += TabButton_DoubleClick;
                tabButton.CloseClick += TabButton_CloseClick;
                tabButton.DotClick += TabButton_DotClick;

                _tabPanel.Controls.Add(tabButton);
                x += tabButton.Width + 5;
            }

            _addTabButton.Location = new Point(x, 5);
        }

        private void TabButton_MouseClick(object sender, MouseEventArgs e)
        {
            var button = sender as TabButton;
            if (button == null) return;

            if (e.Button == MouseButtons.Left)
            {
                _tabControl.SelectedIndex = button.TabIndex;
                if (this.IsHandleCreated)
                {
                    this.BeginInvoke(new Action(() => UpdateTabButtons()));
                }
            }
            else if (e.Button == MouseButtons.Right)
            {
                _selectedTab = _tabControl.TabPages[button.TabIndex];
                _tabContextMenu.Show(button, e.Location);
            }
        }

        private void TabButton_DoubleClick(object sender, EventArgs e)
        {
            var button = sender as TabButton;
            if (button == null) return;

            var tabIndex = button.TabIndex;
            var tabPage = _tabControl.TabPages[tabIndex];
            OnTabRenameRequested(tabIndex, tabPage.Text);
        }

        private void TabButton_CloseClick(object sender, EventArgs e)
        {
            var button = sender as TabButton;
            if (button == null) return;

            if (_tabControl.TabCount <= 1)
            {
                MessageBox.Show(Lang.Get("AtLeastOneTab"), Lang.Get("AppTitle"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var tabPage = _tabControl.TabPages[button.TabIndex];
            var tabData = (TabData)tabPage.Tag;

            var dialog = new Form
            {
                Text = "关闭标签页",
                Size = new Size(400, 180),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(37, 37, 38)
            };

            var label = new Label
            {
                Text = $"如何处理标签页 \"{tabData.Name}\"？",
                Location = new Point(20, 20),
                AutoSize = true,
                ForeColor = Color.White
            };

            var hideButton = new Button
            {
                Text = "临时关闭",
                Location = new Point(20, 60),
                Size = new Size(100, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White
            };
            hideButton.Click += (s, args) =>
            {
                // 保存隐藏状态
                _hiddenTabs[tabData.Id] = true;
                SaveHiddenTabsState();
                _tabControl.TabPages.Remove(tabPage);
                dialog.DialogResult = DialogResult.OK;
                dialog.Close();
                if (this.IsHandleCreated)
                {
                    this.BeginInvoke(new Action(() => UpdateTabButtons()));
                }
            };

            var deleteButton = new Button
            {
                Text = "永久删除",
                Location = new Point(140, 60),
                Size = new Size(100, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(200, 50, 50),
                ForeColor = Color.White
            };
            deleteButton.Click += (s, args) =>
            {
                var confirmResult = MessageBox.Show(
                    $"确定要永久删除标签页 \"{tabData.Name}\" 吗？\n此操作不可恢复！",
                    "确认删除",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                
                if (confirmResult == DialogResult.Yes)
                {
                    _dataService.GetAppData().Tabs.Remove(tabData);
                    _tabControl.TabPages.Remove(tabPage);
                    _dataService.NotifyDataChanged();
                    dialog.DialogResult = DialogResult.OK;
                    dialog.Close();
                    if (this.IsHandleCreated)
                    {
                        this.BeginInvoke(new Action(() => UpdateTabButtons()));
                    }
                }
            };

            var cancelButton = new Button
            {
                Text = "取消",
                Location = new Point(260, 60),
                Size = new Size(100, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White
            };
            cancelButton.Click += (s, args) =>
            {
                dialog.DialogResult = DialogResult.Cancel;
                dialog.Close();
            };

            dialog.Controls.Add(label);
            dialog.Controls.Add(hideButton);
            dialog.Controls.Add(deleteButton);
            dialog.Controls.Add(cancelButton);

            ShowModalDialog(dialog);
        }
        
        private void TabButton_DotClick(object sender, EventArgs e)
        {
            var button = sender as TabButton;
            if (button == null) return;
            
            var tabPage = _tabControl.TabPages[button.TabIndex];
            var tabData = (TabData)tabPage.Tag;
            
            // 创建颜色选择对话框
            var dialog = new Form
            {
                Text = "选择标签颜色",
                Size = new Size(280, 320),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(37, 37, 38)
            };
            
            var label = new Label
            {
                Text = $"为 \"{tabData.Name}\" 选择颜色：",
                Location = new Point(20, 15),
                AutoSize = true,
                ForeColor = Color.White
            };
            
            var flowPanel = new FlowLayoutPanel
            {
                Location = new Point(20, 45),
                Size = new Size(220, 180),
                AutoScroll = true,
                BackColor = Color.FromArgb(45, 45, 48),
                BorderStyle = BorderStyle.FixedSingle
            };
            
            // 预设颜色
            var colors = new[]
            {
                "#757575", // 默认灰色
                "#F44336", // 红色
                "#E91E63", // 粉色
                "#9C27B0", // 紫色
                "#673AB7", // 深紫色
                "#3F51B5", // 靛蓝色
                "#2196F3", // 蓝色
                "#03A9F4", // 浅蓝色
                "#00BCD4", // 青色
                "#009688", // 蓝绿色
                "#4CAF50", // 绿色
                "#8BC34A", // 浅绿色
                "#CDDC39", // 黄绿色
                "#FFEB3B", // 黄色
                "#FFC107", // 琥珀色
                "#FF9800", // 橙色
                "#FF5722", // 深橙色
                "#795548", // 棕色
                "#607D8B"  // 蓝灰色
            };
            
            foreach (var colorHex in colors)
            {
                var colorButton = new Button
                {
                    Size = new Size(40, 40),
                    Margin = new Padding(5),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = ColorTranslator.FromHtml(colorHex),
                    Tag = colorHex,
                    Cursor = Cursors.Hand
                };
                colorButton.FlatAppearance.BorderSize = 2;
                colorButton.FlatAppearance.BorderColor = Color.FromArgb(60, 60, 60);
                
                colorButton.Click += (s, args) =>
                {
                    tabData.DotColor = colorHex;
                    _dataService.NotifyDataChanged(tabData);
                    dialog.DialogResult = DialogResult.OK;
                    dialog.Close();
                    if (this.IsHandleCreated)
                    {
                        this.BeginInvoke(new Action(() => UpdateTabButtons()));
                    }
                };
                
                flowPanel.Controls.Add(colorButton);
            }
            
            var cancelButton = new Button
            {
                Text = Lang.Get("Cancel"),
                Location = new Point(85, 240),
                Size = new Size(90, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White
            };
            cancelButton.Click += (s, args) => dialog.Close();
            
            dialog.Controls.Add(label);
            dialog.Controls.Add(flowPanel);
            dialog.Controls.Add(cancelButton);
            
            ShowModalDialog(dialog);
        }

        private void InitializeContextMenus()
        {
            if (_rowContextMenu != null) _rowContextMenu.Dispose();
            if (_columnContextMenu != null) _columnContextMenu.Dispose();
            if (_tabContextMenu != null) _tabContextMenu.Dispose();
            
            _rowContextMenu = new ContextMenuStrip();
            _rowContextMenu.Items.Add(Lang.Get("AddRowAbove"), null, OnAddRowAboveClick);
            _rowContextMenu.Items.Add(Lang.Get("AddRowBelow"), null, OnAddRowBelowClick);
            _rowContextMenu.Items.Add(new ToolStripSeparator());
            _rowContextMenu.Items.Add(Lang.Get("CopyRow"), null, OnCopyRowClick);
            _rowContextMenu.Items.Add(Lang.Get("DeleteRow"), null, OnDeleteRowClick);

            _columnContextMenu = new ContextMenuStrip();
            _columnContextMenu.Items.Add(Lang.Get("AddColumn"), null, OnAddColumnClick);
            _columnContextMenu.Items.Add(Lang.Get("EditColumn"), null, OnEditColumnClick);
            _columnContextMenu.Items.Add(new ToolStripSeparator());
            _columnContextMenu.Items.Add(Lang.Get("SumColumn"), null, OnSumColumnClick);
            _columnContextMenu.Items.Add(new ToolStripSeparator());
            _columnContextMenu.Items.Add(Lang.Get("DeleteColumn"), null, OnDeleteColumnClick);

            _tabContextMenu = new ContextMenuStrip();
            _tabContextMenu.Items.Add(Lang.Get("RenameTab"), null, OnRenameTabClick);
            _tabContextMenu.Items.Add(new ToolStripSeparator());
            _tabContextMenu.Items.Add(Lang.Get("CloseTab"), null, OnCloseTabClick);
        }

        private void SetupAutoStart()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    key?.SetValue("BfTaskBoard", $"\"{Application.ExecutablePath}\"");
                }
            }
            catch { }
        }
        
        private void SetupKeyboardShortcuts()
        {
            this.KeyPreview = true;
            this.KeyDown += (s, e) =>
            {
                // Ctrl+T - 新建标签页
                if (e.Control && e.KeyCode == Keys.T)
                {
                    e.Handled = true;
                    using (var dialog = new SimpleInputDialog(Lang.Get("NewTab"), Lang.Get("EnterTabName"), $"新标签页 {_dataService.GetAppData().Tabs.Count + 1}"))
                    {
                        if (dialog.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.InputValue))
                        {
                            var tabData = new TabData { Name = dialog.InputValue };
                            tabData.Columns.Add(new ColumnDefinition { Name = Lang.Get("Task"), Type = ColumnType.Text });
                            
                            _dataService.GetAppData().Tabs.Add(tabData);
                            CreateTabPage(tabData);
                            _dataService.NotifyDataChanged(tabData);
                            
                            _tabControl.SelectedIndex = _tabControl.TabCount - 1;
                            if (this.IsHandleCreated)
                            {
                                this.BeginInvoke(new Action(() => UpdateTabButtons()));
                            }
                        }
                    }
                }
                // Ctrl+Q - 切换标签页
                else if (e.Control && e.KeyCode == Keys.Q)
                {
                    e.Handled = true;
                    if (_tabControl.TabCount > 1)
                    {
                        var nextIndex = (_tabControl.SelectedIndex + 1) % _tabControl.TabCount;
                        _tabControl.SelectedIndex = nextIndex;
                        if (this.IsHandleCreated)
                        {
                            this.BeginInvoke(new Action(() => UpdateTabButtons()));
                        }
                    }
                }
                // Ctrl+S - 手动保存
                else if (e.Control && e.KeyCode == Keys.S)
                {
                    e.Handled = true;
                    var tabData = GetCurrentTabData();
                    if (tabData != null)
                    {
                        _dataService.NotifyDataChanged(tabData);
                        UpdateTimestampLabel(tabData);
                        
                        // 高亮显示时间戳3秒
                        foreach (TabPage tabPage in _tabControl.TabPages)
                        {
                            if (tabPage.Tag == tabData)
                            {
                                var panel = tabPage.Controls[0] as Panel;
                                if (panel != null)
                                {
                                    var bottomPanel = panel.Controls.OfType<Panel>().FirstOrDefault(p => p.Dock == DockStyle.Bottom);
                                    if (bottomPanel != null)
                                    {
                                        var timestampLabel = bottomPanel.Controls.OfType<Label>().FirstOrDefault(l => l.Tag == tabData);
                                        if (timestampLabel != null)
                                        {
                                            timestampLabel.Text = $"已保存: {tabData.LastModified:yyyy-MM-dd HH:mm:ss}";
                                            timestampLabel.ForeColor = Color.LightGreen;
                                            timestampLabel.Font = new Font(timestampLabel.Font, FontStyle.Bold);
                                            
                                            var timer = new System.Windows.Forms.Timer();
                                            timer.Interval = 3000;
                                            timer.Tick += (sender, args) =>
                                            {
                                                timestampLabel.Text = $"最后修改: {tabData.LastModified:yyyy-MM-dd HH:mm:ss}";
                                                timestampLabel.ForeColor = Color.FromArgb(150, 150, 150);
                                                timestampLabel.Font = new Font("微软雅黑", 9f);
                                                timer.Stop();
                                                timer.Dispose();
                                            };
                                            timer.Start();
                                        }
                                    }
                                }
                                break;
                            }
                        }
                    }
                }
            };
        }

        private void OpenDataFolder()
        {
            var appDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "BfTaskBoard");
            System.Diagnostics.Process.Start("explorer.exe", appDataFolder);
        }

        private void ShowTabManagement()
        {
            var dialog = new Form
            {
                Text = Lang.Get("TabManagementDialog"),
                Size = new Size(600, 500),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(37, 37, 38)
            };

            var listView = new ListView
            {
                Location = new Point(20, 20),
                Size = new Size(450, 350),
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                BackColor = Color.FromArgb(45, 45, 48),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            listView.Columns.Add(Lang.Get("TabName"), 200);
            listView.Columns.Add(Lang.Get("Status"), 100);
            listView.Columns.Add(Lang.Get("ColumnCount"), 60);
            listView.Columns.Add(Lang.Get("RowCount"), 60);

            // 刷新列表
            void RefreshList()
            {
                listView.Items.Clear();
                var appData = _dataService.GetAppData();
                foreach (var tab in appData.Tabs)
                {
                    var item = new ListViewItem(tab.Name);
                    item.Tag = tab;
                    
                    // 状态
                    var isHidden = _hiddenTabs.ContainsKey(tab.Id) && _hiddenTabs[tab.Id];
                    item.SubItems.Add(isHidden ? Lang.Get("Hidden") : Lang.Get("Visible"));
                    
                    // 列数和行数
                    item.SubItems.Add(tab.Columns.Count.ToString());
                    item.SubItems.Add(tab.Rows.Count.ToString());
                    
                    if (isHidden)
                    {
                        item.ForeColor = Color.Gray;
                    }
                    
                    listView.Items.Add(item);
                }
            }

            RefreshList();

            // 按钮面板
            var buttonPanel = new Panel
            {
                Location = new Point(480, 20),
                Size = new Size(100, 350)
            };

            var moveUpButton = new Button
            {
                Text = Lang.Get("MoveUp"),
                Location = new Point(0, 0),
                Size = new Size(100, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White
            };
            moveUpButton.Click += (s, e) =>
            {
                if (listView.SelectedItems.Count > 0)
                {
                    var index = listView.SelectedItems[0].Index;
                    if (index > 0)
                    {
                        var appData = _dataService.GetAppData();
                        var tab = appData.Tabs[index];
                        appData.Tabs.RemoveAt(index);
                        appData.Tabs.Insert(index - 1, tab);
                        _dataService.NotifyDataChanged();
                        RefreshList();
                        ReloadTabs();
                        listView.Items[index - 1].Selected = true;
                    }
                }
            };

            var moveDownButton = new Button
            {
                Text = Lang.Get("MoveDown"),
                Location = new Point(0, 35),
                Size = new Size(100, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White
            };
            moveDownButton.Click += (s, e) =>
            {
                if (listView.SelectedItems.Count > 0)
                {
                    var index = listView.SelectedItems[0].Index;
                    var appData = _dataService.GetAppData();
                    if (index < appData.Tabs.Count - 1)
                    {
                        var tab = appData.Tabs[index];
                        appData.Tabs.RemoveAt(index);
                        appData.Tabs.Insert(index + 1, tab);
                        _dataService.NotifyDataChanged();
                        RefreshList();
                        ReloadTabs();
                        listView.Items[index + 1].Selected = true;
                    }
                }
            };

            var toggleVisibilityButton = new Button
            {
                Text = Lang.Get("Show") + "/" + Lang.Get("Hide"),
                Location = new Point(0, 80),
                Size = new Size(100, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White
            };
            toggleVisibilityButton.Click += (s, e) =>
            {
                if (listView.SelectedItems.Count > 0)
                {
                    var tab = (TabData)listView.SelectedItems[0].Tag;
                    if (_hiddenTabs.ContainsKey(tab.Id) && _hiddenTabs[tab.Id])
                    {
                        _hiddenTabs[tab.Id] = false;
                    }
                    else
                    {
                        _hiddenTabs[tab.Id] = true;
                    }
                    SaveHiddenTabsState();
                    RefreshList();
                    ReloadTabs();
                }
            };

            var addButton = new Button
            {
                Text = "新建标签",
                Location = new Point(0, 125),
                Size = new Size(100, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White
            };
            addButton.Click += (s, e) =>
            {
                using (var inputDialog = new SimpleInputDialog("新建标签页", "请输入标签页名称:", $"新标签页 {_dataService.GetAppData().Tabs.Count + 1}"))
                {
                    if (inputDialog.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(inputDialog.InputValue))
                    {
                        var tabData = new TabData { Name = inputDialog.InputValue };
                        tabData.Columns.Add(new ColumnDefinition { Name = Lang.Get("Task"), Type = ColumnType.Text });
                        
                        _dataService.GetAppData().Tabs.Add(tabData);
                        _dataService.NotifyDataChanged();
                        RefreshList();
                        ReloadTabs();
                    }
                }
            };

            var deleteButton = new Button
            {
                Text = "删除标签",
                Location = new Point(0, 170),
                Size = new Size(100, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(200, 50, 50),
                ForeColor = Color.White
            };
            deleteButton.Click += (s, e) =>
            {
                if (listView.SelectedItems.Count > 0)
                {
                    var tab = (TabData)listView.SelectedItems[0].Tag;
                    var result = MessageBox.Show(
                        $"确定要永久删除标签页 \"{tab.Name}\" 吗？\n此操作不可恢复！",
                        "确认删除",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);
                    
                    if (result == DialogResult.Yes)
                    {
                        _dataService.GetAppData().Tabs.Remove(tab);
                        if (_hiddenTabs.ContainsKey(tab.Id))
                        {
                            _hiddenTabs.Remove(tab.Id);
                            SaveHiddenTabsState();
                        }
                        _dataService.NotifyDataChanged();
                        RefreshList();
                        ReloadTabs();
                    }
                }
            };

            var quickCreateButton = new Button
            {
                Text = Lang.Get("QuickCreate"),
                Location = new Point(0, 215),
                Size = new Size(100, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 150, 136),
                ForeColor = Color.White
            };
            quickCreateButton.Click += (s, e) =>
            {
                if (listView.SelectedItems.Count > 0)
                {
                    var selectedTab = (TabData)listView.SelectedItems[0].Tag;
                    using (var inputDialog = new SimpleInputDialog(Lang.Get("NewTab"), Lang.Get("EnterTabName"), $"{selectedTab.Name} - 副本"))
                    {
                        if (inputDialog.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(inputDialog.InputValue))
                        {
                            var newTab = new TabData { Name = inputDialog.InputValue };
                            
                            // 复制列配置
                            foreach (var col in selectedTab.Columns)
                            {
                                var newCol = new ColumnDefinition
                                {
                                    Name = col.Name,
                                    Type = col.Type,
                                    Options = new List<OptionItem>()
                                };
                                
                                // 复制选项
                                foreach (var opt in col.Options)
                                {
                                    newCol.Options.Add(new OptionItem { Label = opt.Label, Color = opt.Color });
                                }
                                
                                newTab.Columns.Add(newCol);
                            }
                            
                            _dataService.GetAppData().Tabs.Add(newTab);
                            _dataService.NotifyDataChanged();
                            RefreshList();
                            ReloadTabs();
                            
                            // 选中新创建的标签
                            for (int i = 0; i < listView.Items.Count; i++)
                            {
                                if (((TabData)listView.Items[i].Tag).Id == newTab.Id)
                                {
                                    listView.Items[i].Selected = true;
                                    break;
                                }
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show(Lang.Get("SelectTabFirst"), Lang.Get("AppTitle"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };

            buttonPanel.Controls.Add(moveUpButton);
            buttonPanel.Controls.Add(moveDownButton);
            buttonPanel.Controls.Add(toggleVisibilityButton);
            buttonPanel.Controls.Add(addButton);
            buttonPanel.Controls.Add(deleteButton);
            buttonPanel.Controls.Add(quickCreateButton);

            var closeButton = new Button
            {
                Text = Lang.Get("OK"),
                Location = new Point(500, 390),
                Size = new Size(80, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White
            };
            closeButton.Click += (s, e) => dialog.Close();

            dialog.Controls.Add(listView);
            dialog.Controls.Add(buttonPanel);
            dialog.Controls.Add(closeButton);

            ShowModalDialog(dialog);
        }

        private void ReloadTabs()
        {
            // 保存当前选中的标签索引
            var currentIndex = _tabControl.SelectedIndex;
            
            // 清空现有标签页
            _tabControl.TabPages.Clear();
            
            // 重新加载标签页
            LoadTabs();
            
            // 恢复选中状态
            if (currentIndex >= 0 && currentIndex < _tabControl.TabCount)
            {
                _tabControl.SelectedIndex = currentIndex;
            }
        }


        private void ShowJsonView()
        {
            var dialog = new Form
            {
                Text = "JSON视图 - 配置文件编辑器",
                Size = new Size(900, 700),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.Sizable,
                MinimumSize = new Size(600, 400),
                BackColor = Color.FromArgb(30, 30, 30)
            };

            var textBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.White,
                Font = new Font("Consolas", 10),
                BorderStyle = BorderStyle.None
            };

            try
            {
                var appData = _dataService.GetAppData();
                var json = JsonConvert.SerializeObject(appData, Formatting.Indented);
                textBox.Text = json;
                HighlightJson(textBox);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载JSON失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            var panel = new Panel
            {
                Height = 40,
                Dock = DockStyle.Bottom,
                BackColor = Color.FromArgb(45, 45, 48)
            };

            var saveButton = new Button
            {
                Text = "保存",
                Location = new Point(panel.Width - 180, 5),
                Size = new Size(80, 30),
                Anchor = AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White
            };
            saveButton.Click += (s, e) =>
            {
                try
                {
                    // 验证JSON格式
                    var testParse = JsonConvert.DeserializeObject<AppData>(textBox.Text);
                    
                    // 保存到文件
                    var dataFile = Path.Combine(_dataService.GetDataPath(), "data.json");
                    File.WriteAllText(dataFile, textBox.Text);
                    
                    MessageBox.Show("保存成功！需要重启应用程序才能生效。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    dialog.Close();
                }
                catch (JsonException ex)
                {
                    MessageBox.Show($"JSON格式错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"保存失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            var cancelButton = new Button
            {
                Text = "取消",
                Location = new Point(panel.Width - 90, 5),
                Size = new Size(80, 30),
                Anchor = AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White
            };
            cancelButton.Click += (s, e) => dialog.Close();

            panel.Controls.Add(saveButton);
            panel.Controls.Add(cancelButton);

            dialog.Controls.Add(textBox);
            dialog.Controls.Add(panel);

            textBox.TextChanged += (s, e) => 
            {
                // 简单的错误检查
                try
                {
                    JsonConvert.DeserializeObject<AppData>(textBox.Text);
                    saveButton.Enabled = true;
                    saveButton.BackColor = Color.FromArgb(0, 122, 204);
                }
                catch
                {
                    saveButton.Enabled = false;
                    saveButton.BackColor = Color.FromArgb(100, 100, 100);
                }
            };

            ShowModalDialog(dialog);
        }

        private void HighlightJson(RichTextBox rtb)
        {
            // 简单的JSON语法高亮
            rtb.SelectAll();
            rtb.SelectionColor = Color.White;
            
            // 高亮字符串
            HighlightPattern(rtb, "\"[^\"]*\"", Color.FromArgb(214, 157, 133));
            
            // 高亮数字
            HighlightPattern(rtb, @"\b\d+\.?\d*\b", Color.FromArgb(181, 206, 168));
            
            // 高亮布尔值和null
            HighlightPattern(rtb, @"\b(true|false|null)\b", Color.FromArgb(86, 156, 214));
            
            // 高亮属性名
            HighlightPattern(rtb, "\"[^\"]*\"(?=\\s*:)", Color.FromArgb(156, 220, 254));
        }

        private void HighlightPattern(RichTextBox rtb, string pattern, Color color)
        {
            var regex = new System.Text.RegularExpressions.Regex(pattern);
            foreach (System.Text.RegularExpressions.Match match in regex.Matches(rtb.Text))
            {
                rtb.Select(match.Index, match.Length);
                rtb.SelectionColor = color;
            }
            rtb.Select(0, 0);
        }

        private void ShowExportExcelDialog()
        {
            // 设置EPPlus许可证上下文（非商业用途）
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            var dialog = new Form
            {
                Text = Lang.Get("ExportExcelDialog"),
                Size = new Size(500, 400),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(37, 37, 38)
            };

            var label = new Label
            {
                Text = Lang.Get("SelectAtLeastOne") + ":",
                Location = new Point(20, 20),
                AutoSize = true,
                ForeColor = Color.White
            };

            var checkedListBox = new CheckedListBox
            {
                Location = new Point(20, 50),
                Size = new Size(440, 250),
                BackColor = Color.FromArgb(45, 45, 48),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.None,
                CheckOnClick = true
            };

            // 添加所有标签页（包括隐藏的）
            var appData = _dataService.GetAppData();
            foreach (var tab in appData.Tabs)
            {
                var displayName = tab.Name;
                if (_hiddenTabs.ContainsKey(tab.Id) && _hiddenTabs[tab.Id])
                {
                    displayName += " (" + Lang.Get("Hidden") + ")";
                }
                checkedListBox.Items.Add(new TabExportItem { Tab = tab, DisplayName = displayName });
            }

            // 全选/取消全选复选框
            var selectAllCheckBox = new CheckBox
            {
                Text = Lang.Get("SelectAll"),
                Location = new Point(20, 310),
                AutoSize = true,
                ForeColor = Color.White
            };
            selectAllCheckBox.CheckedChanged += (s, e) =>
            {
                for (int i = 0; i < checkedListBox.Items.Count; i++)
                {
                    checkedListBox.SetItemChecked(i, selectAllCheckBox.Checked);
                }
            };

            var exportButton = new Button
            {
                Text = Lang.Get("Export"),
                Location = new Point(300, 340),
                Size = new Size(80, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White
            };

            exportButton.Click += (s, e) =>
            {
                var selectedTabs = new List<TabData>();
                foreach (TabExportItem item in checkedListBox.CheckedItems)
                {
                    selectedTabs.Add(item.Tab);
                }

                if (selectedTabs.Count == 0)
                {
                    MessageBox.Show(Lang.Get("SelectAtLeastOne"), Lang.Get("AppTitle"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 选择保存位置
                using (var saveDialog = new SaveFileDialog())
                {
                    saveDialog.Filter = "Excel文件 (*.xlsx)|*.xlsx";
                    saveDialog.FileName = $"TaskBoard_Export_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                    
                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            ExportToExcel(selectedTabs, saveDialog.FileName);
                            MessageBox.Show(Lang.Get("ExportSuccess", saveDialog.FileName), Lang.Get("AppTitle"), 
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                            dialog.Close();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(Lang.Get("ExportFailed", ex.Message), Lang.Get("AppTitle"), 
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            };

            var cancelButton = new Button
            {
                Text = "取消",
                Location = new Point(390, 340),
                Size = new Size(80, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White
            };
            cancelButton.Click += (s, e) => dialog.Close();

            dialog.Controls.Add(label);
            dialog.Controls.Add(checkedListBox);
            dialog.Controls.Add(selectAllCheckBox);
            dialog.Controls.Add(exportButton);
            dialog.Controls.Add(cancelButton);

            ShowModalDialog(dialog);
        }

        private void ExportToExcel(List<TabData> tabs, string fileName)
        {
            using (var package = new ExcelPackage())
            {
                foreach (var tab in tabs)
                {
                    var worksheet = package.Workbook.Worksheets.Add(tab.Name);
                    
                    // 设置标题行
                    for (int i = 0; i < tab.Columns.Count; i++)
                    {
                        var col = tab.Columns[i];
                        worksheet.Cells[1, i + 1].Value = col.Name;
                        
                        // 设置标题样式
                        using (var headerCell = worksheet.Cells[1, i + 1])
                        {
                            headerCell.Style.Font.Bold = true;
                            headerCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            headerCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(68, 114, 196));
                            headerCell.Style.Font.Color.SetColor(Color.White);
                            headerCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            headerCell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        }
                    }

                    // 添加数据行
                    for (int rowIndex = 0; rowIndex < tab.Rows.Count; rowIndex++)
                    {
                        var row = tab.Rows[rowIndex];
                        for (int colIndex = 0; colIndex < tab.Columns.Count; colIndex++)
                        {
                            var col = tab.Columns[colIndex];
                            var cellValue = "";
                            
                            if (row.Data.ContainsKey(col.Id))
                            {
                                var value = row.Data[col.Id];
                                
                                switch (col.Type)
                                {
                                    case ColumnType.Text:
                                    case ColumnType.Single:
                                    case ColumnType.TextArea:
                                        cellValue = value?.ToString() ?? "";
                                        break;
                                        
                                    case ColumnType.Multi:
                                        if (value is JArray jArray)
                                        {
                                            cellValue = string.Join(", ", jArray.Select(v => v.ToString()));
                                        }
                                        else if (value is List<string> list)
                                        {
                                            cellValue = string.Join(", ", list);
                                        }
                                        break;
                                        
                                    case ColumnType.Image:
                                        if (value is JArray imgArray && imgArray.Count > 0)
                                        {
                                            cellValue = $"[{imgArray.Count}张图片]";
                                        }
                                        else if (value is List<string> imgList && imgList.Count > 0)
                                        {
                                            cellValue = $"[{imgList.Count}张图片]";
                                        }
                                        break;
                                        
                                    case ColumnType.TodoList:
                                        if (value is JArray todoArray)
                                        {
                                            var completedCount = todoArray.Count(t => t["IsCompleted"]?.Value<bool>() == true);
                                            cellValue = $"[{completedCount}/{todoArray.Count}]";
                                            
                                            // 添加详细待办事项作为注释
                                            var todoDetails = string.Join("\n", 
                                                todoArray.Select(t => 
                                                    $"{(t["IsCompleted"]?.Value<bool>() == true ? "✓" : "○")} {t["Text"]?.ToString()}"));
                                            if (!string.IsNullOrEmpty(todoDetails))
                                            {
                                                worksheet.Cells[rowIndex + 2, colIndex + 1].AddComment(todoDetails);
                                            }
                                        }
                                        break;
                                }
                            }
                            
                            worksheet.Cells[rowIndex + 2, colIndex + 1].Value = cellValue;
                            
                            // 设置数据单元格边框
                            worksheet.Cells[rowIndex + 2, colIndex + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        }
                    }

                    // 自动调整列宽
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    
                    // 设置最小和最大列宽
                    for (int i = 1; i <= tab.Columns.Count; i++)
                    {
                        if (worksheet.Column(i).Width < 10)
                            worksheet.Column(i).Width = 10;
                        if (worksheet.Column(i).Width > 50)
                            worksheet.Column(i).Width = 50;
                    }
                }

                // 保存文件
                var fileInfo = new FileInfo(fileName);
                package.SaveAs(fileInfo);
            }
        }

        // 辅助类用于在CheckedListBox中显示标签
        private class TabExportItem
        {
            public TabData Tab { get; set; }
            public string DisplayName { get; set; }
            
            public override string ToString()
            {
                return DisplayName;
            }
        }

        private void CutAndBackupConfig()
        {
            var result = MessageBox.Show(
                Lang.Get("ConfirmCut"), 
                Lang.Get("OneClickCut"), 
                MessageBoxButtons.YesNo, 
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                try
                {
                    // 创建备份目录
                    var backupDir = Path.Combine(_dataService.GetDataPath(), "backups");
                    Directory.CreateDirectory(backupDir);

                    // 生成备份文件名（包含时间戳）
                    var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    var backupFileName = $"backup_{timestamp}.json";
                    var backupPath = Path.Combine(backupDir, backupFileName);

                    // 获取当前配置
                    var currentData = _dataService.GetAppData();
                    var json = JsonConvert.SerializeObject(currentData, Formatting.Indented);
                    
                    // 保存备份
                    File.WriteAllText(backupPath, json);

                    // 创建新的空白配置
                    var newData = new AppData();
                    
                    // 添加一个默认标签页
                    var defaultTab = new TabData { Name = Lang.Get("NewTaskList") };
                    defaultTab.Columns.Add(new ColumnDefinition { Name = Lang.Get("Task"), Type = ColumnType.Text });
                    newData.Tabs.Add(defaultTab);

                    // 更新当前配置为新的空白配置
                    currentData.Tabs.Clear();
                    currentData.Tabs.AddRange(newData.Tabs);
                    
                    // 保存新配置
                    _dataService.NotifyDataChanged();

                    // 清空隐藏标签状态
                    _hiddenTabs.Clear();
                    SaveHiddenTabsState();

                    // 重新加载界面
                    ReloadTabs();

                    MessageBox.Show(
                        Lang.Get("CutSuccess", backupFileName), 
                        Lang.Get("AppTitle"), 
                        MessageBoxButtons.OK, 
                        MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        $"切割失败：{ex.Message}", 
                        "错误", 
                        MessageBoxButtons.OK, 
                        MessageBoxIcon.Error);
                }
            }
        }

        private void ShowRestoreBackupDialog()
        {
            var dialog = new Form
            {
                Text = Lang.Get("RestoreBackupDialog"),
                Size = new Size(600, 450),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(37, 37, 38)
            };

            var label = new Label
            {
                Text = Lang.Get("SelectBackupFile"),
                Location = new Point(20, 20),
                AutoSize = true,
                ForeColor = Color.White
            };

            var listView = new ListView
            {
                Location = new Point(20, 50),
                Size = new Size(560, 300),
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                BackColor = Color.FromArgb(45, 45, 48),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.None
            };

            // 添加列
            listView.Columns.Add(Lang.Get("BackupFile"), 250);
            listView.Columns.Add(Lang.Get("BackupTime"), 150);
            listView.Columns.Add(Lang.Get("FileSize"), 100);
            listView.Columns.Add(Lang.Get("TabCount"), 80);

            // 加载备份文件列表
            var backupDir = Path.Combine(_dataService.GetDataPath(), "backups");
            if (Directory.Exists(backupDir))
            {
                var backupFiles = Directory.GetFiles(backupDir, "backup_*.json")
                    .OrderByDescending(f => File.GetCreationTime(f))
                    .ToArray();

                foreach (var backupFile in backupFiles)
                {
                    var fileInfo = new FileInfo(backupFile);
                    var fileName = Path.GetFileNameWithoutExtension(backupFile);
                    
                    // 尝试解析时间戳
                    var timeStr = fileName.Replace("backup_", "");
                    DateTime backupTime;
                    if (DateTime.TryParseExact(timeStr, "yyyyMMdd_HHmmss", null, 
                        System.Globalization.DateTimeStyles.None, out backupTime))
                    {
                        timeStr = backupTime.ToString("yyyy-MM-dd HH:mm:ss");
                    }

                    // 尝试读取标签页数
                    var tabCount = Lang.Get("Unknown");
                    try
                    {
                        var json = File.ReadAllText(backupFile);
                        var data = JsonConvert.DeserializeObject<AppData>(json);
                        if (data != null)
                        {
                            tabCount = data.Tabs.Count.ToString();
                        }
                    }
                    catch { }

                    var item = new ListViewItem(new[] 
                    { 
                        fileName,
                        timeStr,
                        $"{fileInfo.Length / 1024:N0} KB",
                        tabCount
                    });
                    item.Tag = backupFile;
                    listView.Items.Add(item);
                }
            }

            var restoreButton = new Button
            {
                Text = Lang.Get("Restore"),
                Location = new Point(420, 370),
                Size = new Size(80, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                Enabled = false
            };

            listView.SelectedIndexChanged += (s, e) =>
            {
                restoreButton.Enabled = listView.SelectedItems.Count > 0;
            };

            restoreButton.Click += (s, e) =>
            {
                if (listView.SelectedItems.Count == 0) return;

                var selectedFile = (string)listView.SelectedItems[0].Tag;
                var fileName = listView.SelectedItems[0].Text;

                var confirmResult = MessageBox.Show(
                    Lang.Get("ConfirmRestore", fileName),
                    Lang.Get("RestoreBackup"),
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (confirmResult == DialogResult.Yes)
                {
                    try
                    {
                        // 先备份当前配置
                        var backupDir = Path.Combine(_dataService.GetDataPath(), "backups");
                        Directory.CreateDirectory(backupDir);
                        
                        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                        var currentBackupPath = Path.Combine(backupDir, $"backup_before_restore_{timestamp}.json");
                        
                        var currentData = _dataService.GetAppData();
                        var currentJson = JsonConvert.SerializeObject(currentData, Formatting.Indented);
                        File.WriteAllText(currentBackupPath, currentJson);

                        // 读取并恢复选中的备份
                        var backupJson = File.ReadAllText(selectedFile);
                        var backupData = JsonConvert.DeserializeObject<AppData>(backupJson);

                        if (backupData != null)
                        {
                            // 清空当前数据
                            currentData.Tabs.Clear();
                            
                            // 恢复备份数据
                            currentData.Tabs.AddRange(backupData.Tabs);
                            
                            // 保存
                            _dataService.NotifyDataChanged();

                            // 清空隐藏标签状态
                            _hiddenTabs.Clear();
                            SaveHiddenTabsState();

                            // 重新加载界面
                            ReloadTabs();

                            dialog.Close();

                            MessageBox.Show(
                                Lang.Get("RestoreSuccess", fileName, $"backup_before_restore_{timestamp}.json"),
                                Lang.Get("AppTitle"),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            $"恢复失败：{ex.Message}",
                            "错误",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                }
            };

            var cancelButton = new Button
            {
                Text = "取消",
                Location = new Point(510, 370),
                Size = new Size(70, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White
            };
            cancelButton.Click += (s, e) => dialog.Close();

            dialog.Controls.Add(label);
            dialog.Controls.Add(listView);
            dialog.Controls.Add(restoreButton);
            dialog.Controls.Add(cancelButton);

            ShowModalDialog(dialog);
        }

        private void InitializeServices()
        {
            _dataService = DataService.Instance;
            this.FormClosing += (s, e) => _dataService.Dispose();
        }

        private void ShowLanguageMenu(Button languageButton)
        {
            var contextMenu = new ContextMenuStrip();
            
            var chineseItem = new ToolStripMenuItem("中文");
            chineseItem.Checked = LanguageService.Instance.CurrentLanguage == "zh-CN";
            chineseItem.Click += (s, e) => LanguageService.Instance.SetLanguage("zh-CN");
            
            var englishItem = new ToolStripMenuItem("English");
            englishItem.Checked = LanguageService.Instance.CurrentLanguage == "en-US";
            englishItem.Click += (s, e) => LanguageService.Instance.SetLanguage("en-US");
            
            contextMenu.Items.Add(chineseItem);
            contextMenu.Items.Add(englishItem);
            
            contextMenu.Show(languageButton, new Point(0, languageButton.Height));
        }

        private void ApplyLanguage()
        {
            // 更新窗口标题
            this.Text = Lang.Get("AppTitle");
            
            // 更新底部按钮的文本
            var bottomPanel = this.Controls.OfType<Panel>().FirstOrDefault(p => p.Dock == DockStyle.Bottom);
            if (bottomPanel != null)
            {
                var buttons = bottomPanel.Controls.OfType<Button>().ToList();
                if (buttons.Count >= 8)
                {
                    buttons[0].Text = Lang.Get("OpenDataFolder");
                    buttons[1].Text = Lang.Get("TabManagement");
                    buttons[2].Text = Lang.Get("JsonView");
                    buttons[3].Text = Lang.Get("ExportExcel");
                    buttons[4].Text = Lang.Get("OneClickCut");
                    buttons[5].Text = Lang.Get("RestoreBackup");
                    buttons[6].Text = "🔍 " + Lang.Get("Search");
                    // buttons[7] 是语言按钮，已经在事件中更新
                }
            }
            
            // 重新初始化右键菜单
            InitializeContextMenus();
            
            // 如果需要完全刷新界面，可以调用ReloadTabs()
            // 但这会导致用户丢失当前的编辑状态，所以暂时不这样做
        }
        
        private void ApplyTheme()
        {
            var colors = ThemeService.Instance.GetColors();
            
            // 应用主窗体颜色
            this.BackColor = colors.FormBackground;
            
            // 应用标签面板颜色
            _tabPanel.BackColor = colors.ControlBackground;
            
            // 应用底部面板颜色
            var bottomPanel = this.Controls.OfType<Panel>().FirstOrDefault(p => p.Dock == DockStyle.Bottom);
            if (bottomPanel != null)
            {
                bottomPanel.BackColor = colors.PanelBackground;
                foreach (var button in bottomPanel.Controls.OfType<Button>())
                {
                    button.BackColor = colors.ButtonBackground;
                    button.ForeColor = colors.Text;
                }
            }
            
            // 更新标签按钮
            UpdateTabButtons();
            
            // 更新所有标签页
            foreach (TabPage tabPage in _tabControl.TabPages)
            {
                tabPage.BackColor = colors.FormBackground;
                
                var panel = tabPage.Controls[0] as Panel;
                if (panel != null)
                {
                    panel.BackColor = colors.FormBackground;
                    
                    // 更新网格
                    var grid = panel.Controls.OfType<DataGridView>().FirstOrDefault();
                    if (grid != null)
                    {
                        grid.BackgroundColor = colors.GridBackground;
                        grid.GridColor = colors.GridLines;
                        
                        // 默认单元格样式
                        grid.DefaultCellStyle.BackColor = colors.GridBackground;
                        grid.DefaultCellStyle.ForeColor = colors.Text;
                        grid.DefaultCellStyle.SelectionBackColor = colors.Selection;
                        grid.DefaultCellStyle.SelectionForeColor = colors.SelectionText;
                        
                        // 列标题样式
                        grid.ColumnHeadersDefaultCellStyle.BackColor = colors.GridHeader;
                        grid.ColumnHeadersDefaultCellStyle.ForeColor = colors.Text;
                        grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = colors.GridHeader;
                        grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = colors.Text;
                        
                        // 行标题样式
                        grid.RowHeadersDefaultCellStyle.BackColor = colors.GridHeader;
                        grid.RowHeadersDefaultCellStyle.ForeColor = colors.Text;
                        grid.RowHeadersDefaultCellStyle.SelectionBackColor = colors.Selection;
                        grid.RowHeadersDefaultCellStyle.SelectionForeColor = colors.SelectionText;
                        
                        // 刷新网格显示
                        grid.Invalidate();
                    }
                    
                    // 更新底部面板
                    var bottomTabPanel = panel.Controls.OfType<Panel>().FirstOrDefault(p => p.Dock == DockStyle.Bottom);
                    if (bottomTabPanel != null)
                    {
                        bottomTabPanel.BackColor = colors.FormBackground;
                        foreach (var control in bottomTabPanel.Controls)
                        {
                            if (control is Button btn)
                            {
                                btn.BackColor = colors.ButtonBackground;
                                btn.ForeColor = colors.Text;
                            }
                            else if (control is Label lbl)
                            {
                                lbl.ForeColor = colors.SecondaryText;
                            }
                        }
                    }
                }
            }
            
            // 更新右键菜单
            UpdateContextMenusTheme();
        }
        
        private void UpdateContextMenusTheme()
        {
            var colors = ThemeService.Instance.GetColors();
            
            // 更新所有右键菜单的颜色
            var menus = new[] { _rowContextMenu, _columnContextMenu, _tabContextMenu };
            foreach (var menu in menus)
            {
                if (menu != null)
                {
                    menu.BackColor = colors.MenuBackground;
                    menu.ForeColor = colors.Text;
                    foreach (ToolStripItem item in menu.Items)
                    {
                        if (item is ToolStripMenuItem menuItem)
                        {
                            menuItem.BackColor = colors.MenuBackground;
                            menuItem.ForeColor = colors.Text;
                        }
                    }
                }
            }
        }

        private void ShowSearchDialog()
        {
            var dialog = new Form
            {
                Text = Lang.Get("SearchDialog"),
                Size = new Size(800, 600),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.Sizable,
                MinimumSize = new Size(600, 400),
                BackColor = Color.FromArgb(37, 37, 38)
            };

            // 搜索输入区域
            var searchPanel = new Panel
            {
                Height = 60,
                Dock = DockStyle.Top,
                BackColor = Color.FromArgb(45, 45, 48),
                Padding = new Padding(10)
            };

            var searchTextBox = new TextBox
            {
                Location = new Point(10, 20),
                Width = 500,
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Microsoft YaHei UI", 10)
            };

            var searchButton = new Button
            {
                Text = Lang.Get("Search"),
                Location = new Point(520, 18),
                Size = new Size(80, 26),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White
            };

            searchPanel.Controls.Add(searchTextBox);
            searchPanel.Controls.Add(searchButton);

            // 结果显示区域
            var resultPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(30, 30, 30),
                Padding = new Padding(10),
                AutoScroll = true
            };

            var resultLabel = new Label
            {
                Text = Lang.Get("EnterSearchKeyword"),
                Location = new Point(10, 10),
                AutoSize = true,
                ForeColor = Color.Gray
            };

            resultPanel.Controls.Add(resultLabel);

            // 搜索功能
            Action performSearch = () =>
            {
                var keyword = searchTextBox.Text.Trim();
                if (string.IsNullOrEmpty(keyword))
                {
                    resultLabel.Text = Lang.Get("EnterSearchKeyword");
                    resultLabel.ForeColor = Color.Gray;
                    return;
                }

                resultPanel.Controls.Clear();
                var results = SearchInAllTabs(keyword);
                
                if (results.Count == 0)
                {
                    resultLabel.Text = Lang.Get("NoSearchResult");
                    resultLabel.ForeColor = Color.OrangeRed;
                    resultPanel.Controls.Add(resultLabel);
                }
                else
                {
                    var tabCount = results.Select(r => r.TabName).Distinct().Count();
                    resultLabel.Text = Lang.Get("FoundInTabs", tabCount);
                    resultLabel.ForeColor = Color.LightGreen;
                    resultPanel.Controls.Add(resultLabel);

                    int y = 40;
                    foreach (var result in results)
                    {
                        var resultItem = new Panel
                        {
                            Location = new Point(10, y),
                            Size = new Size(resultPanel.Width - 40, 60),
                            BackColor = Color.FromArgb(45, 45, 48),
                            Cursor = Cursors.Hand
                        };

                        var tabLabel = new Label
                        {
                            Text = $"📑 {result.TabName}",
                            Location = new Point(10, 5),
                            AutoSize = true,
                            ForeColor = Color.White,
                            Font = new Font("Microsoft YaHei UI", 9, FontStyle.Bold)
                        };

                        var locationLabel = new Label
                        {
                            Text = $"列: {result.ColumnName}, 行: {result.RowIndex + 1}",
                            Location = new Point(10, 25),
                            AutoSize = true,
                            ForeColor = Color.LightGray
                        };

                        var valueLabel = new Label
                        {
                            Text = result.Value.Length > 50 ? result.Value.Substring(0, 50) + "..." : result.Value,
                            Location = new Point(10, 42),
                            AutoSize = true,
                            ForeColor = Color.Silver,
                            Font = new Font("Microsoft YaHei UI", 8)
                        };

                        resultItem.Controls.Add(tabLabel);
                        resultItem.Controls.Add(locationLabel);
                        resultItem.Controls.Add(valueLabel);

                        // 点击定位功能
                        resultItem.Click += (s, e) => NavigateToResult(result, dialog);
                        tabLabel.Click += (s, e) => NavigateToResult(result, dialog);
                        locationLabel.Click += (s, e) => NavigateToResult(result, dialog);
                        valueLabel.Click += (s, e) => NavigateToResult(result, dialog);

                        // 悬浮效果
                        resultItem.MouseEnter += (s, e) => resultItem.BackColor = Color.FromArgb(60, 60, 60);
                        resultItem.MouseLeave += (s, e) => resultItem.BackColor = Color.FromArgb(45, 45, 48);

                        resultPanel.Controls.Add(resultItem);
                        y += 70;
                    }
                }
            };

            searchButton.Click += (s, e) => performSearch();
            searchTextBox.KeyPress += (s, e) =>
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    performSearch();
                    e.Handled = true;
                }
            };

            dialog.Controls.Add(resultPanel);
            dialog.Controls.Add(searchPanel);

            searchTextBox.Focus();
            ShowModalDialog(dialog);
        }

        private class SearchResult
        {
            public string TabId { get; set; }
            public string TabName { get; set; }
            public string ColumnId { get; set; }
            public string ColumnName { get; set; }
            public int RowIndex { get; set; }
            public string Value { get; set; }
        }

        private List<SearchResult> SearchInAllTabs(string keyword)
        {
            var results = new List<SearchResult>();
            var appData = _dataService.GetAppData();
            keyword = keyword.ToLower();

            foreach (var tab in appData.Tabs)
            {
                for (int rowIndex = 0; rowIndex < tab.Rows.Count; rowIndex++)
                {
                    var row = tab.Rows[rowIndex];
                    foreach (var col in tab.Columns)
                    {
                        if (row.Data.ContainsKey(col.Id))
                        {
                            var value = row.Data[col.Id]?.ToString() ?? "";
                            if (value.ToLower().Contains(keyword))
                            {
                                results.Add(new SearchResult
                                {
                                    TabId = tab.Id,
                                    TabName = tab.Name,
                                    ColumnId = col.Id,
                                    ColumnName = col.Name,
                                    RowIndex = rowIndex,
                                    Value = value
                                });
                            }
                        }
                    }
                }
            }

            return results;
        }

        private void NavigateToResult(SearchResult result, Form searchDialog)
        {
            // 先关闭搜索对话框
            searchDialog.Close();

            // 找到对应的标签页
            var tabIndex = -1;
            for (int i = 0; i < _tabControl.TabPages.Count; i++)
            {
                var tabData = (TabData)_tabControl.TabPages[i].Tag;
                if (tabData.Id == result.TabId)
                {
                    tabIndex = i;
                    break;
                }
            }

            if (tabIndex >= 0)
            {
                // 切换到对应标签页
                _tabControl.SelectedIndex = tabIndex;

                // 获取当前的DataGridView
                var grid = GetCurrentGrid();
                if (grid != null && result.RowIndex < grid.Rows.Count)
                {
                    // 清除当前选择
                    grid.ClearSelection();
                    
                    // 选中对应的单元格
                    var colIndex = -1;
                    for (int i = 0; i < grid.Columns.Count; i++)
                    {
                        var col = grid.Columns[i].Tag as ColumnDefinition;
                        if (col != null && col.Id == result.ColumnId)
                        {
                            colIndex = i;
                            break;
                        }
                    }

                    if (colIndex >= 0)
                    {
                        grid.CurrentCell = grid.Rows[result.RowIndex].Cells[colIndex];
                        grid.Rows[result.RowIndex].Selected = true;
                        
                        // 确保单元格可见
                        grid.FirstDisplayedScrollingRowIndex = Math.Max(0, result.RowIndex - 5);
                    }
                }
            }
        }

        private void Grid_ColumnHeaderCellPainting_Handler(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex == -1 && e.ColumnIndex >= 0)
            {
                var grid = (DataGridView)sender;
                var col = grid.Columns[e.ColumnIndex].Tag as ColumnDefinition;
                
                if (col != null && col.Type == ColumnType.Single)
                {
                    // 默认绘制
                    e.Paint(e.CellBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.ContentForeground);
                    
                    // 绘制文本
                    var textBounds = e.CellBounds;
                    textBounds.Width -= 20; // 为图标留出空间
                    TextRenderer.DrawText(e.Graphics, e.Value?.ToString(), e.CellStyle.Font,
                        textBounds, e.CellStyle.ForeColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                    
                    // 如果鼠标悬浮或有筛选，显示筛选图标
                    var filterKey = $"{((TabData)_tabControl.SelectedTab.Tag).Id}_{col.Id}";
                    var hasFilter = _columnFilters.ContainsKey(filterKey) && _columnFilters[filterKey].Count > 0;
                    
                    if (_hoveredColumnIndex == e.ColumnIndex || hasFilter)
                    {
                        var iconBounds = new Rectangle(e.CellBounds.Right - 18, 
                            e.CellBounds.Top + (e.CellBounds.Height - 16) / 2, 16, 16);
                        
                        // 绘制筛选图标（简单的漏斗形状）
                        using (var pen = new Pen(hasFilter ? Color.FromArgb(0, 122, 204) : Color.Gray, 2))
                        {
                            e.Graphics.DrawLine(pen, iconBounds.Left + 2, iconBounds.Top + 2, 
                                iconBounds.Right - 2, iconBounds.Top + 2);
                            e.Graphics.DrawLine(pen, iconBounds.Left + 2, iconBounds.Top + 2, 
                                iconBounds.Left + 8, iconBounds.Bottom - 2);
                            e.Graphics.DrawLine(pen, iconBounds.Right - 2, iconBounds.Top + 2, 
                                iconBounds.Right - 8, iconBounds.Bottom - 2);
                        }
                    }
                    
                    e.Handled = true;
                }
                else if (col != null && col.Type == ColumnType.Text)
                {
                    // 为文本列绘制排序图标
                    e.Paint(e.CellBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.ContentForeground);
                    
                    // 绘制文本
                    var textBounds = e.CellBounds;
                    textBounds.Width -= 20; // 为图标留出空间
                    TextRenderer.DrawText(e.Graphics, e.Value?.ToString(), e.CellStyle.Font,
                        textBounds, e.CellStyle.ForeColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                    
                    if (_hoveredColumnIndex == e.ColumnIndex)
                    {
                        var iconBounds = new Rectangle(e.CellBounds.Right - 18, 
                            e.CellBounds.Top + (e.CellBounds.Height - 16) / 2, 16, 16);
                        
                        // 绘制排序图标（上下箭头）
                        using (var pen = new Pen(Color.Gray, 2))
                        {
                            // 向上箭头
                            e.Graphics.DrawLine(pen, iconBounds.Left + 8, iconBounds.Top + 3, 
                                iconBounds.Left + 4, iconBounds.Top + 7);
                            e.Graphics.DrawLine(pen, iconBounds.Left + 8, iconBounds.Top + 3, 
                                iconBounds.Left + 12, iconBounds.Top + 7);
                            // 向下箭头
                            e.Graphics.DrawLine(pen, iconBounds.Left + 8, iconBounds.Bottom - 3, 
                                iconBounds.Left + 4, iconBounds.Bottom - 7);
                            e.Graphics.DrawLine(pen, iconBounds.Left + 8, iconBounds.Bottom - 3, 
                                iconBounds.Left + 12, iconBounds.Bottom - 7);
                        }
                    }
                    
                    e.Handled = true;
                }
            }
        }

        private void Grid_CellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex == -1 && e.ColumnIndex >= 0)
            {
                Grid_ColumnHeaderMouseMove_Handler(sender, e);
            }
        }
        
        private void Grid_ColumnHeaderMouseMove_Handler(object sender, DataGridViewCellMouseEventArgs e)
        {
            var grid = (DataGridView)sender;
            
            if (e.RowIndex == -1 && e.ColumnIndex >= 0)
            {
                var col = grid.Columns[e.ColumnIndex].Tag as ColumnDefinition;
                if (col != null && (col.Type == ColumnType.Single || col.Type == ColumnType.Text))
                {
                    if (_hoveredColumnIndex != e.ColumnIndex)
                    {
                        var oldIndex = _hoveredColumnIndex;
                        _hoveredColumnIndex = e.ColumnIndex;
                        
                        // 刷新旧列头和新列头
                        if (oldIndex >= 0 && oldIndex < grid.ColumnCount)
                        {
                            grid.InvalidateColumn(oldIndex);
                        }
                        grid.InvalidateColumn(_hoveredColumnIndex);
                    }
                    return;
                }
            }
            
            if (_hoveredColumnIndex >= 0)
            {
                var oldIndex = _hoveredColumnIndex;
                _hoveredColumnIndex = -1;
                if (oldIndex < grid.ColumnCount)
                {
                    grid.InvalidateColumn(oldIndex);
                }
            }
        }

        private void LoadTabs()
        {
            var appData = _dataService.GetAppData();
            
            if (appData.Tabs.Count == 0)
            {
                CreateDefaultTab();
            }

            // 加载保存的隐藏状态
            LoadHiddenTabsState();

            foreach (var tabData in appData.Tabs)
            {
                // 跳过隐藏的标签页
                if (_hiddenTabs.ContainsKey(tabData.Id) && _hiddenTabs[tabData.Id])
                {
                    continue;
                }
                CreateTabPage(tabData);
            }

            if (_tabControl.TabCount > 0)
            {
                _tabControl.SelectedIndex = 0;
            }

            // 延迟更新标签按钮，确保控件已创建
            if (this.IsHandleCreated)
            {
                this.BeginInvoke(new Action(() => UpdateTabButtons()));
            }
            else
            {
                this.HandleCreated += (s, e) => UpdateTabButtons();
            }
        }

        private void LoadHiddenTabsState()
        {
            try
            {
                var hiddenTabsFile = Path.Combine(_dataService.GetDataPath(), "hidden_tabs.json");
                if (File.Exists(hiddenTabsFile))
                {
                    var json = File.ReadAllText(hiddenTabsFile);
                    _hiddenTabs = JsonConvert.DeserializeObject<Dictionary<string, bool>>(json) ?? new Dictionary<string, bool>();
                }
            }
            catch { }
        }

        private void SaveHiddenTabsState()
        {
            try
            {
                var hiddenTabsFile = Path.Combine(_dataService.GetDataPath(), "hidden_tabs.json");
                var json = JsonConvert.SerializeObject(_hiddenTabs, Formatting.Indented);
                File.WriteAllText(hiddenTabsFile, json);
            }
            catch { }
        }

        private void CreateDefaultTab()
        {
            var defaultTab = new TabData { Name = Lang.Get("DailyPlan") };
            defaultTab.Columns.Add(new ColumnDefinition { Name = Lang.Get("Task"), Type = ColumnType.Text });
            defaultTab.Columns.Add(new ColumnDefinition 
            { 
                Name = Lang.Get("Status"), 
                Type = ColumnType.Single,
                Options = new List<OptionItem>
                {
                    new OptionItem { Label = Lang.Get("NotStarted"), Color = "#757575" },
                    new OptionItem { Label = Lang.Get("InProgress"), Color = "#FF6F00" },
                    new OptionItem { Label = Lang.Get("Completed"), Color = "#388E3C" }
                }
            });
            
            _dataService.GetAppData().Tabs.Add(defaultTab);
            _dataService.SaveData();
        }

        private void CreateTabPage(TabData tabData)
        {
            var tabPage = new TabPage(tabData.Name);
            tabPage.Tag = tabData;
            tabPage.BackColor = Color.FromArgb(37, 37, 38);

            DataGridView grid = new ModernDataGridView();
            grid.Dock = DockStyle.Fill;
            grid.Margin = new Padding(10);
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.CellValueChanged += Grid_CellValueChanged;
            grid.CellMouseDown += Grid_CellMouseDown;
            grid.CellClick += Grid_CellClick;
            grid.MouseClick += Grid_MouseClick;
            grid.ColumnHeaderMouseClick += Grid_ColumnHeaderMouseClick;
            grid.CellPainting += Grid_CellPainting;
            grid.CellMouseMove += Grid_CellMouseMove;
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(45, 45, 48);
            grid.DefaultCellStyle.SelectionForeColor = Color.White;
            // 行高将根据内容动态调整

            var panel = new Panel();
            panel.Dock = DockStyle.Fill;
            panel.Padding = new Padding(10);
            panel.BackColor = Color.FromArgb(37, 37, 38);

            var bottomPanel = new Panel();
            bottomPanel.Height = 40;
            bottomPanel.Dock = DockStyle.Bottom;
            bottomPanel.BackColor = Color.FromArgb(37, 37, 38);

            var addRowButton = new Button();
            addRowButton.Text = "添加新行";
            addRowButton.Location = new Point(10, 5);
            addRowButton.FlatStyle = FlatStyle.Flat;
            addRowButton.BackColor = Color.FromArgb(60, 60, 60);
            addRowButton.ForeColor = Color.White;
            addRowButton.Click += (s, e) => AddNewRow(tabData, grid);
            bottomPanel.Controls.Add(addRowButton);
            
            // 添加最后修改时间标签
            var timestampLabel = new Label();
            timestampLabel.Text = $"最后修改: {tabData.LastModified:yyyy-MM-dd HH:mm:ss}";
            timestampLabel.AutoSize = true;
            timestampLabel.Location = new Point(bottomPanel.Width - 200, 10);
            timestampLabel.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            timestampLabel.ForeColor = Color.FromArgb(150, 150, 150);
            timestampLabel.Font = new Font("微软雅黑", 9f);
            timestampLabel.Tag = tabData; // 保存引用以便更新
            bottomPanel.Controls.Add(timestampLabel);

            UpdateGridColumns(grid, tabData);
            UpdateGridRows(grid, tabData);

            panel.Controls.Add(grid);
            panel.Controls.Add(bottomPanel);
            tabPage.Controls.Add(panel);
            _tabControl.TabPages.Add(tabPage);
        }

        private void Grid_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            var grid = (DataGridView)sender;
            
            // 处理列标题绘制
            if (e.RowIndex == -1 && e.ColumnIndex >= 0)
            {
                Grid_ColumnHeaderCellPainting_Handler(sender, e);
                return;
            }
            
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var col = grid.Columns[e.ColumnIndex].Tag as ColumnDefinition;
            if (col == null) return;

            if (col.Type == ColumnType.Single)
            {
                e.Paint(e.CellBounds, DataGridViewPaintParts.Background | DataGridViewPaintParts.Border);
                e.Handled = true;

                var value = e.FormattedValue?.ToString();
                if (string.IsNullOrEmpty(value)) return;

                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;

                var option = col.Options.FirstOrDefault(o => o.Label == value);
                if (option == null) return;

                var textSize = g.MeasureString(value, grid.Font);
                var tagWidth = (int)textSize.Width + 16;
                var tagHeight = 20;
                var x = e.CellBounds.X + (e.CellBounds.Width - tagWidth) / 2;
                var y = e.CellBounds.Y + (e.CellBounds.Height - tagHeight) / 2;
                var tagRect = new Rectangle(x, y, tagWidth, tagHeight);

                using (var brush = new SolidBrush(ColorTranslator.FromHtml(option.Color)))
                using (var path = GetRoundedRectPath(tagRect, 10))
                {
                    g.FillPath(brush, path);
                }

                using (var brush = new SolidBrush(GetContrastColor(ColorTranslator.FromHtml(option.Color))))
                {
                    var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                    g.DrawString(value, grid.Font, brush, tagRect, format);
                }
            }
            else if (col.Type == ColumnType.Image)
            {
                e.Paint(e.CellBounds, DataGridViewPaintParts.Background | DataGridViewPaintParts.Border);
                e.Handled = true;

                var rowData = grid.Rows[e.RowIndex].Tag as RowData;
                if (rowData != null && rowData.Data.ContainsKey(col.Id))
                {
                    var images = rowData.Data[col.Id] as List<string> ?? new List<string>();
                    var count = images.Count;

                    var g = e.Graphics;
                    var text = count > 0 ? $"📷 {count} 张图片" : "📷 添加图片";
                    var textSize = g.MeasureString(text, grid.Font);
                    var textRect = new RectangleF(
                        e.CellBounds.X + (e.CellBounds.Width - textSize.Width) / 2,
                        e.CellBounds.Y + (e.CellBounds.Height - textSize.Height) / 2,
                        textSize.Width,
                        textSize.Height);

                    using (var brush = new SolidBrush(count > 0 ? Color.LightBlue : Color.Gray))
                    {
                        g.DrawString(text, grid.Font, brush, textRect);
                    }
                }
            }
            else if (col.Type == ColumnType.TodoList)
            {
                e.Paint(e.CellBounds, DataGridViewPaintParts.Background | DataGridViewPaintParts.Border);
                e.Handled = true;

                var rowData = grid.Rows[e.RowIndex].Tag as RowData;
                if (rowData != null && rowData.Data.ContainsKey(col.Id))
                {
                    var todos = GetTodoListFromData(rowData.Data[col.Id]);
                    if (todos.Count > 0)
                    {
                        var g = e.Graphics;
                        g.SmoothingMode = SmoothingMode.AntiAlias;
                        var y = e.CellBounds.Y + 5;

                        foreach (var todo in todos)
                        {
                            // 绘制圆点
                            var dotRect = new Rectangle(e.CellBounds.X + 5, y + 2, 12, 12);
                            using (var brush = new SolidBrush(todo.IsCompleted ? Color.FromArgb(0, 200, 0) : Color.FromArgb(200, 200, 200)))
                            {
                                g.FillEllipse(brush, dotRect);
                            }

                            // 绘制文本
                            var textRect = new Rectangle(e.CellBounds.X + 22, y, e.CellBounds.Width - 25, 16);
                            var font = todo.IsCompleted ? new Font(grid.Font, FontStyle.Strikeout) : grid.Font;
                            var textColor = todo.IsCompleted ? Color.Gray : Color.White;
                            using (var brush = new SolidBrush(textColor))
                            {
                                var format = new StringFormat { Trimming = StringTrimming.EllipsisCharacter };
                                g.DrawString(todo.Text, font, brush, textRect, format);
                            }

                            y += 18;
                        }
                    }
                }
            }
            else if (col.Type == ColumnType.TextArea)
            {
                e.Paint(e.CellBounds, DataGridViewPaintParts.Background | DataGridViewPaintParts.Border);
                e.Handled = true;

                var value = e.FormattedValue?.ToString();
                if (!string.IsNullOrEmpty(value))
                {
                    var g = e.Graphics;
                    // 只显示第一行，超出部分显示省略号
                    var lines = value.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    var displayText = lines.Length > 0 ? lines[0] : "";
                    
                    if (lines.Length > 1)
                    {
                        displayText += "...";
                    }
                    
                    var textRect = new Rectangle(e.CellBounds.X + 5, e.CellBounds.Y, e.CellBounds.Width - 10, e.CellBounds.Height);
                    var format = new StringFormat 
                    { 
                        LineAlignment = StringAlignment.Center,
                        Trimming = StringTrimming.EllipsisCharacter
                    };
                    
                    using (var brush = new SolidBrush(e.State.HasFlag(DataGridViewElementStates.Selected) ? Color.White : grid.DefaultCellStyle.ForeColor))
                    {
                        g.DrawString(displayText, grid.Font, brush, textRect, format);
                    }
                }
            }
        }

        private GraphicsPath GetRoundedRectPath(Rectangle rect, int radius)
        {
            var path = new GraphicsPath();
            path.AddArc(rect.X, rect.Y, radius, radius, 180, 90);
            path.AddArc(rect.Right - radius, rect.Y, radius, radius, 270, 90);
            path.AddArc(rect.Right - radius, rect.Bottom - radius, radius, radius, 0, 90);
            path.AddArc(rect.X, rect.Bottom - radius, radius, radius, 90, 90);
            path.CloseFigure();
            return path;
        }

        private Color GetContrastColor(Color color)
        {
            var luminance = (0.299 * color.R + 0.587 * color.G + 0.114 * color.B) / 255;
            return luminance > 0.5 ? Color.Black : Color.White;
        }

        private List<TodoItem> GetTodoListFromData(object data)
        {
            if (data == null) return new List<TodoItem>();

            if (data is List<TodoItem> todoList)
                return todoList;

            if (data is JArray jArray)
            {
                return jArray.ToObject<List<TodoItem>>() ?? new List<TodoItem>();
            }

            return new List<TodoItem>();
        }

        private void UpdateGridColumns(DataGridView grid, TabData tabData)
        {
            grid.Columns.Clear();

            foreach (var col in tabData.Columns)
            {
                DataGridViewColumn gridCol = null;

                switch (col.Type)
                {
                    case ColumnType.Text:
                        gridCol = new DataGridViewTextBoxColumn();
                        break;
                    case ColumnType.Single:
                        gridCol = new DataGridViewTextBoxColumn();
                        gridCol.ReadOnly = true;
                        break;
                    case ColumnType.Multi:
                        continue; // 跳过多选类型
                    case ColumnType.Image:
                        gridCol = new DataGridViewTextBoxColumn();
                        gridCol.ReadOnly = true;
                        break;
                    case ColumnType.TodoList:
                        gridCol = new DataGridViewTextBoxColumn();
                        gridCol.ReadOnly = true;
                        break;
                    case ColumnType.TextArea:
                        gridCol = new DataGridViewTextBoxColumn();
                        gridCol.ReadOnly = true;
                        break;
                }

                if (gridCol != null)
                {
                    gridCol.Name = col.Id;
                    gridCol.HeaderText = col.Name;
                    gridCol.Tag = col;
                    gridCol.SortMode = DataGridViewColumnSortMode.NotSortable;
                    gridCol.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    grid.Columns.Add(gridCol);
                }
            }
        }

        private void UpdateGridRows(DataGridView grid, TabData tabData)
        {
            grid.Rows.Clear();

            foreach (var rowData in tabData.Rows)
            {
                var row = grid.Rows[grid.Rows.Add()];
                row.Tag = rowData;

                // 计算行高
                int maxHeight = 30; // 默认最小高度

                foreach (var col in tabData.Columns)
                {
                    if (col.Type == ColumnType.Multi) continue; // 跳过多选类型
                    
                    if (rowData.Data.ContainsKey(col.Id))
                    {
                        var value = rowData.Data[col.Id];
                        
                        if (col.Type == ColumnType.TodoList)
                        {
                            // 计算TodoList所需高度
                            var todos = GetTodoListFromData(value);
                            if (todos.Count > 0)
                            {
                                // 每个待办项占用18像素，加上顶部和底部边距
                                int todoHeight = (todos.Count * 18) + 10;
                                maxHeight = Math.Max(maxHeight, todoHeight);
                            }
                        }
                        else if (col.Type == ColumnType.Image)
                        {
                            // 图片列不显示文本，由 CellPainting 处理
                        }
                        else if (col.Type != ColumnType.Image)
                        {
                            row.Cells[col.Id].Value = value?.ToString();
                        }
                    }
                }

                // 设置行高
                row.Height = maxHeight;
            }
        }

        private void Grid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            var grid = (DataGridView)sender;
            var tabData = GetCurrentTabData();
            var rowData = (RowData)grid.Rows[e.RowIndex].Tag;
            var colId = grid.Columns[e.ColumnIndex].Name;

            rowData.Data[colId] = grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            _dataService.NotifyDataChanged(tabData);
            UpdateTimestampLabel(tabData);
        }

        private void Grid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            // 如果是TodoList的圆点点击已经处理过，则跳过
            if (_todoClickHandled)
            {
                _todoClickHandled = false;
                return;
            }

            var grid = (DataGridView)sender;
            var col = (ColumnDefinition)grid.Columns[e.ColumnIndex].Tag;

            if (col.Type == ColumnType.Image)
            {
                var rowData = (RowData)grid.Rows[e.RowIndex].Tag;
                ShowImageManagementDialog(rowData, col.Id, grid, e.RowIndex);
            }
            else if (col.Type == ColumnType.Single)
            {
                var rowData = (RowData)grid.Rows[e.RowIndex].Tag;
                ShowSingleSelectDialog(rowData, col, grid, e.RowIndex, e.ColumnIndex);
            }
            else if (col.Type == ColumnType.TodoList)
            {
                var rowData = (RowData)grid.Rows[e.RowIndex].Tag;
                ShowTodoListDialog(rowData, col.Id, grid, e.RowIndex);
            }
            else if (col.Type == ColumnType.TextArea)
            {
                var rowData = (RowData)grid.Rows[e.RowIndex].Tag;
                ShowTextAreaDialog(rowData, col.Id, col.Name, grid, e.RowIndex, e.ColumnIndex);
            }
        }

        private void ShowModalDialog(Form dialog)
        {
            // 应用主题到对话框
            ApplyThemeToDialog(dialog);
            
            // 创建半透明蒙层
            var overlay = new Form();
            overlay.StartPosition = FormStartPosition.Manual;
            overlay.FormBorderStyle = FormBorderStyle.None;
            overlay.Bounds = this.Bounds;
            overlay.BackColor = Color.Black;
            overlay.Opacity = 0.5;
            overlay.ShowInTaskbar = false;
            overlay.Show(this);

            dialog.StartPosition = FormStartPosition.CenterParent;
            dialog.ShowDialog(this);
            
            overlay.Close();
            overlay.Dispose();
        }
        
        private void ApplyThemeToDialog(Form dialog)
        {
            var colors = ThemeService.Instance.GetColors();
            
            dialog.BackColor = colors.FormBackground;
            dialog.ForeColor = colors.Text;
            
            ApplyThemeToControls(dialog.Controls, colors);
        }
        
        private void ApplyThemeToControls(Control.ControlCollection controls, ThemeColors colors)
        {
            foreach (Control control in controls)
            {
                if (control is Button button)
                {
                    if (button.BackColor == Color.FromArgb(0, 122, 204) || button.Text.Contains("确定") || button.Text.Contains("OK"))
                    {
                        button.BackColor = colors.PrimaryButton;
                    }
                    else if (button.BackColor == Color.FromArgb(200, 50, 50) || button.Text.Contains("删除"))
                    {
                        button.BackColor = colors.DangerButton;
                    }
                    else
                    {
                        button.BackColor = colors.ButtonBackground;
                    }
                    button.ForeColor = colors.Text;
                    button.FlatAppearance.BorderColor = colors.Border;
                }
                else if (control is Label label)
                {
                    label.ForeColor = colors.Text;
                }
                else if (control is TextBox textBox)
                {
                    textBox.BackColor = colors.InputBackground;
                    textBox.ForeColor = colors.Text;
                }
                else if (control is ComboBox comboBox)
                {
                    comboBox.BackColor = colors.InputBackground;
                    comboBox.ForeColor = colors.Text;
                }
                else if (control is ListBox listBox)
                {
                    listBox.BackColor = colors.InputBackground;
                    listBox.ForeColor = colors.Text;
                }
                else if (control is Panel panel)
                {
                    panel.BackColor = colors.PanelBackground;
                    if (control.Controls.Count > 0)
                    {
                        ApplyThemeToControls(control.Controls, colors);
                    }
                }
                else if (control is FlowLayoutPanel flowPanel)
                {
                    flowPanel.BackColor = colors.PanelBackground;
                    if (control.Controls.Count > 0)
                    {
                        ApplyThemeToControls(control.Controls, colors);
                    }
                }
                else if (control is DataGridView grid)
                {
                    grid.BackgroundColor = colors.GridBackground;
                    grid.GridColor = colors.GridLines;
                    grid.DefaultCellStyle.BackColor = colors.GridBackground;
                    grid.DefaultCellStyle.ForeColor = colors.Text;
                    grid.DefaultCellStyle.SelectionBackColor = colors.Selection;
                    grid.DefaultCellStyle.SelectionForeColor = colors.SelectionText;
                    grid.ColumnHeadersDefaultCellStyle.BackColor = colors.GridHeader;
                    grid.ColumnHeadersDefaultCellStyle.ForeColor = colors.Text;
                }
                else if (control is ListView listView)
                {
                    listView.BackColor = colors.InputBackground;
                    listView.ForeColor = colors.Text;
                }
                else if (control is RichTextBox richTextBox)
                {
                    richTextBox.BackColor = colors.InputBackground;
                    richTextBox.ForeColor = colors.Text;
                }
                
                // 递归处理子控件
                if (control.Controls.Count > 0 && !(control is Panel) && !(control is FlowLayoutPanel))
                {
                    ApplyThemeToControls(control.Controls, colors);
                }
            }
        }

        private void ShowSingleSelectDialog(RowData rowData, ColumnDefinition col, DataGridView grid, int rowIndex, int colIndex)
        {
            var dialog = new Form();
            dialog.Text = $"选择{col.Name}";
            dialog.Size = new Size(450, 500);
            dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
            dialog.MaximizeBox = false;
            dialog.MinimizeBox = false;
            dialog.BackColor = Color.FromArgb(37, 37, 38);

            // 主面板
            var mainPanel = new Panel();
            mainPanel.Location = new Point(20, 20);
            mainPanel.Size = new Size(410, 360);
            mainPanel.BackColor = Color.FromArgb(45, 45, 48);
            
            // 列表框
            var listBox = new ListBox();
            listBox.Location = new Point(0, 0);
            listBox.Size = new Size(320, 360);
            listBox.BackColor = Color.FromArgb(45, 45, 48);
            listBox.ForeColor = Color.White;
            listBox.BorderStyle = BorderStyle.None;
            listBox.Font = new Font("Microsoft YaHei UI", 10);
            
            var currentValue = rowData.Data.ContainsKey(col.Id) ? rowData.Data[col.Id]?.ToString() : "";
            
            // 操作按钮面板
            var buttonPanel = new Panel();
            buttonPanel.Location = new Point(320, 0);
            buttonPanel.Size = new Size(90, 360);
            buttonPanel.BackColor = Color.FromArgb(37, 37, 38);
            
            var addButton = new Button();
            addButton.Text = "添加";
            addButton.Location = new Point(10, 10);
            addButton.Size = new Size(70, 30);
            addButton.FlatStyle = FlatStyle.Flat;
            addButton.BackColor = Color.FromArgb(0, 122, 204);
            addButton.ForeColor = Color.White;
            
            var editButton = new Button();
            editButton.Text = "编辑";
            editButton.Location = new Point(10, 50);
            editButton.Size = new Size(70, 30);
            editButton.FlatStyle = FlatStyle.Flat;
            editButton.BackColor = Color.FromArgb(60, 60, 60);
            editButton.ForeColor = Color.White;
            editButton.Enabled = false;
            
            var deleteButton = new Button();
            deleteButton.Text = "删除";
            deleteButton.Location = new Point(10, 90);
            deleteButton.Size = new Size(70, 30);
            deleteButton.FlatStyle = FlatStyle.Flat;
            deleteButton.BackColor = Color.FromArgb(60, 60, 60);
            deleteButton.ForeColor = Color.White;
            deleteButton.Enabled = false;
            
            buttonPanel.Controls.Add(addButton);
            buttonPanel.Controls.Add(editButton);
            buttonPanel.Controls.Add(deleteButton);
            
            mainPanel.Controls.Add(listBox);
            mainPanel.Controls.Add(buttonPanel);
            
            // 刷新列表
            Action refreshList = () =>
            {
                listBox.Items.Clear();
                int selectedIndex = -1;
                for (int i = 0; i < col.Options.Count; i++)
                {
                    listBox.Items.Add(col.Options[i]);
                    if (col.Options[i].Label == currentValue)
                    {
                        selectedIndex = i;
                    }
                }
                if (selectedIndex >= 0)
                {
                    listBox.SelectedIndex = selectedIndex;
                }
            };
            
            refreshList();
            
            // 自定义绘制列表项
            listBox.DrawMode = DrawMode.OwnerDrawFixed;
            listBox.ItemHeight = 35;
            listBox.DrawItem += (s, e) =>
            {
                if (e.Index < 0) return;
                
                var g = e.Graphics;
                var isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
                var option = col.Options[e.Index];
                
                // 背景
                var bgColor = isSelected ? Color.FromArgb(0, 122, 204) : Color.FromArgb(45, 45, 48);
                g.FillRectangle(new SolidBrush(bgColor), e.Bounds);
                
                // 颜色标签
                var tagRect = new Rectangle(e.Bounds.X + 10, e.Bounds.Y + 7, 20, 20);
                using (var brush = new SolidBrush(ColorTranslator.FromHtml(option.Color)))
                using (var path = GetRoundedRectPath(tagRect, 10))
                {
                    g.FillPath(brush, path);
                }
                
                // 文本
                var textRect = new Rectangle(e.Bounds.X + 40, e.Bounds.Y, e.Bounds.Width - 40, e.Bounds.Height);
                var format = new StringFormat { LineAlignment = StringAlignment.Center };
                g.DrawString(option.Label, listBox.Font, new SolidBrush(Color.White), textRect, format);
            };
            
            // 选择改变时启用/禁用按钮
            listBox.SelectedIndexChanged += (s, e) =>
            {
                editButton.Enabled = listBox.SelectedIndex >= 0;
                deleteButton.Enabled = listBox.SelectedIndex >= 0;
            };
            
            // 添加选项
            addButton.Click += (s, e) =>
            {
                ShowOptionEditDialog(null, (newOption) =>
                {
                    col.Options.Add(newOption);
                    refreshList();
                    var tabData = GetCurrentTabData();
                    _dataService.NotifyDataChanged(tabData);
                    UpdateTimestampLabel(tabData);
                });
            };
            
            // 编辑选项
            editButton.Click += (s, e) =>
            {
                if (listBox.SelectedIndex >= 0)
                {
                    var option = col.Options[listBox.SelectedIndex];
                    ShowOptionEditDialog(option, (editedOption) =>
                    {
                        option.Label = editedOption.Label;
                        option.Color = editedOption.Color;
                        refreshList();
                        // 如果当前单元格的值是这个选项，更新显示
                        if (rowData.Data.ContainsKey(col.Id) && rowData.Data[col.Id]?.ToString() == option.Label)
                        {
                            grid.InvalidateCell(colIndex, rowIndex);
                        }
                        var tabData = GetCurrentTabData();
                        _dataService.NotifyDataChanged(tabData);
                        UpdateTimestampLabel(tabData);
                    });
                }
            };
            
            // 删除选项
            deleteButton.Click += (s, e) =>
            {
                if (listBox.SelectedIndex >= 0)
                {
                    var option = col.Options[listBox.SelectedIndex];
                    var result = MessageBox.Show($"确定要删除选项 \"{option.Label}\" 吗？", "确认删除", 
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    
                    if (result == DialogResult.Yes)
                    {
                        col.Options.RemoveAt(listBox.SelectedIndex);
                        refreshList();
                        var tabData = GetCurrentTabData();
                        _dataService.NotifyDataChanged(tabData);
                        UpdateTimestampLabel(tabData);
                    }
                }
            };
            
            // 双击选择
            listBox.DoubleClick += (s, e) =>
            {
                if (listBox.SelectedIndex >= 0)
                {
                    var selectedOption = col.Options[listBox.SelectedIndex];
                    rowData.Data[col.Id] = selectedOption.Label;
                    grid.Rows[rowIndex].Cells[colIndex].Value = selectedOption.Label;
                    var tabData = GetCurrentTabData();
                    _dataService.NotifyDataChanged(tabData);
                    UpdateTimestampLabel(tabData);
                    grid.InvalidateCell(colIndex, rowIndex);
                    dialog.Close();
                }
            };
            
            // 底部按钮
            var okButton = new Button();
            okButton.Text = "确定";
            okButton.Location = new Point(120, 400);
            okButton.Size = new Size(90, 35);
            okButton.FlatStyle = FlatStyle.Flat;
            okButton.BackColor = Color.FromArgb(0, 122, 204);
            okButton.ForeColor = Color.White;
            okButton.Click += (s, e) =>
            {
                if (listBox.SelectedIndex >= 0)
                {
                    var selectedOption = col.Options[listBox.SelectedIndex];
                    rowData.Data[col.Id] = selectedOption.Label;
                    grid.Rows[rowIndex].Cells[colIndex].Value = selectedOption.Label;
                    var tabData = GetCurrentTabData();
                    _dataService.NotifyDataChanged(tabData);
                    UpdateTimestampLabel(tabData);
                    grid.InvalidateCell(colIndex, rowIndex);
                }
                dialog.Close();
            };
            
            var cancelButton = new Button();
            cancelButton.Text = "取消";
            cancelButton.Location = new Point(240, 400);
            cancelButton.Size = new Size(90, 35);
            cancelButton.FlatStyle = FlatStyle.Flat;
            cancelButton.BackColor = Color.FromArgb(60, 60, 60);
            cancelButton.ForeColor = Color.White;
            cancelButton.Click += (s, e) => dialog.Close();

            dialog.Controls.Add(mainPanel);
            dialog.Controls.Add(okButton);
            dialog.Controls.Add(cancelButton);

            ShowModalDialog(dialog);
        }
        
        private void ShowOptionEditDialog(OptionItem existingOption, Action<OptionItem> onSave)
        {
            var dialog = new Form();
            dialog.Text = existingOption == null ? "添加选项" : "编辑选项";
            dialog.Size = new Size(350, 250);
            dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
            dialog.MaximizeBox = false;
            dialog.MinimizeBox = false;
            dialog.BackColor = Color.FromArgb(37, 37, 38);
            dialog.StartPosition = FormStartPosition.CenterParent;
            
            var nameLabel = new Label();
            nameLabel.Text = "选项名称：";
            nameLabel.Location = new Point(20, 20);
            nameLabel.AutoSize = true;
            nameLabel.ForeColor = Color.White;
            
            var nameTextBox = new TextBox();
            nameTextBox.Location = new Point(20, 45);
            nameTextBox.Size = new Size(290, 25);
            nameTextBox.BackColor = Color.FromArgb(45, 45, 48);
            nameTextBox.ForeColor = Color.White;
            nameTextBox.BorderStyle = BorderStyle.FixedSingle;
            nameTextBox.Text = existingOption?.Label ?? "";
            
            var colorLabel = new Label();
            colorLabel.Text = "背景颜色：";
            colorLabel.Location = new Point(20, 80);
            colorLabel.AutoSize = true;
            colorLabel.ForeColor = Color.White;
            
            var colorPanel = new Panel();
            colorPanel.Location = new Point(20, 105);
            colorPanel.Size = new Size(290, 40);
            colorPanel.BackColor = existingOption != null ? ColorTranslator.FromHtml(existingOption.Color) : Color.FromArgb(100, 100, 100);
            colorPanel.BorderStyle = BorderStyle.FixedSingle;
            colorPanel.Cursor = Cursors.Hand;
            
            var selectedColor = existingOption?.Color ?? "#646464";
            
            colorPanel.Click += (s, e) =>
            {
                // 显示颜色选择器
                var colorDialog = new Form();
                colorDialog.Text = "选择颜色";
                colorDialog.Size = new Size(280, 320);
                colorDialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                colorDialog.MaximizeBox = false;
                colorDialog.MinimizeBox = false;
                colorDialog.BackColor = Color.FromArgb(37, 37, 38);
                colorDialog.StartPosition = FormStartPosition.CenterParent;
                
                var colorFlowPanel = new FlowLayoutPanel();
                colorFlowPanel.Location = new Point(20, 20);
                colorFlowPanel.Size = new Size(220, 200);
                colorFlowPanel.AutoScroll = true;
                colorFlowPanel.BackColor = Color.FromArgb(45, 45, 48);
                colorFlowPanel.BorderStyle = BorderStyle.FixedSingle;
                
                // 预设颜色
                var colors = new[]
                {
                    "#F44336", "#E91E63", "#9C27B0", "#673AB7",
                    "#3F51B5", "#2196F3", "#03A9F4", "#00BCD4",
                    "#009688", "#4CAF50", "#8BC34A", "#CDDC39",
                    "#FFEB3B", "#FFC107", "#FF9800", "#FF5722",
                    "#795548", "#9E9E9E", "#607D8B", "#455A64",
                    "#37474F", "#263238", "#212121", "#757575"
                };
                
                foreach (var colorHex in colors)
                {
                    var colorButton = new Button();
                    colorButton.Size = new Size(40, 40);
                    colorButton.Margin = new Padding(5);
                    colorButton.FlatStyle = FlatStyle.Flat;
                    colorButton.BackColor = ColorTranslator.FromHtml(colorHex);
                    colorButton.Tag = colorHex;
                    colorButton.Cursor = Cursors.Hand;
                    colorButton.FlatAppearance.BorderSize = 2;
                    colorButton.FlatAppearance.BorderColor = Color.FromArgb(60, 60, 60);
                    
                    colorButton.Click += (sender, args) =>
                    {
                        selectedColor = colorHex;
                        colorPanel.BackColor = ColorTranslator.FromHtml(colorHex);
                        colorDialog.Close();
                    };
                    
                    colorFlowPanel.Controls.Add(colorButton);
                }
                
                var colorCancelButton = new Button();
                colorCancelButton.Text = "取消";
                colorCancelButton.Location = new Point(85, 240);
                colorCancelButton.Size = new Size(90, 30);
                colorCancelButton.FlatStyle = FlatStyle.Flat;
                colorCancelButton.BackColor = Color.FromArgb(60, 60, 60);
                colorCancelButton.ForeColor = Color.White;
                colorCancelButton.Click += (sender, args) => colorDialog.Close();
                
                colorDialog.Controls.Add(colorFlowPanel);
                colorDialog.Controls.Add(colorCancelButton);
                
                colorDialog.ShowDialog(dialog);
            };
            
            var okButton = new Button();
            okButton.Text = "确定";
            okButton.Location = new Point(65, 165);
            okButton.Size = new Size(90, 35);
            okButton.FlatStyle = FlatStyle.Flat;
            okButton.BackColor = Color.FromArgb(0, 122, 204);
            okButton.ForeColor = Color.White;
            okButton.Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(nameTextBox.Text))
                {
                    MessageBox.Show("请输入选项名称", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                var option = new OptionItem
                {
                    Label = nameTextBox.Text.Trim(),
                    Color = selectedColor
                };
                
                onSave(option);
                dialog.Close();
            };
            
            var cancelButton = new Button();
            cancelButton.Text = "取消";
            cancelButton.Location = new Point(185, 165);
            cancelButton.Size = new Size(90, 35);
            cancelButton.FlatStyle = FlatStyle.Flat;
            cancelButton.BackColor = Color.FromArgb(60, 60, 60);
            cancelButton.ForeColor = Color.White;
            cancelButton.Click += (s, e) => dialog.Close();
            
            dialog.Controls.Add(nameLabel);
            dialog.Controls.Add(nameTextBox);
            dialog.Controls.Add(colorLabel);
            dialog.Controls.Add(colorPanel);
            dialog.Controls.Add(okButton);
            dialog.Controls.Add(cancelButton);
            
            nameTextBox.Focus();
            dialog.ShowDialog();
        }

        private void ShowImageManagementDialog(RowData rowData, string colId, DataGridView grid, int rowIndex)
        {
            var dialog = new Form();
            dialog.Text = "图片管理";
            dialog.Size = new Size(800, 600);
            dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
            dialog.MaximizeBox = false;
            dialog.MinimizeBox = false;
            dialog.BackColor = Color.FromArgb(37, 37, 38);

            var flowPanel = new FlowLayoutPanel();
            flowPanel.Dock = DockStyle.Fill;
            flowPanel.AutoScroll = true;
            flowPanel.Padding = new Padding(10);
            flowPanel.BackColor = Color.FromArgb(37, 37, 38);

            if (!rowData.Data.ContainsKey(colId))
            {
                rowData.Data[colId] = new List<string>();
            }

            var images = rowData.Data[colId] as List<string> ?? new List<string>();

            void RefreshImages()
            {
                flowPanel.Controls.Clear();

                foreach (var imagePath in images)
                {
                    var imagePanel = new Panel();
                    imagePanel.Size = new Size(150, 150);
                    imagePanel.BackColor = Color.FromArgb(45, 45, 48);
                    imagePanel.Margin = new Padding(5);

                    var picBox = new PictureBox();
                    picBox.Size = new Size(140, 140);
                    picBox.Location = new Point(5, 5);
                    picBox.SizeMode = PictureBoxSizeMode.Zoom;
                    picBox.Cursor = Cursors.Hand;

                    try
                    {
                        var fullPath = Path.Combine(_dataService.GetImagePath(), imagePath);
                        if (File.Exists(fullPath))
                        {
                            using (var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read))
                            {
                                picBox.Image = Image.FromStream(stream);
                            }
                        }
                    }
                    catch { }

                    picBox.Click += (s, e) =>
                    {
                        var viewer = new Form();
                        viewer.Text = "查看图片";
                        viewer.Size = new Size(800, 600);
                        viewer.StartPosition = FormStartPosition.CenterParent;
                        viewer.BackColor = Color.Black;
                        
                        var viewPic = new PictureBox();
                        viewPic.Dock = DockStyle.Fill;
                        viewPic.SizeMode = PictureBoxSizeMode.Zoom;
                        viewPic.Image = picBox.Image;
                        viewPic.BackColor = Color.Black;
                        
                        viewer.Controls.Add(viewPic);
                        ShowModalDialog(viewer);
                    };

                    var deleteBtn = new Button();
                    deleteBtn.Text = "×";
                    deleteBtn.Size = new Size(25, 25);
                    deleteBtn.Location = new Point(120, 5);
                    deleteBtn.FlatStyle = FlatStyle.Flat;
                    deleteBtn.BackColor = Color.FromArgb(200, 50, 50);
                    deleteBtn.ForeColor = Color.White;
                    deleteBtn.Cursor = Cursors.Hand;
                    deleteBtn.Click += (s, e) =>
                    {
                        images.Remove(imagePath);
                        RefreshImages();
                    };

                    imagePanel.Controls.Add(picBox);
                    imagePanel.Controls.Add(deleteBtn);
                    flowPanel.Controls.Add(imagePanel);
                }

                var addButton = new Button();
                addButton.Size = new Size(150, 150);
                addButton.BackColor = Color.FromArgb(45, 45, 48);
                addButton.ForeColor = Color.White;
                addButton.Text = "+\n添加图片\n(支持Ctrl+V)";
                addButton.Font = new Font("Arial", 16);
                addButton.FlatStyle = FlatStyle.Flat;
                addButton.FlatAppearance.BorderSize = 1;
                addButton.FlatAppearance.BorderColor = Color.FromArgb(70, 70, 70);
                addButton.Cursor = Cursors.Hand;
                addButton.Margin = new Padding(5);
                
                addButton.Click += (s, e) =>
                {
                    using (var openDialog = new OpenFileDialog())
                    {
                        openDialog.Filter = "图片文件|*.jpg;*.jpeg;*.png;*.gif;*.bmp";
                        openDialog.Multiselect = true;
                        if (openDialog.ShowDialog() == DialogResult.OK)
                        {
                            foreach (var file in openDialog.FileNames)
                            {
                                var fileName = $"{Guid.NewGuid():N}{Path.GetExtension(file)}";
                                var destPath = Path.Combine(_dataService.GetImagePath(), fileName);
                                File.Copy(file, destPath, true);
                                images.Add(fileName);
                            }
                            RefreshImages();
                        }
                    }
                };

                flowPanel.Controls.Add(addButton);
            }

            RefreshImages();

            var bottomPanel = new Panel();
            bottomPanel.Height = 50;
            bottomPanel.Dock = DockStyle.Bottom;
            bottomPanel.BackColor = Color.FromArgb(45, 45, 48);

            var okButton = new Button();
            okButton.Text = "确定";
            okButton.Size = new Size(80, 30);
            okButton.Location = new Point(630, 10);
            okButton.FlatStyle = FlatStyle.Flat;
            okButton.BackColor = Color.FromArgb(0, 122, 204);
            okButton.ForeColor = Color.White;
            okButton.Click += (s, e) =>
            {
                rowData.Data[colId] = images;
                var tabData = GetCurrentTabData();
                _dataService.NotifyDataChanged(tabData);
                UpdateTimestampLabel(tabData);
                grid.InvalidateCell(grid.Columns[colId].Index, rowIndex);
                dialog.Close();
            };

            var cancelButton = new Button();
            cancelButton.Text = "取消";
            cancelButton.Size = new Size(80, 30);
            cancelButton.Location = new Point(540, 10);
            cancelButton.FlatStyle = FlatStyle.Flat;
            cancelButton.BackColor = Color.FromArgb(60, 60, 60);
            cancelButton.ForeColor = Color.White;
            cancelButton.Click += (s, e) => dialog.Close();

            bottomPanel.Controls.Add(okButton);
            bottomPanel.Controls.Add(cancelButton);

            dialog.Controls.Add(flowPanel);
            dialog.Controls.Add(bottomPanel);
            
            // 支持粘贴图片
            dialog.KeyPreview = true;
            dialog.KeyDown += (s, e) =>
            {
                if (e.Control && e.KeyCode == Keys.V)
                {
                    if (Clipboard.ContainsImage())
                    {
                        var image = Clipboard.GetImage();
                        var fileName = $"{Guid.NewGuid():N}.png";
                        var destPath = Path.Combine(_dataService.GetImagePath(), fileName);
                        image.Save(destPath, System.Drawing.Imaging.ImageFormat.Png);
                        images.Add(fileName);
                        RefreshImages();
                        e.Handled = true;
                    }
                    else if (Clipboard.ContainsFileDropList())
                    {
                        var files = Clipboard.GetFileDropList();
                        var imageExtensions = new[] { ".jpg", ".jpeg", ".png", ".gif", ".bmp" };
                        
                        foreach (string file in files)
                        {
                            var ext = Path.GetExtension(file).ToLower();
                            if (imageExtensions.Contains(ext))
                            {
                                var fileName = $"{Guid.NewGuid():N}{ext}";
                                var destPath = Path.Combine(_dataService.GetImagePath(), fileName);
                                File.Copy(file, destPath, true);
                                images.Add(fileName);
                            }
                        }
                        RefreshImages();
                        e.Handled = true;
                    }
                }
            };

            ShowModalDialog(dialog);
        }

        private void ShowTodoListDialog(RowData rowData, string colId, DataGridView grid, int rowIndex)
        {
            var dialog = new Form();
            dialog.Text = "待办事项管理";
            dialog.Size = new Size(500, 600);
            dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
            dialog.MaximizeBox = false;
            dialog.MinimizeBox = false;
            dialog.BackColor = Color.FromArgb(37, 37, 38);

            if (!rowData.Data.ContainsKey(colId))
            {
                rowData.Data[colId] = new List<TodoItem>();
            }

            var todos = GetTodoListFromData(rowData.Data[colId]);

            var todoControl = new TodoListControl(Color.FromArgb(37, 37, 38), () =>
            {
                rowData.Data[colId] = todos;
                _dataService.NotifyDataChanged();
                
                // 更新行高以适应新的TodoList项目数
                int maxHeight = 30; // 默认最小高度
                foreach (var col in GetCurrentTabData().Columns)
                {
                    if (col.Type == ColumnType.TodoList && rowData.Data.ContainsKey(col.Id))
                    {
                        var todoList = GetTodoListFromData(rowData.Data[col.Id]);
                        if (todoList.Count > 0)
                        {
                            int todoHeight = (todoList.Count * 18) + 10;
                            maxHeight = Math.Max(maxHeight, todoHeight);
                        }
                    }
                }
                grid.Rows[rowIndex].Height = maxHeight;
                
                grid.InvalidateCell(grid.Columns[colId].Index, rowIndex);
            });
            todoControl.Dock = DockStyle.Fill;
            todoControl.Todos = todos;

            var bottomPanel = new Panel();
            bottomPanel.Height = 50;
            bottomPanel.Dock = DockStyle.Bottom;
            bottomPanel.BackColor = Color.FromArgb(45, 45, 48);

            var okButton = new Button();
            okButton.Text = "确定";
            okButton.Size = new Size(80, 30);
            okButton.Location = new Point(410, 10);
            okButton.FlatStyle = FlatStyle.Flat;
            okButton.BackColor = Color.FromArgb(0, 122, 204);
            okButton.ForeColor = Color.White;
            okButton.Click += (s, e) =>
            {
                rowData.Data[colId] = todos;
                _dataService.NotifyDataChanged();
                
                // 更新行高以适应新的TodoList项目数
                int maxHeight = 30; // 默认最小高度
                foreach (var col in GetCurrentTabData().Columns)
                {
                    if (col.Type == ColumnType.TodoList && rowData.Data.ContainsKey(col.Id))
                    {
                        var todoList = GetTodoListFromData(rowData.Data[col.Id]);
                        if (todoList.Count > 0)
                        {
                            int todoHeight = (todoList.Count * 18) + 10;
                            maxHeight = Math.Max(maxHeight, todoHeight);
                        }
                    }
                }
                grid.Rows[rowIndex].Height = maxHeight;
                
                grid.InvalidateCell(grid.Columns[colId].Index, rowIndex);
                dialog.Close();
            };

            bottomPanel.Controls.Add(okButton);

            dialog.Controls.Add(todoControl);
            dialog.Controls.Add(bottomPanel);

            ShowModalDialog(dialog);
        }

        private void ShowTextAreaDialog(RowData rowData, string colId, string colName, DataGridView grid, int rowIndex, int colIndex)
        {
            var dialog = new Form();
            dialog.Text = $"编辑{colName} - 自动保存";
            dialog.Size = new Size(600, 400);
            dialog.FormBorderStyle = FormBorderStyle.Sizable;
            dialog.MinimumSize = new Size(400, 300);
            dialog.BackColor = Color.FromArgb(37, 37, 38);

            var textBox = new TextBox();
            textBox.Multiline = true;
            textBox.ScrollBars = ScrollBars.Vertical;
            textBox.Dock = DockStyle.Fill;
            textBox.BackColor = Color.FromArgb(45, 45, 48);
            textBox.ForeColor = Color.White;
            textBox.BorderStyle = BorderStyle.None;
            textBox.Font = new Font("Consolas", 10);

            // 获取当前值
            var currentValue = rowData.Data.ContainsKey(colId) ? rowData.Data[colId]?.ToString() : "";
            textBox.Text = currentValue;

            // 状态栏
            var statusBar = new Panel();
            statusBar.Height = 25;
            statusBar.Dock = DockStyle.Bottom;
            statusBar.BackColor = Color.FromArgb(30, 30, 30);

            var statusLabel = new Label();
            statusLabel.Text = "自动保存已启用";
            statusLabel.ForeColor = Color.LightGray;
            statusLabel.Location = new Point(10, 5);
            statusLabel.AutoSize = true;
            statusBar.Controls.Add(statusLabel);

            // 保存定时器
            var saveTimer = new System.Windows.Forms.Timer();
            saveTimer.Interval = 500; // 500毫秒防抖
            var textChanged = false;

            // 文本改变时触发自动保存
            textBox.TextChanged += (s, e) =>
            {
                textChanged = true;
                saveTimer.Stop();
                saveTimer.Start();
                statusLabel.Text = "正在输入...";
                statusLabel.ForeColor = Color.Yellow;
            };

            saveTimer.Tick += (s, e) =>
            {
                saveTimer.Stop();
                if (textChanged)
                {
                    textChanged = false;
                    // 保存数据
                    rowData.Data[colId] = textBox.Text;
                    grid.Rows[rowIndex].Cells[colIndex].Value = textBox.Text;
                    _dataService.NotifyDataChanged();
                    grid.InvalidateCell(colIndex, rowIndex);
                    
                    // 更新状态
                    statusLabel.Text = $"已自动保存 - {DateTime.Now:HH:mm:ss}";
                    statusLabel.ForeColor = Color.LightGreen;
                }
            };

            // 关闭对话框时停止定时器
            dialog.FormClosing += (s, e) =>
            {
                saveTimer.Stop();
                // 最后一次保存
                if (textChanged)
                {
                    rowData.Data[colId] = textBox.Text;
                    grid.Rows[rowIndex].Cells[colIndex].Value = textBox.Text;
                    _dataService.NotifyDataChanged();
                    grid.InvalidateCell(colIndex, rowIndex);
                }
                saveTimer.Dispose();
            };

            dialog.Controls.Add(textBox);
            dialog.Controls.Add(statusBar);

            textBox.Focus();
            textBox.SelectAll();

            // 支持 Ctrl+S 手动保存
            textBox.KeyDown += (s, e) =>
            {
                if (e.Control && e.KeyCode == Keys.S)
                {
                    e.Handled = true;
                    saveTimer.Stop();
                    if (textChanged)
                    {
                        textChanged = false;
                        rowData.Data[colId] = textBox.Text;
                        grid.Rows[rowIndex].Cells[colIndex].Value = textBox.Text;
                        _dataService.NotifyDataChanged();
                        grid.InvalidateCell(colIndex, rowIndex);
                        
                        statusLabel.Text = $"已手动保存 - {DateTime.Now:HH:mm:ss}";
                        statusLabel.ForeColor = Color.LightGreen;
                    }
                }
            };

            ShowModalDialog(dialog);
        }

        private void Grid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0 || e.Button != MouseButtons.Left) return;

            var grid = (DataGridView)sender;
            var col = grid.Columns[e.ColumnIndex].Tag as ColumnDefinition;
            if (col == null || col.Type != ColumnType.TodoList) return;

            var rowData = grid.Rows[e.RowIndex].Tag as RowData;
            if (rowData == null || !rowData.Data.ContainsKey(col.Id)) return;

            var todos = GetTodoListFromData(rowData.Data[col.Id]);
            if (todos.Count == 0) return;

            // 计算点击位置
            var cellBounds = grid.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, false);
            var relativeY = e.Y;
            var y = 5;

            foreach (var todo in todos)
            {
                if (y + 20 > cellBounds.Height) break;

                // 检查是否点击了圆点
                var dotRect = new Rectangle(5, y + 2, 12, 12);
                if (dotRect.Contains(e.X, relativeY))
                {
                    _todoClickHandled = true; // 标记已处理，阻止CellClick事件
                    
                    // 使用BeginInvoke延迟显示对话框，避免阻塞事件链
                    this.BeginInvoke(new Action(() =>
                    {
                        var message = todo.IsCompleted ? "确认取消完成状态？" : "确认完成此待办事项？";
                        var result = MessageBox.Show(message, "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            todo.IsCompleted = !todo.IsCompleted;
                            rowData.Data[col.Id] = todos;
                            _dataService.NotifyDataChanged();
                            grid.InvalidateCell(e.ColumnIndex, e.RowIndex);
                        }
                        // 重置标志
                        _todoClickHandled = false;
                    }));
                    return;
                }

                y += 18;
            }
        }

        private void Grid_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                var grid = (DataGridView)sender;
                var hit = grid.HitTest(e.X, e.Y);
                
                if (hit.RowIndex >= 0)
                {
                    grid.ClearSelection();
                    grid.Rows[hit.RowIndex].Selected = true;
                    _selectedRow = grid.Rows[hit.RowIndex];
                    _rowContextMenu.Show(grid, e.Location);
                }
            }
        }

        private void Grid_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex < 0) return;
            
            var grid = (DataGridView)sender;
            var col = grid.Columns[e.ColumnIndex].Tag as ColumnDefinition;
            
            if (e.Button == MouseButtons.Left && col != null)
            {
                // 获取鼠标点击的位置（相对于网格控件）
                var mousePos = grid.PointToClient(Cursor.Position);
                
                // 获取列头的矩形区域
                var colRect = grid.GetColumnDisplayRectangle(e.ColumnIndex, false);
                
                // 检查是否点击在图标区域（列右侧25像素内）
                if (mousePos.X >= colRect.Left && mousePos.X <= colRect.Right && 
                    mousePos.X >= colRect.Right - 25) // 图标区域
                {
                    if (col.Type == ColumnType.Single)
                    {
                        ShowFilterDialog(grid, col, e.ColumnIndex);
                    }
                    else if (col.Type == ColumnType.Text)
                    {
                        ShowSortDialog(grid, col, e.ColumnIndex);
                    }
                }
            }
            else if (e.Button == MouseButtons.Right)
            {
                _selectedColumnIndex = e.ColumnIndex;
                
                // 根据列类型设置"列求和"菜单项的启用状态
                foreach (ToolStripItem item in _columnContextMenu.Items)
                {
                    if (item.Text == Lang.Get("SumColumn"))
                    {
                        item.Enabled = col.Type == ColumnType.Text;
                        break;
                    }
                }
                
                _columnContextMenu.Show(grid, grid.PointToClient(Cursor.Position));
            }
        }

        private void ShowFilterDialog(DataGridView grid, ColumnDefinition col, int columnIndex)
        {
            var dialog = new Form
            {
                Text = $"{Lang.Get("Filter")} - {col.Name}",
                Size = new Size(300, 400),
                StartPosition = FormStartPosition.Manual,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            // 应用主题
            ApplyThemeToDialog(dialog);

            // 将对话框显示在屏幕正中间
            dialog.StartPosition = FormStartPosition.CenterScreen;

            var colors = ThemeService.Instance.GetColors();
            
            var checkedListBox = new CheckedListBox
            {
                Location = new Point(10, 10),
                Size = new Size(264, 300),
                BackColor = colors.InputBackground,
                ForeColor = colors.Text,
                BorderStyle = BorderStyle.None,
                CheckOnClick = true
            };

            // 获取该列的所有唯一值
            var uniqueValues = new HashSet<string>();
            var tabData = (TabData)_tabControl.SelectedTab.Tag;
            foreach (var row in tabData.Rows)
            {
                if (row.Data.ContainsKey(col.Id))
                {
                    var value = row.Data[col.Id]?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(value))
                    {
                        uniqueValues.Add(value);
                    }
                }
            }

            // 获取当前筛选
            var filterKey = $"{tabData.Id}_{col.Id}";
            var currentFilter = _columnFilters.ContainsKey(filterKey) ? _columnFilters[filterKey] : new List<string>();

            // 添加选项
            foreach (var option in col.Options)
            {
                if (uniqueValues.Contains(option.Label))
                {
                    var isChecked = currentFilter.Count == 0 || currentFilter.Contains(option.Label);
                    checkedListBox.Items.Add(option.Label, isChecked);
                }
            }

            var buttonPanel = new Panel
            {
                Height = 40,
                Dock = DockStyle.Bottom,
                BackColor = colors.PanelBackground
            };

            var applyButton = new Button
            {
                Text = Lang.Get("OK"),
                Location = new Point(110, 5),
                Size = new Size(80, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = colors.PrimaryButton,
                ForeColor = Color.White
            };
            applyButton.FlatAppearance.BorderColor = colors.PrimaryButton;

            var clearButton = new Button
            {
                Text = Lang.Get("Clear"),
                Location = new Point(20, 5),
                Size = new Size(80, 30),
                FlatStyle = FlatStyle.Flat,
                BackColor = colors.ButtonBackground,
                ForeColor = colors.Text
            };
            clearButton.FlatAppearance.BorderColor = colors.Border;

            applyButton.Click += (s, e) =>
            {
                var selectedItems = new List<string>();
                foreach (string item in checkedListBox.CheckedItems)
                {
                    selectedItems.Add(item);
                }

                if (selectedItems.Count == checkedListBox.Items.Count || selectedItems.Count == 0)
                {
                    // 全选或全不选，移除筛选
                    _columnFilters.Remove(filterKey);
                }
                else
                {
                    _columnFilters[filterKey] = selectedItems;
                }

                ApplyFilters(grid);
                dialog.Close();
            };

            clearButton.Click += (s, e) =>
            {
                _columnFilters.Remove(filterKey);
                ApplyFilters(grid);
                dialog.Close();
            };

            buttonPanel.Controls.Add(clearButton);
            buttonPanel.Controls.Add(applyButton);
            dialog.Controls.Add(checkedListBox);
            dialog.Controls.Add(buttonPanel);

            dialog.ShowDialog();
        }

        private void ShowSortDialog(DataGridView grid, ColumnDefinition col, int columnIndex)
        {
            var menu = new ContextMenuStrip();
            
            var ascItem = new ToolStripMenuItem($"↑ {Lang.Get("SortAscending")}");
            ascItem.Click += (s, e) => SortColumn(grid, columnIndex, true);
            
            var descItem = new ToolStripMenuItem($"↓ {Lang.Get("SortDescending")}");
            descItem.Click += (s, e) => SortColumn(grid, columnIndex, false);
            
            var clearItem = new ToolStripMenuItem(Lang.Get("ClearSort"));
            clearItem.Click += (s, e) => 
            {
                // 重新加载数据以清除排序
                var tabData = GetCurrentTabData();
                UpdateGridRows(grid, tabData);
            };
            
            menu.Items.Add(ascItem);
            menu.Items.Add(descItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(clearItem);
            
            // 将菜单显示在屏幕正中间
            var screen = Screen.FromControl(grid);
            var centerX = screen.WorkingArea.Left + (screen.WorkingArea.Width / 2);
            var centerY = screen.WorkingArea.Top + (screen.WorkingArea.Height / 2);
            
            menu.Show(centerX, centerY);
        }

        private void SortColumn(DataGridView grid, int columnIndex, bool ascending)
        {
            grid.Sort(grid.Columns[columnIndex], ascending ? 
                System.ComponentModel.ListSortDirection.Ascending : 
                System.ComponentModel.ListSortDirection.Descending);
        }

        private void ApplyFilters(DataGridView grid)
        {
            var tabData = (TabData)_tabControl.SelectedTab.Tag;
            
            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row.IsNewRow) continue;
                
                bool visible = true;
                var rowData = (RowData)row.Tag;
                
                // 检查所有筛选条件
                foreach (var filter in _columnFilters)
                {
                    var parts = filter.Key.Split('_');
                    if (parts[0] == tabData.Id)
                    {
                        var colId = parts[1];
                        if (rowData.Data.ContainsKey(colId))
                        {
                            var value = rowData.Data[colId]?.ToString() ?? "";
                            if (!filter.Value.Contains(value))
                            {
                                visible = false;
                                break;
                            }
                        }
                    }
                }
                
                row.Visible = visible;
            }
        }

        private void OnAddTabClicked(object sender, EventArgs e)
        {
            using (var dialog = new SimpleInputDialog(Lang.Get("NewTab"), Lang.Get("EnterTabName"), $"{Lang.Get("NewTab")} {_tabControl.TabCount + 1}"))
            {
                if (dialog.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.InputValue))
                {
                    var tabData = new TabData { Name = dialog.InputValue };
                    tabData.Columns.Add(new ColumnDefinition { Name = Lang.Get("Task"), Type = ColumnType.Text });
                    
                    _dataService.GetAppData().Tabs.Add(tabData);
                    CreateTabPage(tabData);
                    _dataService.NotifyDataChanged();
                    
                    _tabControl.SelectedIndex = _tabControl.TabCount - 1;
                    if (this.IsHandleCreated)
                    {
                        this.BeginInvoke(new Action(() => UpdateTabButtons()));
                    }
                }
            }
        }

        private void OnTabRenameRequested(int tabIndex, string currentName)
        {
            using (var dialog = new SimpleInputDialog("重命名标签页", "请输入新的标签页名称:", currentName))
            {
                if (dialog.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.InputValue))
                {
                    var tabData = (TabData)_tabControl.TabPages[tabIndex].Tag;
                    tabData.Name = dialog.InputValue;
                    _tabControl.TabPages[tabIndex].Text = dialog.InputValue;
                    _dataService.NotifyDataChanged();
                    if (this.IsHandleCreated)
                    {
                        this.BeginInvoke(new Action(() => UpdateTabButtons()));
                    }
                }
            }
        }

        private void OnRenameTabClick(object sender, EventArgs e)
        {
            if (_selectedTab == null) return;

            using (var dialog = new SimpleInputDialog("重命名标签页", "请输入新的标签页名称:", _selectedTab.Text))
            {
                if (dialog.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.InputValue))
                {
                    var tabData = (TabData)_selectedTab.Tag;
                    tabData.Name = dialog.InputValue;
                    _selectedTab.Text = dialog.InputValue;
                    _dataService.NotifyDataChanged();
                    if (this.IsHandleCreated)
                    {
                        this.BeginInvoke(new Action(() => UpdateTabButtons()));
                    }
                }
            }
        }

        private void OnCloseTabClick(object sender, EventArgs e)
        {
            if (_selectedTab == null || _tabControl.TabCount <= 1) return;

            var result = MessageBox.Show($"确定要关闭标签页 \"{_selectedTab.Text}\" 吗？\n数据会保留，下次可以重新打开。", 
                "确认关闭", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            
            if (result == DialogResult.Yes)
            {
                var tabData = (TabData)_selectedTab.Tag;
                _dataService.GetAppData().Tabs.Remove(tabData);
                _tabControl.TabPages.Remove(_selectedTab);
                _dataService.NotifyDataChanged();
                if (this.IsHandleCreated)
                {
                    this.BeginInvoke(new Action(() => UpdateTabButtons()));
                }
            }
        }

        private void AddNewRow(TabData tabData, DataGridView grid)
        {
            var newRow = new RowData();
            tabData.Rows.Add(newRow);
            UpdateGridRows((ModernDataGridView)grid, tabData);
            _dataService.NotifyDataChanged(tabData);
            UpdateTimestampLabel(tabData);
        }

        private void OnDeleteRowClick(object sender, EventArgs e)
        {
            if (_selectedRow == null) return;

            var result = MessageBox.Show(Lang.Get("ConfirmDeleteRow"), Lang.Get("Delete"), 
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            
            if (result == DialogResult.Yes)
            {
                var tabData = GetCurrentTabData();
                var rowData = (RowData)_selectedRow.Tag;
                tabData.Rows.Remove(rowData);
                
                var grid = GetCurrentGrid();
                UpdateGridRows(grid, tabData);
                _dataService.NotifyDataChanged(tabData);
                UpdateTimestampLabel(tabData);
            }
        }

        private void OnCopyRowClick(object sender, EventArgs e)
        {
            if (_selectedRow == null) return;

            var tabData = GetCurrentTabData();
            var sourceRow = (RowData)_selectedRow.Tag;
            var newRow = new RowData();
            
            foreach (var kvp in sourceRow.Data)
            {
                newRow.Data[kvp.Key] = kvp.Value;
            }
            
            var index = tabData.Rows.IndexOf(sourceRow);
            tabData.Rows.Insert(index + 1, newRow);
            
            var grid = GetCurrentGrid();
            UpdateGridRows(grid, tabData);
            _dataService.NotifyDataChanged();
        }

        private void OnAddRowAboveClick(object sender, EventArgs e)
        {
            if (_selectedRow == null) return;

            var tabData = GetCurrentTabData();
            var currentRow = (RowData)_selectedRow.Tag;
            var newRow = new RowData();
            
            var index = tabData.Rows.IndexOf(currentRow);
            tabData.Rows.Insert(index, newRow);
            
            var grid = GetCurrentGrid();
            UpdateGridRows(grid, tabData);
            _dataService.NotifyDataChanged();
        }

        private void OnAddRowBelowClick(object sender, EventArgs e)
        {
            if (_selectedRow == null) return;

            var tabData = GetCurrentTabData();
            var currentRow = (RowData)_selectedRow.Tag;
            var newRow = new RowData();
            
            var index = tabData.Rows.IndexOf(currentRow);
            tabData.Rows.Insert(index + 1, newRow);
            
            var grid = GetCurrentGrid();
            UpdateGridRows(grid, tabData);
            _dataService.NotifyDataChanged();
        }

        private void OnAddColumnClick(object sender, EventArgs e)
        {
            using (var dialog = new SimpleColumnEditDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var tabData = GetCurrentTabData();
                    var grid = GetCurrentGrid();
                    
                    tabData.Columns.Add(dialog.GetColumnDefinition());
                    UpdateGridColumns(grid, tabData);
                    UpdateGridRows(grid, tabData);
                    _dataService.NotifyDataChanged(tabData);
                    UpdateTimestampLabel(tabData);
                }
            }
        }

        private void OnEditColumnClick(object sender, EventArgs e)
        {
            var grid = GetCurrentGrid();
            var col = (ColumnDefinition)grid.Columns[_selectedColumnIndex].Tag;
            
            using (var dialog = new SimpleColumnEditDialog(col))
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var tabData = GetCurrentTabData();
                    var editedColumn = dialog.GetColumnDefinition();
                    
                    // 更新列的属性
                    col.Name = editedColumn.Name;
                    col.Type = editedColumn.Type;
                    
                    UpdateGridColumns(grid, tabData);
                    UpdateGridRows(grid, tabData);
                    _dataService.NotifyDataChanged(tabData);
                    UpdateTimestampLabel(tabData);
                }
            }
        }

        private void OnSumColumnClick(object sender, EventArgs e)
        {
            var grid = GetCurrentGrid();
            var col = (ColumnDefinition)grid.Columns[_selectedColumnIndex].Tag;
            var tabData = GetCurrentTabData();
            
            double sum = 0;
            int validCount = 0;
            int invalidCount = 0;
            var invalidValues = new List<string>();
            
            // 遍历所有行，尝试将单元格值转换为数字
            foreach (var row in tabData.Rows)
            {
                if (row.Data.ContainsKey(col.Id))
                {
                    var cellValue = row.Data[col.Id]?.ToString();
                    if (!string.IsNullOrWhiteSpace(cellValue))
                    {
                        // 尝试解析为数字
                        if (double.TryParse(cellValue.Trim(), out double value))
                        {
                            sum += value;
                            validCount++;
                        }
                        else
                        {
                            invalidCount++;
                            if (invalidValues.Count < 5) // 只记录前5个无效值作为示例
                            {
                                invalidValues.Add(cellValue);
                            }
                        }
                    }
                }
            }
            
            // 构建结果消息
            var message = $"列 \"{col.Name}\" 求和结果：\n\n";
            message += $"总和：{sum:F2}\n";
            message += $"有效数值个数：{validCount}\n";
            
            if (invalidCount > 0)
            {
                message += $"非数值个数：{invalidCount}\n";
                if (invalidValues.Count > 0)
                {
                    message += $"非数值示例：{string.Join(", ", invalidValues.Select(v => $"\"{v}\""))}";
                    if (invalidCount > 5)
                    {
                        message += " ...";
                    }
                }
            }
            
            // 显示结果对话框
            var resultDialog = new Form
            {
                Text = "列求和结果",
                Size = new Size(400, 250),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                BackColor = Color.FromArgb(37, 37, 38)
            };
            
            var resultLabel = new Label
            {
                Text = message,
                Location = new Point(20, 20),
                Size = new Size(360, 150),
                ForeColor = Color.White,
                Font = new Font("Microsoft YaHei UI", 10)
            };
            
            var okButton = new Button
            {
                Text = "确定",
                Location = new Point(160, 180),
                Size = new Size(80, 30),
                DialogResult = DialogResult.OK,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White
            };
            
            resultDialog.Controls.Add(resultLabel);
            resultDialog.Controls.Add(okButton);
            
            ShowModalDialog(resultDialog);
        }

        private void OnDeleteColumnClick(object sender, EventArgs e)
        {
            var grid = GetCurrentGrid();
            var col = (ColumnDefinition)grid.Columns[_selectedColumnIndex].Tag;
            
            var result = MessageBox.Show(Lang.Get("ConfirmDeleteColumn", col.Name), 
                Lang.Get("Delete"), MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            
            if (result == DialogResult.Yes)
            {
                var tabData = GetCurrentTabData();
                tabData.Columns.Remove(col);
                
                foreach (var row in tabData.Rows)
                {
                    row.Data.Remove(col.Id);
                }
                
                UpdateGridColumns(grid, tabData);
                UpdateGridRows(grid, tabData);
                _dataService.NotifyDataChanged(tabData);
                UpdateTimestampLabel(tabData);
            }
        }

        private TabData GetCurrentTabData()
        {
            return (TabData)_tabControl.SelectedTab?.Tag;
        }

        private DataGridView GetCurrentGrid()
        {
            var panel = _tabControl.SelectedTab?.Controls[0] as Panel;
            return panel?.Controls.OfType<DataGridView>().FirstOrDefault();
        }
        
        private void UpdateTimestampLabel(TabData tabData)
        {
            if (tabData == null) return;
            
            // 查找对应的标签页
            foreach (TabPage tabPage in _tabControl.TabPages)
            {
                if (tabPage.Tag == tabData)
                {
                    var panel = tabPage.Controls[0] as Panel;
                    if (panel != null)
                    {
                        var bottomPanel = panel.Controls.OfType<Panel>().FirstOrDefault(p => p.Dock == DockStyle.Bottom);
                        if (bottomPanel != null)
                        {
                            var timestampLabel = bottomPanel.Controls.OfType<Label>().FirstOrDefault(l => l.Tag == tabData);
                            if (timestampLabel != null)
                            {
                                timestampLabel.Text = $"最后修改: {tabData.LastModified:yyyy-MM-dd HH:mm:ss}";
                            }
                        }
                    }
                    break;
                }
            }
        }
        
        private void ShowAICreateTableDialog()
        {
            var dialog = new Form
            {
                Text = Lang.Get("AICreateTable"),
                Size = new Size(600, 500),
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };
            
            // 应用主题
            ApplyThemeToDialog(dialog);
            var colors = ThemeService.Instance.GetColors();
            
            // AI提供商选择
            var providerLabel = new Label
            {
                Text = Lang.Get("AIProvider") + ":",
                Location = new Point(20, 20),
                Size = new Size(100, 25),
                ForeColor = colors.Text
            };
            
            var providerCombo = new ComboBox
            {
                Location = new Point(130, 20),
                Size = new Size(150, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = colors.InputBackground,
                ForeColor = colors.Text
            };
            providerCombo.Items.Add("OpenAI");
            providerCombo.Items.Add("DeepSeek");
            providerCombo.Items.Add("Claude");
            
            // API Key输入
            var apiKeyLabel = new Label
            {
                Text = Lang.Get("APIKey") + ":",
                Location = new Point(20, 60),
                Size = new Size(100, 25),
                ForeColor = colors.Text
            };
            
            var apiKeyTextBox = new TextBox
            {
                Location = new Point(130, 60),
                Size = new Size(430, 25),
                PasswordChar = '*',
                BackColor = colors.InputBackground,
                ForeColor = colors.Text
            };
            
            // 需求输入
            var requirementLabel = new Label
            {
                Text = Lang.Get("TableRequirement") + ":",
                Location = new Point(20, 100),
                Size = new Size(560, 25),
                ForeColor = colors.Text
            };
            
            var requirementTextBox = new TextBox
            {
                Location = new Point(20, 130),
                Size = new Size(540, 250),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = colors.InputBackground,
                ForeColor = colors.Text
            };
            requirementTextBox.PlaceholderText = Lang.Get("EnterRequirement");
            
            // 生成按钮
            var generateButton = new Button
            {
                Text = Lang.Get("GenerateTable"),
                Location = new Point(380, 400),
                Size = new Size(100, 35),
                FlatStyle = FlatStyle.Flat,
                BackColor = colors.PrimaryButton,
                ForeColor = Color.White
            };
            generateButton.FlatAppearance.BorderColor = colors.PrimaryButton;
            
            // 取消按钮
            var cancelButton = new Button
            {
                Text = Lang.Get("Cancel"),
                Location = new Point(490, 400),
                Size = new Size(70, 35),
                FlatStyle = FlatStyle.Flat,
                BackColor = colors.ButtonBackground,
                ForeColor = colors.Text
            };
            cancelButton.FlatAppearance.BorderColor = colors.Border;
            
            // 加载保存的配置
            LoadAIConfig(providerCombo, apiKeyTextBox);
            
            // 查看日志按钮
            var viewLogButton = new Button
            {
                Text = "📋",
                Location = new Point(520, 20),
                Size = new Size(40, 25),
                FlatStyle = FlatStyle.Flat,
                BackColor = colors.ButtonBackground,
                ForeColor = colors.Text,
                // ToolTip will be set below
            };
            viewLogButton.FlatAppearance.BorderColor = colors.Border;
            viewLogButton.Click += (s, e) =>
            {
                var logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                    "BfTaskBoard", "ai_logs");
                if (Directory.Exists(logPath))
                {
                    Process.Start("explorer.exe", logPath);
                }
                else
                {
                    MessageBox.Show("No logs found yet.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
            
            // Set tooltip for log button
            var toolTip = new ToolTip();
            toolTip.SetToolTip(viewLogButton, "View AI Logs");
            
            generateButton.Click += async (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(apiKeyTextBox.Text))
                {
                    MessageBox.Show(Lang.Get("InvalidAPIKey"), Lang.Get("Error"), 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                if (string.IsNullOrWhiteSpace(requirementTextBox.Text))
                {
                    MessageBox.Show(Lang.Get("PleaseEnterRequirement"), Lang.Get("Error"), 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                // 保存配置
                SaveAIConfig(providerCombo.SelectedIndex, apiKeyTextBox.Text);
                
                // 显示进度
                generateButton.Enabled = false;
                generateButton.Text = Lang.Get("GeneratingTable");
                
                try
                {
                    var config = new AIService.AIConfig
                    {
                        Provider = (AIService.AIProvider)providerCombo.SelectedIndex,
                        ApiKey = apiKeyTextBox.Text
                    };
                    
                    var tabData = await AIService.GenerateTableAsync(config, requirementTextBox.Text);
                    
                    // 添加到应用
                    _dataService.GetAppData().Tabs.Add(tabData);
                    CreateTabPage(tabData);
                    _dataService.NotifyDataChanged(tabData);
                    
                    dialog.Close();
                    MessageBox.Show(Lang.Get("TableGeneratedSuccess"), Lang.Get("Success"), 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(Lang.Get("TableGenerateFailed", ex.Message), 
                        Lang.Get("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    generateButton.Enabled = true;
                    generateButton.Text = Lang.Get("GenerateTable");
                }
            };
            
            cancelButton.Click += (s, e) => dialog.Close();
            
            dialog.Controls.AddRange(new Control[] {
                providerLabel, providerCombo, viewLogButton,
                apiKeyLabel, apiKeyTextBox,
                requirementLabel, requirementTextBox,
                generateButton, cancelButton
            });
            
            dialog.ShowDialog();
        }
        
        private void LoadAIConfig(ComboBox providerCombo, TextBox apiKeyTextBox)
        {
            try
            {
                var configPath = Path.Combine(_dataService.GetDataPath(), "ai_config.json");
                if (File.Exists(configPath))
                {
                    var json = File.ReadAllText(configPath);
                    var config = JsonConvert.DeserializeObject<dynamic>(json);
                    if (config != null && config.provider != null)
                    {
                        providerCombo.SelectedIndex = (int)config.provider;
                        if (config.apiKey != null)
                        {
                            apiKeyTextBox.Text = config.apiKey.ToString();
                        }
                    }
                    else
                    {
                        providerCombo.SelectedIndex = 0;
                    }
                }
                else
                {
                    providerCombo.SelectedIndex = 0;
                }
            }
            catch
            {
                providerCombo.SelectedIndex = 0;
            }
        }
        
        private void SaveAIConfig(int provider, string apiKey)
        {
            try
            {
                var configPath = Path.Combine(_dataService.GetDataPath(), "ai_config.json");
                var config = new { provider = provider, apiKey = apiKey };
                var json = JsonConvert.SerializeObject(config, Formatting.Indented);
                File.WriteAllText(configPath, json);
            }
            catch { }
        }
    }

    // 自定义标签按钮类
    public class TabButton : Panel
    {
        public new int TabIndex { get; set; }
        public string TabId { get; set; }
        public bool IsActive { get; set; }
        private Button closeButton;
        private Label titleLabel;
        private Panel dotPanel;
        private Color dotColor;
        private bool isHovering = false;

        public event EventHandler<MouseEventArgs> TabClick;
        public event EventHandler TabDoubleClick;
        public event EventHandler CloseClick;
        public event EventHandler DotClick;

        public TabButton(string text, int index, string tabId, string dotColorHex = "#757575")
        {
            TabIndex = index;
            TabId = tabId;
            Height = 30;
            Width = 130; // 稍微增加宽度以容纳圆点
            Cursor = Cursors.Hand;
            
            // 解析颜色
            try
            {
                dotColor = ColorTranslator.FromHtml(dotColorHex);
            }
            catch
            {
                dotColor = Color.FromArgb(117, 117, 117); // 默认灰色
            }
            
            // 圆点面板
            dotPanel = new Panel
            {
                Size = new Size(10, 10),
                Location = new Point(8, 10),
                BackColor = Color.Transparent,
                Cursor = Cursors.Hand
            };
            dotPanel.Paint += DotPanel_Paint;
            dotPanel.Click += (s, e) => DotClick?.Invoke(this, e);

            titleLabel = new Label
            {
                Text = text,
                AutoSize = false,
                AutoEllipsis = true,  // 启用省略号
                TextAlign = ContentAlignment.MiddleLeft,
                Location = new Point(25, 5), // 调整Y轴位置使文本与圆点水平对齐
                Size = new Size(80, 20), // 调整高度
                ForeColor = Color.White,
                BackColor = Color.Transparent
            };

            closeButton = new Button
            {
                Text = "×",
                Size = new Size(20, 20),
                Location = new Point(Width - 25, 5),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Transparent,
                ForeColor = Color.Gray,
                Visible = false,
                Cursor = Cursors.Hand
            };
            closeButton.FlatAppearance.BorderSize = 0;
            closeButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(80, 80, 80);

            Controls.Add(dotPanel);
            Controls.Add(titleLabel);
            Controls.Add(closeButton);

            titleLabel.MouseClick += (s, e) => TabClick?.Invoke(this, e);
            titleLabel.DoubleClick += (s, e) => TabDoubleClick?.Invoke(this, e);
            closeButton.Click += (s, e) => 
            {
                CloseClick?.Invoke(this, e);
            };

            MouseEnter += OnMouseEnter;
            MouseLeave += OnMouseLeave;
            titleLabel.MouseEnter += OnMouseEnter;
            titleLabel.MouseLeave += OnMouseLeave;
            dotPanel.MouseEnter += OnMouseEnter;

            UpdateAppearance();
        }
        
        private void DotPanel_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            using (var brush = new SolidBrush(dotColor))
            {
                e.Graphics.FillEllipse(brush, 0, 0, 9, 9);
            }
        }

        private void OnMouseEnter(object sender, EventArgs e)
        {
            isHovering = true;
            closeButton.Visible = true;
            UpdateAppearance();
        }

        private void OnMouseLeave(object sender, EventArgs e)
        {
            if (!ClientRectangle.Contains(PointToClient(MousePosition)))
            {
                isHovering = false;
                closeButton.Visible = false;
                UpdateAppearance();
            }
        }

        public void UpdateAppearance()
        {
            var colors = ThemeService.Instance.GetColors();
            
            if (IsActive)
            {
                BackColor = colors.TabActiveBackground;
            }
            else if (isHovering)
            {
                BackColor = colors.TabHoverBackground;
            }
            else
            {
                BackColor = colors.TabBackground;
            }
            
            // 更新文本颜色
            titleLabel.ForeColor = colors.Text;
            closeButton.ForeColor = colors.Text;
        }
    }

    // 简单的输入对话框，不依赖 MaterialSkin
    public class SimpleInputDialog : Form
    {
        private TextBox _textBox;
        public string InputValue => _textBox.Text;

        public SimpleInputDialog(string title, string prompt, string defaultValue = "")
        {
            var colors = ThemeService.Instance.GetColors();
            
            this.Text = title;
            this.Size = new Size(400, 200);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = colors.FormBackground;

            var promptLabel = new Label();
            promptLabel.Text = prompt;
            promptLabel.Location = new Point(20, 20);
            promptLabel.AutoSize = true;
            promptLabel.ForeColor = colors.Text;

            _textBox = new TextBox();
            _textBox.Text = defaultValue;
            _textBox.Location = new Point(20, 50);
            _textBox.Width = 360;
            _textBox.BackColor = colors.InputBackground;
            _textBox.ForeColor = colors.Text;
            _textBox.BorderStyle = BorderStyle.FixedSingle;

            var okButton = new Button();
            okButton.Text = Lang.Get("OK");
            okButton.Location = new Point(225, 100);
            okButton.Size = new Size(75, 30);
            okButton.DialogResult = DialogResult.OK;
            okButton.FlatStyle = FlatStyle.Flat;
            okButton.BackColor = colors.PrimaryButton;
            okButton.ForeColor = Color.White;
            okButton.FlatAppearance.BorderColor = colors.PrimaryButton;

            var cancelButton = new Button();
            cancelButton.Text = Lang.Get("Cancel");
            cancelButton.Location = new Point(305, 100);
            cancelButton.Size = new Size(75, 30);
            cancelButton.DialogResult = DialogResult.Cancel;
            cancelButton.FlatStyle = FlatStyle.Flat;
            cancelButton.BackColor = colors.ButtonBackground;
            cancelButton.ForeColor = colors.Text;
            cancelButton.FlatAppearance.BorderColor = colors.Border;

            this.Controls.Add(promptLabel);
            this.Controls.Add(_textBox);
            this.Controls.Add(okButton);
            this.Controls.Add(cancelButton);

            _textBox.Focus();
            _textBox.SelectAll();
        }
    }

    // 简单的列编辑对话框
    public class SimpleColumnEditDialog : Form
    {
        private TextBox _nameTextBox;
        private ComboBox _typeComboBox;
        private ColumnDefinition _column;

        public SimpleColumnEditDialog(ColumnDefinition column = null)
        {
            _column = column ?? new ColumnDefinition();
            InitializeUI();
        }

        private void InitializeUI()
        {
            var colors = ThemeService.Instance.GetColors();
            
            this.Text = _column.Id == _column.Name ? Lang.Get("AddColumnDialog") : Lang.Get("EditColumnDialog");
            this.Size = new Size(400, 200);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = colors.FormBackground;

            var nameLabel = new Label
            {
                Text = Lang.Get("ColumnName"),
                Location = new Point(20, 20),
                AutoSize = true,
                ForeColor = colors.Text
            };

            _nameTextBox = new TextBox
            {
                Text = _column.Name,
                Location = new Point(20, 45),
                Width = 360,
                BackColor = colors.InputBackground,
                ForeColor = colors.Text,
                BorderStyle = BorderStyle.FixedSingle
            };

            var typeLabel = new Label
            {
                Text = Lang.Get("ColumnType"),
                Location = new Point(20, 75),
                AutoSize = true,
                ForeColor = colors.Text
            };

            _typeComboBox = new ComboBox
            {
                Location = new Point(20, 100),
                Width = 360,
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = colors.InputBackground,
                ForeColor = colors.Text
            };

            // 添加列类型选项
            _typeComboBox.Items.Add(new ComboBoxItem(Lang.Get("Text"), ColumnType.Text));
            _typeComboBox.Items.Add(new ComboBoxItem(Lang.Get("SingleSelect"), ColumnType.Single));
            _typeComboBox.Items.Add(new ComboBoxItem(Lang.Get("Image"), ColumnType.Image));
            _typeComboBox.Items.Add(new ComboBoxItem(Lang.Get("TodoList"), ColumnType.TodoList));
            _typeComboBox.Items.Add(new ComboBoxItem(Lang.Get("TextArea"), ColumnType.TextArea));

            // 设置当前选择
            for (int i = 0; i < _typeComboBox.Items.Count; i++)
            {
                if (((ComboBoxItem)_typeComboBox.Items[i]).Value == _column.Type)
                {
                    _typeComboBox.SelectedIndex = i;
                    break;
                }
            }

            var okButton = new Button
            {
                Text = Lang.Get("OK"),
                Location = new Point(225, 140),
                Size = new Size(75, 30),
                DialogResult = DialogResult.OK,
                FlatStyle = FlatStyle.Flat,
                BackColor = colors.PrimaryButton,
                ForeColor = Color.White
            };
            okButton.FlatAppearance.BorderColor = colors.PrimaryButton;

            var cancelButton = new Button
            {
                Text = Lang.Get("Cancel"),
                Location = new Point(305, 140),
                Size = new Size(75, 30),
                DialogResult = DialogResult.Cancel,
                FlatStyle = FlatStyle.Flat,
                BackColor = colors.ButtonBackground,
                ForeColor = colors.Text
            };
            cancelButton.FlatAppearance.BorderColor = colors.Border;

            this.Controls.AddRange(new Control[] { nameLabel, _nameTextBox, typeLabel, _typeComboBox, okButton, cancelButton });

            _nameTextBox.Focus();
            _nameTextBox.SelectAll();
        }

        public ColumnDefinition GetColumnDefinition()
        {
            _column.Name = _nameTextBox.Text;
            if (_typeComboBox.SelectedItem is ComboBoxItem item)
            {
                _column.Type = item.Value;
            }
            return _column;
        }

        private class ComboBoxItem
        {
            public string Text { get; set; }
            public ColumnType Value { get; set; }

            public ComboBoxItem(string text, ColumnType value)
            {
                Text = text;
                Value = value;
            }

            public override string ToString()
            {
                return Text;
            }
        }
    }
}