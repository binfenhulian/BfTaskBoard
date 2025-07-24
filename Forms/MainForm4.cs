using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;
using System.Collections.Generic;
using System.IO;
using MaterialSkin;
using MaterialSkin.Controls;
using TaskBoard.Controls;
using TaskBoard.Forms;
using TaskBoard.Models;
using TaskBoard.Services;
using Microsoft.Win32;
using Newtonsoft.Json.Linq;

namespace TaskBoard
{
    public partial class MainForm4 : MaterialForm
    {
        private DataService _dataService;
        private TabControl _tabControl;
        private Panel _tabPanel;
        private Button _addTabButton;
        private MaterialContextMenuStrip _rowContextMenu;
        private MaterialContextMenuStrip _columnContextMenu;
        private MaterialContextMenuStrip _tabContextMenu;
        private DataGridViewRow _selectedRow;
        private int _selectedColumnIndex;
        private TabPage _selectedTab;

        public MainForm4()
        {
            InitializeComponent();
            InitializeMaterialDesign();
            InitializeServices();
            SetupAutoStart();
            LoadTabs();
        }

        private void InitializeMaterialDesign()
        {
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.DARK;
            materialSkinManager.ColorScheme = new ColorScheme(
                Primary.BlueGrey800, Primary.BlueGrey900,
                Primary.BlueGrey500, Accent.LightBlue200,
                TextShade.WHITE);
        }

        private void InitializeComponent()
        {
            this.Text = "TaskBoard 2.0 - 现代任务管理";
            this.Size = new Size(1400, 900);
            this.StartPosition = FormStartPosition.CenterScreen;

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

            var openDataFolderButton = new MaterialButton();
            openDataFolderButton.Text = "打开数据文件夹";
            openDataFolderButton.Location = new Point(10, 2);
            openDataFolderButton.AutoSize = false;
            openDataFolderButton.Size = new Size(120, 26);
            openDataFolderButton.Click += (s, e) => OpenDataFolder();

            bottomPanel.Controls.Add(openDataFolderButton);

            this.Controls.Add(_tabControl);
            this.Controls.Add(_tabPanel);
            this.Controls.Add(bottomPanel);
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

            int x = 40;
            for (int i = 0; i < _tabControl.TabCount; i++)
            {
                var tabButton = new Button();
                tabButton.Text = _tabControl.TabPages[i].Text;
                tabButton.Tag = i;
                tabButton.Location = new Point(x, 5);
                tabButton.Height = 30;
                tabButton.AutoSize = true;
                tabButton.MinimumSize = new Size(100, 30);
                tabButton.FlatStyle = FlatStyle.Flat;
                tabButton.FlatAppearance.BorderSize = 0;
                tabButton.BackColor = _tabControl.SelectedIndex == i ? Color.FromArgb(45, 45, 45) : Color.FromArgb(35, 35, 35);
                tabButton.ForeColor = Color.White;
                tabButton.Cursor = Cursors.Hand;
                tabButton.MouseClick += TabButton_MouseClick;
                tabButton.DoubleClick += TabButton_DoubleClick;

                _tabPanel.Controls.Add(tabButton);
                x += tabButton.Width + 5;
            }

            _addTabButton.Location = new Point(x, 5);
        }

        private void TabButton_MouseClick(object sender, MouseEventArgs e)
        {
            var button = sender as Button;
            if (button == null) return;

            if (e.Button == MouseButtons.Left)
            {
                _tabControl.SelectedIndex = (int)button.Tag;
                if (this.IsHandleCreated)
                {
                    this.BeginInvoke(new Action(() => UpdateTabButtons()));
                }
            }
            else if (e.Button == MouseButtons.Right)
            {
                _selectedTab = _tabControl.TabPages[(int)button.Tag];
                _tabContextMenu.Show(button, e.Location);
            }
        }

        private void TabButton_DoubleClick(object sender, EventArgs e)
        {
            var button = sender as Button;
            if (button == null) return;

            var tabIndex = (int)button.Tag;
            var tabPage = _tabControl.TabPages[tabIndex];
            OnTabRenameRequested(tabIndex, tabPage.Text);
        }

        private void InitializeContextMenus()
        {
            _rowContextMenu = new MaterialContextMenuStrip();
            var deleteRowItem = new MaterialToolStripMenuItem { Text = "删除行" };
            deleteRowItem.Click += OnDeleteRowClick;
            var copyRowItem = new MaterialToolStripMenuItem { Text = "复制行" };
            copyRowItem.Click += OnCopyRowClick;
            var addRowAboveItem = new MaterialToolStripMenuItem { Text = "在上方插入行" };
            addRowAboveItem.Click += OnAddRowAboveClick;
            var addRowBelowItem = new MaterialToolStripMenuItem { Text = "在下方插入行" };
            addRowBelowItem.Click += OnAddRowBelowClick;
            
            _rowContextMenu.Items.Add(addRowAboveItem);
            _rowContextMenu.Items.Add(addRowBelowItem);
            _rowContextMenu.Items.Add(new ToolStripSeparator());
            _rowContextMenu.Items.Add(copyRowItem);
            _rowContextMenu.Items.Add(deleteRowItem);

            _columnContextMenu = new MaterialContextMenuStrip();
            var addColumnItem = new MaterialToolStripMenuItem { Text = "添加列" };
            addColumnItem.Click += OnAddColumnClick;
            var editColumnItem = new MaterialToolStripMenuItem { Text = "编辑列" };
            editColumnItem.Click += OnEditColumnClick;
            var deleteColumnItem = new MaterialToolStripMenuItem { Text = "删除列" };
            deleteColumnItem.Click += OnDeleteColumnClick;
            
            _columnContextMenu.Items.Add(addColumnItem);
            _columnContextMenu.Items.Add(editColumnItem);
            _columnContextMenu.Items.Add(new ToolStripSeparator());
            _columnContextMenu.Items.Add(deleteColumnItem);

            _tabContextMenu = new MaterialContextMenuStrip();
            var renameTabItem = new MaterialToolStripMenuItem { Text = "重命名" };
            renameTabItem.Click += OnRenameTabClick;
            var closeTabItem = new MaterialToolStripMenuItem { Text = "关闭标签页" };
            closeTabItem.Click += OnCloseTabClick;
            
            _tabContextMenu.Items.Add(renameTabItem);
            _tabContextMenu.Items.Add(new ToolStripSeparator());
            _tabContextMenu.Items.Add(closeTabItem);
        }

        private void SetupAutoStart()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    key?.SetValue("TaskBoard", $"\"{Application.ExecutablePath}\"");
                }
            }
            catch { }
        }

        private void OpenDataFolder()
        {
            var appDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TaskBoard");
            System.Diagnostics.Process.Start("explorer.exe", appDataFolder);
        }

        private void InitializeServices()
        {
            _dataService = new DataService();
            this.FormClosing += (s, e) => _dataService.Dispose();
        }

        private void LoadTabs()
        {
            var appData = _dataService.GetAppData();
            
            if (appData.Tabs.Count == 0)
            {
                CreateDefaultTab();
            }

            foreach (var tabData in appData.Tabs)
            {
                CreateTabPage(tabData);
            }

            if (_tabControl.TabCount > 0)
            {
                _tabControl.SelectedIndex = 0;
            }

            // 延迟更新标签按钮，确保控件已创建
            this.BeginInvoke(new Action(() => UpdateTabButtons()));
        }

        private void CreateDefaultTab()
        {
            var defaultTab = new TabData { Name = "每日计划" };
            defaultTab.Columns.Add(new ColumnDefinition { Name = "任务", Type = ColumnType.Text });
            defaultTab.Columns.Add(new ColumnDefinition 
            { 
                Name = "状态", 
                Type = ColumnType.Single,
                Options = new List<OptionItem>
                {
                    new OptionItem { Label = "未开始", Color = "#757575" },
                    new OptionItem { Label = "进行中", Color = "#FF6F00" },
                    new OptionItem { Label = "已完成", Color = "#388E3C" }
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

            var grid = new ModernDataGridView();
            grid.Dock = DockStyle.Fill;
            grid.Margin = new Padding(10);
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.CellValueChanged += Grid_CellValueChanged;
            grid.CellClick += Grid_CellClick;
            grid.MouseClick += Grid_MouseClick;
            grid.ColumnHeaderMouseClick += Grid_ColumnHeaderMouseClick;
            grid.CellPainting += Grid_CellPainting;
            grid.CellMouseClick += Grid_CellMouseClick;
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(45, 45, 48);
            grid.DefaultCellStyle.SelectionForeColor = Color.White;

            var panel = new Panel();
            panel.Dock = DockStyle.Fill;
            panel.Padding = new Padding(10);
            panel.BackColor = Color.FromArgb(37, 37, 38);

            var bottomPanel = new Panel();
            bottomPanel.Height = 40;
            bottomPanel.Dock = DockStyle.Bottom;
            bottomPanel.BackColor = Color.FromArgb(37, 37, 38);

            var addRowButton = new MaterialButton();
            addRowButton.Text = "添加新行";
            addRowButton.Location = new Point(10, 5);
            addRowButton.Click += (s, e) => AddNewRow(tabData, grid);
            bottomPanel.Controls.Add(addRowButton);

            UpdateGridColumns(grid, tabData);
            UpdateGridRows(grid, tabData);

            panel.Controls.Add(grid);
            panel.Controls.Add(bottomPanel);
            tabPage.Controls.Add(panel);
            _tabControl.TabPages.Add(tabPage);
        }

        private void Grid_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            var grid = (DataGridView)sender;
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
                        var maxY = e.CellBounds.Bottom - 5;

                        foreach (var todo in todos)
                        {
                            if (y + 20 > maxY) break;

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

        private void UpdateGridColumns(ModernDataGridView grid, TabData tabData)
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

        private void UpdateGridRows(ModernDataGridView grid, TabData tabData)
        {
            grid.Rows.Clear();

            foreach (var rowData in tabData.Rows)
            {
                var row = grid.Rows[grid.Rows.Add()];
                row.Tag = rowData;

                foreach (var col in tabData.Columns)
                {
                    if (col.Type == ColumnType.Multi) continue; // 跳过多选类型
                    
                    if (rowData.Data.ContainsKey(col.Id))
                    {
                        var value = rowData.Data[col.Id];
                        
                        if (col.Type == ColumnType.Image)
                        {
                            // 图片列不显示文本，由 CellPainting 处理
                        }
                        else if (col.Type != ColumnType.Image)
                        {
                            row.Cells[col.Id].Value = value?.ToString();
                        }
                    }
                }
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
            _dataService.NotifyDataChanged();
        }

        private void Grid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

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
        }

        private void ShowModalDialog(Form dialog)
        {
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

        private void ShowSingleSelectDialog(RowData rowData, ColumnDefinition col, DataGridView grid, int rowIndex, int colIndex)
        {
            var dialog = new Form();
            dialog.Text = $"选择{col.Name}";
            dialog.Size = new Size(300, 400);
            dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
            dialog.MaximizeBox = false;
            dialog.MinimizeBox = false;
            dialog.BackColor = Color.FromArgb(37, 37, 38);

            var listBox = new ListBox();
            listBox.Location = new Point(20, 20);
            listBox.Size = new Size(260, 300);
            listBox.BackColor = Color.FromArgb(45, 45, 48);
            listBox.ForeColor = Color.White;
            listBox.BorderStyle = BorderStyle.None;
            listBox.Font = new Font("Microsoft YaHei UI", 10);
            
            var currentValue = rowData.Data.ContainsKey(col.Id) ? rowData.Data[col.Id]?.ToString() : "";
            
            foreach (var option in col.Options)
            {
                var index = listBox.Items.Add(option.Label);
                if (option.Label == currentValue)
                {
                    listBox.SelectedIndex = index;
                }
            }

            listBox.DrawMode = DrawMode.OwnerDrawFixed;
            listBox.ItemHeight = 30;
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
                var tagRect = new Rectangle(e.Bounds.X + 10, e.Bounds.Y + 5, 20, 20);
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

            listBox.SelectedIndexChanged += (s, e) =>
            {
                if (listBox.SelectedIndex >= 0)
                {
                    var selectedOption = col.Options[listBox.SelectedIndex];
                    rowData.Data[col.Id] = selectedOption.Label;
                    grid.Rows[rowIndex].Cells[colIndex].Value = selectedOption.Label;
                    _dataService.NotifyDataChanged();
                    grid.InvalidateCell(colIndex, rowIndex);
                    dialog.Close();
                }
            };

            var cancelButton = new Button();
            cancelButton.Text = "取消";
            cancelButton.Size = new Size(80, 30);
            cancelButton.Location = new Point(200, 330);
            cancelButton.FlatStyle = FlatStyle.Flat;
            cancelButton.BackColor = Color.FromArgb(60, 60, 60);
            cancelButton.ForeColor = Color.White;
            cancelButton.Click += (s, e) => dialog.Close();

            dialog.Controls.Add(listBox);
            dialog.Controls.Add(cancelButton);

            ShowModalDialog(dialog);
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
                addButton.Text = "+\n添加图片";
                addButton.Font = new Font("Arial", 20);
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
                _dataService.NotifyDataChanged();
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
                grid.InvalidateCell(grid.Columns[colId].Index, rowIndex);
                dialog.Close();
            };

            bottomPanel.Controls.Add(okButton);

            dialog.Controls.Add(todoControl);
            dialog.Controls.Add(bottomPanel);

            ShowModalDialog(dialog);
        }

        private void Grid_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
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
                    var message = todo.IsCompleted ? "确认取消完成状态？" : "确认完成此待办事项？";
                    var result = MessageBox.Show(message, "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        todo.IsCompleted = !todo.IsCompleted;
                        rowData.Data[col.Id] = todos;
                        _dataService.NotifyDataChanged();
                        grid.InvalidateCell(e.ColumnIndex, e.RowIndex);
                    }
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
            if (e.Button == MouseButtons.Right)
            {
                _selectedColumnIndex = e.ColumnIndex;
                var grid = (DataGridView)sender;
                _columnContextMenu.Show(grid, grid.PointToClient(Cursor.Position));
            }
        }

        private void OnAddTabClicked(object sender, EventArgs e)
        {
            using (var dialog = new InputDialog("新建标签页", "请输入标签页名称:", $"新标签页 {_tabControl.TabCount + 1}"))
            {
                if (dialog.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.InputValue))
                {
                    var tabData = new TabData { Name = dialog.InputValue };
                    tabData.Columns.Add(new ColumnDefinition { Name = "任务", Type = ColumnType.Text });
                    
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
            using (var dialog = new InputDialog("重命名标签页", "请输入新的标签页名称:", currentName))
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

            using (var dialog = new InputDialog("重命名标签页", "请输入新的标签页名称:", _selectedTab.Text))
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
            _dataService.NotifyDataChanged();
        }

        private void OnDeleteRowClick(object sender, EventArgs e)
        {
            if (_selectedRow == null) return;

            var result = MessageBox.Show("确定要删除选中的行吗？", "确认删除", 
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            
            if (result == DialogResult.Yes)
            {
                var tabData = GetCurrentTabData();
                var rowData = (RowData)_selectedRow.Tag;
                tabData.Rows.Remove(rowData);
                
                var grid = GetCurrentGrid();
                UpdateGridRows(grid, tabData);
                _dataService.NotifyDataChanged();
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
            using (var dialog = new ColumnEditDialog2())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var tabData = GetCurrentTabData();
                    var grid = GetCurrentGrid();
                    
                    tabData.Columns.Add(dialog.GetColumnDefinition());
                    UpdateGridColumns(grid, tabData);
                    UpdateGridRows(grid, tabData);
                    _dataService.NotifyDataChanged();
                }
            }
        }

        private void OnEditColumnClick(object sender, EventArgs e)
        {
            var grid = GetCurrentGrid();
            var col = (ColumnDefinition)grid.Columns[_selectedColumnIndex].Tag;
            
            using (var dialog = new ColumnEditDialog2(col))
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var tabData = GetCurrentTabData();
                    UpdateGridColumns(grid, tabData);
                    UpdateGridRows(grid, tabData);
                    _dataService.NotifyDataChanged();
                }
            }
        }

        private void OnDeleteColumnClick(object sender, EventArgs e)
        {
            var grid = GetCurrentGrid();
            var col = (ColumnDefinition)grid.Columns[_selectedColumnIndex].Tag;
            
            var result = MessageBox.Show($"确定要删除列 \"{col.Name}\" 吗？\n这将删除该列的所有数据。", 
                "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            
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
                _dataService.NotifyDataChanged();
            }
        }

        private TabData GetCurrentTabData()
        {
            return (TabData)_tabControl.SelectedTab?.Tag;
        }

        private ModernDataGridView GetCurrentGrid()
        {
            var panel = _tabControl.SelectedTab?.Controls[0] as Panel;
            return panel?.Controls.OfType<ModernDataGridView>().FirstOrDefault();
        }
    }

    // 更新的列编辑对话框，去掉多选类型
    public class ColumnEditDialog2 : Form
    {
        private TextBox _nameTextBox;
        private ComboBox _typeComboBox;
        private ListBox _optionsListBox;
        private Button _addOptionButton;
        private Button _removeOptionButton;
        private Panel _optionsPanel;
        private ColumnDefinition _column;

        public ColumnEditDialog2(ColumnDefinition column = null)
        {
            _column = column ?? new ColumnDefinition();
            InitializeComponent();
            LoadColumnData();
        }

        private void InitializeComponent()
        {
            this.Text = "编辑列";
            this.Size = new Size(400, 500);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = Color.FromArgb(37, 37, 38);
            this.ForeColor = Color.White;

            var nameLabel = new Label();
            nameLabel.Text = "列名:";
            nameLabel.Location = new Point(12, 15);
            nameLabel.Size = new Size(60, 23);

            _nameTextBox = new TextBox();
            _nameTextBox.Location = new Point(80, 12);
            _nameTextBox.Size = new Size(280, 23);
            _nameTextBox.BackColor = Color.FromArgb(45, 45, 48);
            _nameTextBox.ForeColor = Color.White;
            _nameTextBox.BorderStyle = BorderStyle.FixedSingle;

            var typeLabel = new Label();
            typeLabel.Text = "类型:";
            typeLabel.Location = new Point(12, 50);
            typeLabel.Size = new Size(60, 23);

            _typeComboBox = new ComboBox();
            _typeComboBox.Location = new Point(80, 47);
            _typeComboBox.Size = new Size(280, 23);
            _typeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            _typeComboBox.Items.AddRange(new[] { "文本输入", "单选状态", "图片", "待办列表" });
            _typeComboBox.BackColor = Color.FromArgb(45, 45, 48);
            _typeComboBox.ForeColor = Color.White;
            _typeComboBox.SelectedIndexChanged += TypeComboBox_SelectedIndexChanged;

            _optionsPanel = new Panel();
            _optionsPanel.Location = new Point(12, 85);
            _optionsPanel.Size = new Size(360, 300);
            _optionsPanel.BorderStyle = BorderStyle.FixedSingle;

            var optionsLabel = new Label();
            optionsLabel.Text = "选项列表:";
            optionsLabel.Location = new Point(0, 0);
            optionsLabel.Size = new Size(100, 23);

            _optionsListBox = new ListBox();
            _optionsListBox.Location = new Point(0, 25);
            _optionsListBox.Size = new Size(358, 220);
            _optionsListBox.BackColor = Color.FromArgb(45, 45, 48);
            _optionsListBox.ForeColor = Color.White;
            _optionsListBox.BorderStyle = BorderStyle.None;

            _addOptionButton = new Button();
            _addOptionButton.Text = "添加选项";
            _addOptionButton.Location = new Point(0, 250);
            _addOptionButton.Size = new Size(100, 30);
            _addOptionButton.FlatStyle = FlatStyle.Flat;
            _addOptionButton.BackColor = Color.FromArgb(0, 122, 204);
            _addOptionButton.Click += AddOptionButton_Click;

            _removeOptionButton = new Button();
            _removeOptionButton.Text = "删除选项";
            _removeOptionButton.Location = new Point(110, 250);
            _removeOptionButton.Size = new Size(100, 30);
            _removeOptionButton.FlatStyle = FlatStyle.Flat;
            _removeOptionButton.BackColor = Color.FromArgb(60, 60, 60);
            _removeOptionButton.Click += RemoveOptionButton_Click;

            _optionsPanel.Controls.Add(optionsLabel);
            _optionsPanel.Controls.Add(_optionsListBox);
            _optionsPanel.Controls.Add(_addOptionButton);
            _optionsPanel.Controls.Add(_removeOptionButton);

            var okButton = new Button();
            okButton.Text = "确定";
            okButton.Location = new Point(210, 400);
            okButton.Size = new Size(75, 30);
            okButton.DialogResult = DialogResult.OK;
            okButton.FlatStyle = FlatStyle.Flat;
            okButton.BackColor = Color.FromArgb(0, 122, 204);
            okButton.Click += OkButton_Click;

            var cancelButton = new Button();
            cancelButton.Text = "取消";
            cancelButton.Location = new Point(295, 400);
            cancelButton.Size = new Size(75, 30);
            cancelButton.DialogResult = DialogResult.Cancel;
            cancelButton.FlatStyle = FlatStyle.Flat;
            cancelButton.BackColor = Color.FromArgb(60, 60, 60);

            this.Controls.Add(nameLabel);
            this.Controls.Add(_nameTextBox);
            this.Controls.Add(typeLabel);
            this.Controls.Add(_typeComboBox);
            this.Controls.Add(_optionsPanel);
            this.Controls.Add(okButton);
            this.Controls.Add(cancelButton);
        }

        private void LoadColumnData()
        {
            _nameTextBox.Text = _column.Name;
            
            // 映射类型
            switch (_column.Type)
            {
                case ColumnType.Text:
                    _typeComboBox.SelectedIndex = 0;
                    break;
                case ColumnType.Single:
                    _typeComboBox.SelectedIndex = 1;
                    break;
                case ColumnType.Image:
                    _typeComboBox.SelectedIndex = 2;
                    break;
                case ColumnType.TodoList:
                    _typeComboBox.SelectedIndex = 3;
                    break;
                default:
                    _typeComboBox.SelectedIndex = 0;
                    break;
            }
            
            foreach (var option in _column.Options)
            {
                _optionsListBox.Items.Add($"{option.Label} - {option.Color}");
            }
            
            UpdateOptionsVisibility();
        }

        private void TypeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateOptionsVisibility();
        }

        private void UpdateOptionsVisibility()
        {
            var needsOptions = _typeComboBox.SelectedIndex == 1; // 只有单选需要选项
            _optionsPanel.Visible = needsOptions;
        }

        private void AddOptionButton_Click(object sender, EventArgs e)
        {
            using (var dialog = new OptionEditDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var option = dialog.GetOption();
                    _column.Options.Add(option);
                    _optionsListBox.Items.Add($"{option.Label} - {option.Color}");
                }
            }
        }

        private void RemoveOptionButton_Click(object sender, EventArgs e)
        {
            if (_optionsListBox.SelectedIndex >= 0)
            {
                _column.Options.RemoveAt(_optionsListBox.SelectedIndex);
                _optionsListBox.Items.RemoveAt(_optionsListBox.SelectedIndex);
            }
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            _column.Name = _nameTextBox.Text;
            
            // 映射类型
            switch (_typeComboBox.SelectedIndex)
            {
                case 0:
                    _column.Type = ColumnType.Text;
                    break;
                case 1:
                    _column.Type = ColumnType.Single;
                    break;
                case 2:
                    _column.Type = ColumnType.Image;
                    break;
                case 3:
                    _column.Type = ColumnType.TodoList;
                    break;
            }
        }

        public ColumnDefinition GetColumnDefinition() => _column;
    }
}