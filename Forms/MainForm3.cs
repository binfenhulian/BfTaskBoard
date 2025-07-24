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

namespace TaskBoard
{
    public partial class MainForm3 : MaterialForm
    {
        private DataService _dataService;
        private ModernTabControl _tabControl;
        private MaterialContextMenuStrip _rowContextMenu;
        private MaterialContextMenuStrip _columnContextMenu;
        private MaterialContextMenuStrip _tabContextMenu;
        private DataGridViewRow _selectedRow;
        private int _selectedColumnIndex;
        private TabPage _selectedTab;

        public MainForm3()
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

            _tabControl = new ModernTabControl();
            _tabControl.Dock = DockStyle.Fill;
            _tabControl.Font = new Font("Microsoft YaHei UI", 9F);
            _tabControl.AddTabClicked += OnAddTabClicked;
            _tabControl.TabRenameRequested += OnTabRenameRequested;
            _tabControl.MouseClick += TabControl_MouseClick;

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
            this.Controls.Add(bottomPanel);
        }

        private void TabControl_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                for (int i = 0; i < _tabControl.TabCount; i++)
                {
                    var tabRect = _tabControl.GetTabRect(i);
                    if (tabRect.Contains(e.Location))
                    {
                        _selectedTab = _tabControl.TabPages[i];
                        _tabContextMenu.Show(_tabControl, e.Location);
                        break;
                    }
                }
            }
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
            defaultTab.Columns.Add(new ColumnDefinition 
            { 
                Name = "标签", 
                Type = ColumnType.Multi,
                Options = new List<OptionItem>
                {
                    new OptionItem { Label = "重要", Color = "#D32F2F" },
                    new OptionItem { Label = "紧急", Color = "#F57C00" },
                    new OptionItem { Label = "学习", Color = "#1976D2" }
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

            if (col.Type == ColumnType.Single || col.Type == ColumnType.Multi)
            {
                e.Paint(e.CellBounds, DataGridViewPaintParts.Background | DataGridViewPaintParts.Border);
                e.Handled = true;

                var value = e.FormattedValue?.ToString();
                if (string.IsNullOrEmpty(value)) return;

                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;

                var items = col.Type == ColumnType.Single ? new[] { value } : value.Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries);
                var x = e.CellBounds.X + 5;
                var y = e.CellBounds.Y + (e.CellBounds.Height - 20) / 2;

                foreach (var item in items)
                {
                    var option = col.Options.FirstOrDefault(o => o.Label == item);
                    if (option == null) continue;

                    var textSize = g.MeasureString(item, grid.Font);
                    var tagWidth = (int)textSize.Width + 16;
                    var tagRect = new Rectangle(x, y, tagWidth, 20);

                    using (var brush = new SolidBrush(ColorTranslator.FromHtml(option.Color)))
                    using (var path = GetRoundedRectPath(tagRect, 10))
                    {
                        g.FillPath(brush, path);
                    }

                    using (var brush = new SolidBrush(GetContrastColor(ColorTranslator.FromHtml(option.Color))))
                    {
                        var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                        g.DrawString(item, grid.Font, brush, tagRect, format);
                    }

                    x += tagWidth + 5;
                    if (x + 100 > e.CellBounds.Right) break;
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
                        var comboCol = new DataGridViewComboBoxColumn();
                        comboCol.FlatStyle = FlatStyle.Flat;
                        comboCol.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
                        foreach (var option in col.Options)
                        {
                            comboCol.Items.Add(option.Label);
                        }
                        gridCol = comboCol;
                        break;
                    case ColumnType.Multi:
                        gridCol = new DataGridViewTextBoxColumn();
                        gridCol.ReadOnly = true;
                        break;
                    case ColumnType.Image:
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
                    if (rowData.Data.ContainsKey(col.Id))
                    {
                        var value = rowData.Data[col.Id];
                        
                        if (col.Type == ColumnType.Multi && value is Newtonsoft.Json.Linq.JArray jArray)
                        {
                            row.Cells[col.Id].Value = string.Join(", ", jArray.ToObject<string[]>());
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
            else if (col.Type == ColumnType.Multi)
            {
                var rowData = (RowData)grid.Rows[e.RowIndex].Tag;
                ShowMultiSelectDialog(rowData, col, grid, e.RowIndex, e.ColumnIndex);
            }
        }

        private void ShowImageManagementDialog(RowData rowData, string colId, DataGridView grid, int rowIndex)
        {
            var dialog = new Form();
            dialog.Text = "图片管理";
            dialog.Size = new Size(800, 600);
            dialog.StartPosition = FormStartPosition.CenterParent;
            dialog.BackColor = Color.FromArgb(37, 37, 38);

            var flowPanel = new FlowLayoutPanel();
            flowPanel.Dock = DockStyle.Fill;
            flowPanel.AutoScroll = true;
            flowPanel.Padding = new Padding(10);
            flowPanel.BackColor = Color.FromArgb(37, 37, 38);

            var images = rowData.Data.ContainsKey(colId) ? 
                (rowData.Data[colId] as List<string> ?? new List<string>()) : 
                new List<string>();

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

                    var fullPath = Path.Combine(_dataService.GetImagePath(), imagePath);
                    if (File.Exists(fullPath))
                    {
                        using (var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read))
                        {
                            picBox.Image = Image.FromStream(stream);
                        }
                    }

                    picBox.Click += (s, e) =>
                    {
                        var viewer = new Form();
                        viewer.Text = "查看图片";
                        viewer.Size = new Size(800, 600);
                        viewer.StartPosition = FormStartPosition.CenterParent;
                        var viewPic = new PictureBox();
                        viewPic.Dock = DockStyle.Fill;
                        viewPic.SizeMode = PictureBoxSizeMode.Zoom;
                        viewPic.Image = picBox.Image;
                        viewer.Controls.Add(viewPic);
                        viewer.ShowDialog();
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

                var addPanel = new Panel();
                addPanel.Size = new Size(150, 150);
                addPanel.BackColor = Color.FromArgb(45, 45, 48);
                addPanel.Margin = new Padding(5);
                addPanel.Cursor = Cursors.Hand;

                var addLabel = new Label();
                addLabel.Text = "+\n添加图片";
                addLabel.Dock = DockStyle.Fill;
                addLabel.TextAlign = ContentAlignment.MiddleCenter;
                addLabel.ForeColor = Color.White;
                addLabel.Font = new Font("Arial", 20);

                addPanel.Controls.Add(addLabel);
                addPanel.Click += (s, e) =>
                {
                    using (var openDialog = new OpenFileDialog())
                    {
                        openDialog.Filter = "图片文件|*.jpg;*.jpeg;*.png;*.gif;*.bmp";
                        openDialog.Multiselect = true;
                        if (openDialog.ShowDialog() == DialogResult.OK)
                        {
                            foreach (var file in openDialog.FileNames)
                            {
                                var fileName = $"{DateTime.Now:yyyyMMdd-HHmmss}-{Path.GetFileName(file)}";
                                var destPath = Path.Combine(_dataService.GetImagePath(), fileName);
                                File.Copy(file, destPath, true);
                                images.Add(fileName);
                            }
                            RefreshImages();
                        }
                    }
                };

                flowPanel.Controls.Add(addPanel);
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
            dialog.ShowDialog();
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
                }
            }
        }

        private void OnTabRenameRequested(object sender, TabRenameEventArgs e)
        {
            using (var dialog = new InputDialog("重命名标签页", "请输入新的标签页名称:", e.CurrentName))
            {
                if (dialog.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.InputValue))
                {
                    var tabData = (TabData)_tabControl.TabPages[e.TabIndex].Tag;
                    tabData.Name = dialog.InputValue;
                    _tabControl.TabPages[e.TabIndex].Text = dialog.InputValue;
                    _dataService.NotifyDataChanged();
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
            using (var dialog = new ColumnEditDialog())
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
            
            using (var dialog = new ColumnEditDialog(col))
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

        private void ShowMultiSelectDialog(RowData rowData, ColumnDefinition col, DataGridView grid, int rowIndex, int colIndex)
        {
            var dialog = new MaterialForm();
            dialog.Text = $"选择{col.Name}";
            dialog.Size = new Size(350, 450);
            dialog.StartPosition = FormStartPosition.CenterParent;
            dialog.MaximizeBox = false;
            dialog.MinimizeBox = false;
            dialog.Sizable = false;

            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(dialog);

            var checkedListBox = new CheckedListBox();
            checkedListBox.Location = new Point(20, 80);
            checkedListBox.Size = new Size(310, 280);
            checkedListBox.BackColor = Color.FromArgb(50, 50, 50);
            checkedListBox.ForeColor = Color.White;
            checkedListBox.BorderStyle = BorderStyle.None;
            
            foreach (var option in col.Options)
            {
                checkedListBox.Items.Add(option.Label);
            }

            if (rowData.Data.ContainsKey(col.Id) && rowData.Data[col.Id] is Newtonsoft.Json.Linq.JArray jArray)
            {
                var selected = jArray.ToObject<string[]>();
                foreach (var item in selected)
                {
                    var index = checkedListBox.Items.IndexOf(item);
                    if (index >= 0) checkedListBox.SetItemChecked(index, true);
                }
            }

            var btnOK = new MaterialButton();
            btnOK.Text = "确定";
            btnOK.Location = new Point(175, 380);
            btnOK.Click += (s, e) =>
            {
                var selected = new List<string>();
                foreach (var item in checkedListBox.CheckedItems)
                {
                    selected.Add(item.ToString());
                }
                
                rowData.Data[col.Id] = selected;
                grid.Rows[rowIndex].Cells[colIndex].Value = string.Join(", ", selected);
                _dataService.NotifyDataChanged();
                grid.InvalidateCell(colIndex, rowIndex);
                dialog.Close();
            };

            var btnCancel = new MaterialButton();
            btnCancel.Text = "取消";
            btnCancel.Location = new Point(255, 380);
            btnCancel.Click += (s, e) => dialog.Close();

            dialog.Controls.Add(checkedListBox);
            dialog.Controls.Add(btnOK);
            dialog.Controls.Add(btnCancel);

            dialog.ShowDialog();
        }
    }
}