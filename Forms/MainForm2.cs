using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Collections.Generic;
using MaterialSkin;
using MaterialSkin.Controls;
using TaskBoard.Controls;
using TaskBoard.Forms;
using TaskBoard.Models;
using TaskBoard.Services;

namespace TaskBoard
{
    public partial class MainForm2 : MaterialForm
    {
        private DataService _dataService;
        private ModernTabControl _tabControl;
        private MaterialContextMenuStrip _rowContextMenu;
        private MaterialContextMenuStrip _columnContextMenu;
        private DataGridViewRow _selectedRow;
        private int _selectedColumnIndex;

        public MainForm2()
        {
            InitializeComponent();
            InitializeMaterialDesign();
            InitializeServices();
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

            InitializeContextMenus();

            this.Controls.Add(_tabControl);
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
            grid.CellValueChanged += Grid_CellValueChanged;
            grid.CellClick += Grid_CellClick;
            grid.MouseClick += Grid_MouseClick;
            grid.ColumnHeaderMouseClick += Grid_ColumnHeaderMouseClick;

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
                        var btnCol = new DataGridViewButtonColumn();
                        btnCol.Text = "📷 查看";
                        btnCol.UseColumnTextForButtonValue = true;
                        btnCol.FlatStyle = FlatStyle.Flat;
                        gridCol = btnCol;
                        break;
                }

                if (gridCol != null)
                {
                    gridCol.Name = col.Id;
                    gridCol.HeaderText = col.Name;
                    gridCol.Tag = col;
                    gridCol.Width = 150;
                    gridCol.SortMode = DataGridViewColumnSortMode.NotSortable;
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
                ShowImageDialog(rowData, col.Id);
            }
            else if (col.Type == ColumnType.Multi)
            {
                var rowData = (RowData)grid.Rows[e.RowIndex].Tag;
                ShowMultiSelectDialog(rowData, col, grid, e.RowIndex, e.ColumnIndex);
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

        private void ShowImageDialog(RowData rowData, string colId)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "图片文件|*.jpg;*.jpeg;*.png;*.gif;*.bmp";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var fileName = $"{DateTime.Now:yyyyMMdd-HHmmss}-{System.IO.Path.GetFileName(dialog.FileName)}";
                    var destPath = System.IO.Path.Combine(_dataService.GetImagePath(), fileName);
                    System.IO.File.Copy(dialog.FileName, destPath, true);
                    
                    rowData.Data[colId] = fileName;
                    _dataService.NotifyDataChanged();
                    
                    MessageBox.Show("图片已上传", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
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