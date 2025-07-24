using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using TaskBoard.Models;
using TaskBoard.Services;

namespace TaskBoard
{
    public partial class MainForm : Form
    {
        private DataService _dataService;
        private TabControl _tabControl;
        private ToolStrip _toolStrip;

        public MainForm()
        {
            InitializeComponent();
            InitializeServices();
            LoadTabs();
        }

        private void InitializeComponent()
        {
            this.Text = "TaskBoard - 任务管理";
            this.Size = new Size(1200, 800);
            this.StartPosition = FormStartPosition.CenterScreen;

            _toolStrip = new ToolStrip();
            _toolStrip.Items.Add(new ToolStripButton("新建标签页", null, OnNewTab));
            _toolStrip.Items.Add(new ToolStripButton("设置", null, OnSettings));
            _toolStrip.Items.Add(new ToolStripSeparator());
            _toolStrip.Items.Add(new ToolStripButton("添加行", null, OnAddRow));
            _toolStrip.Items.Add(new ToolStripButton("删除行", null, OnDeleteRow));
            _toolStrip.Items.Add(new ToolStripSeparator());
            _toolStrip.Items.Add(new ToolStripButton("添加列", null, OnAddColumn));
            _toolStrip.Items.Add(new ToolStripButton("编辑列", null, OnEditColumn));

            _tabControl = new TabControl();
            _tabControl.Dock = DockStyle.Fill;

            this.Controls.Add(_tabControl);
            this.Controls.Add(_toolStrip);
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
                var defaultTab = new TabData { Name = "每日计划" };
                defaultTab.Columns.Add(new ColumnDefinition { Name = "任务", Type = ColumnType.Text });
                defaultTab.Columns.Add(new ColumnDefinition 
                { 
                    Name = "状态", 
                    Type = ColumnType.Single,
                    Options = new System.Collections.Generic.List<OptionItem>
                    {
                        new OptionItem { Label = "未开始", Color = "#CCCCCC" },
                        new OptionItem { Label = "进行中", Color = "#FF9900" },
                        new OptionItem { Label = "已完成", Color = "#00CC00" }
                    }
                });
                appData.Tabs.Add(defaultTab);
                _dataService.SaveData();
            }

            foreach (var tabData in appData.Tabs)
            {
                CreateTabPage(tabData);
            }
        }

        private void CreateTabPage(TabData tabData)
        {
            var tabPage = new TabPage(tabData.Name);
            tabPage.Tag = tabData;

            var grid = new DataGridView();
            grid.Dock = DockStyle.Fill;
            grid.AllowUserToAddRows = false;
            grid.AutoGenerateColumns = false;
            grid.CellValueChanged += Grid_CellValueChanged;
            grid.CellClick += Grid_CellClick;

            UpdateGridColumns(grid, tabData);
            UpdateGridRows(grid, tabData);

            tabPage.Controls.Add(grid);
            _tabControl.TabPages.Add(tabPage);
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
                        var comboCol = new DataGridViewComboBoxColumn();
                        foreach (var option in col.Options)
                        {
                            comboCol.Items.Add(option.Label);
                        }
                        gridCol = comboCol;
                        break;
                    case ColumnType.Multi:
                        gridCol = new DataGridViewTextBoxColumn();
                        break;
                    case ColumnType.Image:
                        gridCol = new DataGridViewButtonColumn();
                        ((DataGridViewButtonColumn)gridCol).Text = "查看图片";
                        ((DataGridViewButtonColumn)gridCol).UseColumnTextForButtonValue = true;
                        break;
                }

                if (gridCol != null)
                {
                    gridCol.Name = col.Id;
                    gridCol.HeaderText = col.Name;
                    gridCol.Tag = col;
                    gridCol.Width = 150;
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
            var tabPage = (TabPage)grid.Parent;
            var tabData = (TabData)tabPage.Tag;
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

        private void ShowImageDialog(RowData rowData, string colId)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "图片文件|*.jpg;*.jpeg;*.png;*.gif;*.bmp";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var fileName = $"{DateTime.Now:yyyyMMdd-HHmmss}-{Path.GetFileName(dialog.FileName)}";
                    var destPath = Path.Combine(_dataService.GetImagePath(), fileName);
                    File.Copy(dialog.FileName, destPath, true);
                    
                    rowData.Data[colId] = fileName;
                    _dataService.NotifyDataChanged();
                    
                    MessageBox.Show("图片已上传", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void ShowMultiSelectDialog(RowData rowData, ColumnDefinition col, DataGridView grid, int rowIndex, int colIndex)
        {
            var dialog = new Form();
            dialog.Text = $"选择{col.Name}";
            dialog.Size = new Size(300, 400);
            dialog.StartPosition = FormStartPosition.CenterParent;

            var checkedListBox = new CheckedListBox();
            checkedListBox.Dock = DockStyle.Fill;
            
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

            var btnPanel = new Panel();
            btnPanel.Height = 40;
            btnPanel.Dock = DockStyle.Bottom;

            var btnOK = new Button();
            btnOK.Text = "确定";
            btnOK.DialogResult = DialogResult.OK;
            btnOK.Location = new Point(100, 5);

            btnPanel.Controls.Add(btnOK);
            dialog.Controls.Add(checkedListBox);
            dialog.Controls.Add(btnPanel);

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                var selected = new System.Collections.Generic.List<string>();
                foreach (var item in checkedListBox.CheckedItems)
                {
                    selected.Add(item.ToString());
                }
                
                rowData.Data[col.Id] = selected;
                grid.Rows[rowIndex].Cells[colIndex].Value = string.Join(", ", selected);
                _dataService.NotifyDataChanged();
            }
        }

        private void OnNewTab(object sender, EventArgs e)
        {
            var tabData = new TabData { Name = $"新标签页 {_tabControl.TabCount + 1}" };
            tabData.Columns.Add(new ColumnDefinition { Name = "任务", Type = ColumnType.Text });
            
            _dataService.GetAppData().Tabs.Add(tabData);
            CreateTabPage(tabData);
            _dataService.NotifyDataChanged();
        }

        private void OnSettings(object sender, EventArgs e)
        {
            using (var dialog = new SettingsDialog())
            {
                dialog.ShowDialog();
            }
        }

        private void OnAddRow(object sender, EventArgs e)
        {
            if (_tabControl.SelectedTab == null) return;

            var tabData = (TabData)_tabControl.SelectedTab.Tag;
            var grid = (DataGridView)_tabControl.SelectedTab.Controls[0];
            
            var newRow = new RowData();
            tabData.Rows.Add(newRow);
            
            UpdateGridRows(grid, tabData);
            _dataService.NotifyDataChanged();
        }

        private void OnDeleteRow(object sender, EventArgs e)
        {
            if (_tabControl.SelectedTab == null) return;

            var grid = (DataGridView)_tabControl.SelectedTab.Controls[0];
            if (grid.CurrentRow == null) return;

            var tabData = (TabData)_tabControl.SelectedTab.Tag;
            var rowData = (RowData)grid.CurrentRow.Tag;
            
            tabData.Rows.Remove(rowData);
            UpdateGridRows(grid, tabData);
            _dataService.NotifyDataChanged();
        }

        private void OnAddColumn(object sender, EventArgs e)
        {
            if (_tabControl.SelectedTab == null) return;

            using (var dialog = new ColumnEditDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var tabData = (TabData)_tabControl.SelectedTab.Tag;
                    var grid = (DataGridView)_tabControl.SelectedTab.Controls[0];
                    
                    tabData.Columns.Add(dialog.GetColumnDefinition());
                    UpdateGridColumns(grid, tabData);
                    UpdateGridRows(grid, tabData);
                    _dataService.NotifyDataChanged();
                }
            }
        }

        private void OnEditColumn(object sender, EventArgs e)
        {
            if (_tabControl.SelectedTab == null) return;

            var grid = (DataGridView)_tabControl.SelectedTab.Controls[0];
            if (grid.CurrentCell == null) return;

            var col = (ColumnDefinition)grid.Columns[grid.CurrentCell.ColumnIndex].Tag;
            
            using (var dialog = new ColumnEditDialog(col))
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var tabData = (TabData)_tabControl.SelectedTab.Tag;
                    UpdateGridColumns(grid, tabData);
                    UpdateGridRows(grid, tabData);
                    _dataService.NotifyDataChanged();
                }
            }
        }
    }
}