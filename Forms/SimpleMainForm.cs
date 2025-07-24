using System;
using System.Drawing;
using System.Windows.Forms;
using TaskBoard.Models;
using TaskBoard.Services;

namespace TaskBoard
{
    public class SimpleMainForm : Form
    {
        private DataService _dataService;
        private TabControl _tabControl;

        public SimpleMainForm()
        {
            try
            {
                InitializeComponent();
                InitializeServices();
                LoadTabs();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化失败: {ex.Message}", "错误");
            }
        }

        private void InitializeComponent()
        {
            this.Text = "TaskBoard 2.0 (简化版)";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;

            _tabControl = new TabControl();
            _tabControl.Dock = DockStyle.Fill;

            var statusLabel = new Label();
            statusLabel.Text = "状态: 正在运行";
            statusLabel.Dock = DockStyle.Bottom;
            statusLabel.Height = 25;
            statusLabel.BackColor = Color.LightGray;

            this.Controls.Add(_tabControl);
            this.Controls.Add(statusLabel);
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
                var defaultTab = new TabData { Name = "默认标签页" };
                appData.Tabs.Add(defaultTab);
                _dataService.SaveData();
            }

            foreach (var tabData in appData.Tabs)
            {
                var tabPage = new TabPage(tabData.Name);
                tabPage.Tag = tabData;
                
                var label = new Label();
                label.Text = $"标签页: {tabData.Name}";
                label.AutoSize = true;
                label.Location = new Point(20, 20);
                
                tabPage.Controls.Add(label);
                _tabControl.TabPages.Add(tabPage);
            }
        }
    }
}