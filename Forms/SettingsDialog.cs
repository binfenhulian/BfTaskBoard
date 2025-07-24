using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;

namespace TaskBoard
{
    public class SettingsDialog : Form
    {
        private CheckBox _autoStartCheckBox;
        private const string AppName = "TaskBoard";
        private readonly string _executablePath = Application.ExecutablePath;

        public SettingsDialog()
        {
            InitializeComponent();
            LoadSettings();
        }

        private void InitializeComponent()
        {
            this.Text = "设置";
            this.Size = new Size(400, 300);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var groupBox = new GroupBox();
            groupBox.Text = "启动选项";
            groupBox.Location = new Point(12, 12);
            groupBox.Size = new Size(360, 100);

            _autoStartCheckBox = new CheckBox();
            _autoStartCheckBox.Text = "开机自动启动";
            _autoStartCheckBox.Location = new Point(20, 30);
            _autoStartCheckBox.Size = new Size(300, 23);
            _autoStartCheckBox.CheckedChanged += AutoStartCheckBox_CheckedChanged;

            groupBox.Controls.Add(_autoStartCheckBox);

            var infoLabel = new Label();
            infoLabel.Text = "数据保存位置:\n" + 
                            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\TaskBoard";
            infoLabel.Location = new Point(12, 130);
            infoLabel.Size = new Size(360, 60);

            var okButton = new Button();
            okButton.Text = "确定";
            okButton.Location = new Point(297, 220);
            okButton.Size = new Size(75, 30);
            okButton.DialogResult = DialogResult.OK;

            this.Controls.Add(groupBox);
            this.Controls.Add(infoLabel);
            this.Controls.Add(okButton);
        }

        private void LoadSettings()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", false))
                {
                    _autoStartCheckBox.Checked = key?.GetValue(AppName) != null;
                }
            }
            catch
            {
                _autoStartCheckBox.Checked = false;
            }
        }

        private void AutoStartCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (_autoStartCheckBox.Checked)
                    {
                        key?.SetValue(AppName, $"\"{_executablePath}\"");
                    }
                    else
                    {
                        key?.DeleteValue(AppName, false);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置开机启动失败: {ex.Message}", "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                _autoStartCheckBox.Checked = !_autoStartCheckBox.Checked;
            }
        }
    }
}