using System;
using System.Drawing;
using System.Windows.Forms;
using TaskBoard.Models;

namespace TaskBoard
{
    public class ColumnEditDialog : Form
    {
        private TextBox _nameTextBox;
        private ComboBox _typeComboBox;
        private ListBox _optionsListBox;
        private Button _addOptionButton;
        private Button _removeOptionButton;
        private Panel _optionsPanel;
        private ColumnDefinition _column;

        public ColumnEditDialog(ColumnDefinition column = null)
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

            var nameLabel = new Label();
            nameLabel.Text = "列名:";
            nameLabel.Location = new Point(12, 15);
            nameLabel.Size = new Size(60, 23);

            _nameTextBox = new TextBox();
            _nameTextBox.Location = new Point(80, 12);
            _nameTextBox.Size = new Size(280, 23);

            var typeLabel = new Label();
            typeLabel.Text = "类型:";
            typeLabel.Location = new Point(12, 50);
            typeLabel.Size = new Size(60, 23);

            _typeComboBox = new ComboBox();
            _typeComboBox.Location = new Point(80, 47);
            _typeComboBox.Size = new Size(280, 23);
            _typeComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            _typeComboBox.Items.AddRange(new[] { "文本输入", "单选状态", "多选标签", "图片" });
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

            _addOptionButton = new Button();
            _addOptionButton.Text = "添加选项";
            _addOptionButton.Location = new Point(0, 250);
            _addOptionButton.Size = new Size(100, 30);
            _addOptionButton.Click += AddOptionButton_Click;

            _removeOptionButton = new Button();
            _removeOptionButton.Text = "删除选项";
            _removeOptionButton.Location = new Point(110, 250);
            _removeOptionButton.Size = new Size(100, 30);
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
            okButton.Click += OkButton_Click;

            var cancelButton = new Button();
            cancelButton.Text = "取消";
            cancelButton.Location = new Point(295, 400);
            cancelButton.Size = new Size(75, 30);
            cancelButton.DialogResult = DialogResult.Cancel;

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
            _typeComboBox.SelectedIndex = (int)_column.Type;
            
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
            var needsOptions = _typeComboBox.SelectedIndex == 1 || _typeComboBox.SelectedIndex == 2;
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
            _column.Type = (ColumnType)_typeComboBox.SelectedIndex;
        }

        public ColumnDefinition GetColumnDefinition() => _column;
    }

    public class OptionEditDialog : Form
    {
        private TextBox _labelTextBox;
        private Button _colorButton;
        private string _selectedColor = "#FFFFFF";

        public OptionEditDialog()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "编辑选项";
            this.Size = new Size(300, 150);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

            var labelLabel = new Label();
            labelLabel.Text = "标签:";
            labelLabel.Location = new Point(12, 15);
            labelLabel.Size = new Size(60, 23);

            _labelTextBox = new TextBox();
            _labelTextBox.Location = new Point(80, 12);
            _labelTextBox.Size = new Size(180, 23);

            var colorLabel = new Label();
            colorLabel.Text = "颜色:";
            colorLabel.Location = new Point(12, 50);
            colorLabel.Size = new Size(60, 23);

            _colorButton = new Button();
            _colorButton.Location = new Point(80, 47);
            _colorButton.Size = new Size(180, 23);
            _colorButton.BackColor = Color.White;
            _colorButton.Click += ColorButton_Click;

            var okButton = new Button();
            okButton.Text = "确定";
            okButton.Location = new Point(105, 85);
            okButton.Size = new Size(75, 23);
            okButton.DialogResult = DialogResult.OK;

            var cancelButton = new Button();
            cancelButton.Text = "取消";
            cancelButton.Location = new Point(185, 85);
            cancelButton.Size = new Size(75, 23);
            cancelButton.DialogResult = DialogResult.Cancel;

            this.Controls.Add(labelLabel);
            this.Controls.Add(_labelTextBox);
            this.Controls.Add(colorLabel);
            this.Controls.Add(_colorButton);
            this.Controls.Add(okButton);
            this.Controls.Add(cancelButton);
        }

        private void ColorButton_Click(object sender, EventArgs e)
        {
            using (var dialog = new ColorDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    _colorButton.BackColor = dialog.Color;
                    _selectedColor = $"#{dialog.Color.R:X2}{dialog.Color.G:X2}{dialog.Color.B:X2}";
                }
            }
        }

        public OptionItem GetOption()
        {
            return new OptionItem
            {
                Label = _labelTextBox.Text,
                Color = _selectedColor
            };
        }
    }
}