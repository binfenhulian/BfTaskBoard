using System;
using System.Drawing;
using System.Windows.Forms;
using MaterialSkin;
using MaterialSkin.Controls;

namespace TaskBoard.Forms
{
    public class InputDialog : MaterialForm
    {
        private MaterialTextBox2 _textBox;
        private MaterialButton _okButton;
        private MaterialButton _cancelButton;

        public string InputValue => _textBox.Text;

        public InputDialog(string title, string prompt, string defaultValue = "")
        {
            InitializeComponent(title, prompt, defaultValue);
            
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.DARK;
            materialSkinManager.ColorScheme = new ColorScheme(
                Primary.Blue800, Primary.Blue900,
                Primary.Blue500, Accent.LightBlue200,
                TextShade.WHITE);
        }

        private void InitializeComponent(string title, string prompt, string defaultValue)
        {
            this.Text = title;
            this.Size = new Size(400, 200);
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Sizable = false;

            var promptLabel = new MaterialLabel();
            promptLabel.Text = prompt;
            promptLabel.Location = new Point(20, 80);
            promptLabel.AutoSize = true;

            _textBox = new MaterialTextBox2();
            _textBox.Text = defaultValue;
            _textBox.Location = new Point(20, 110);
            _textBox.Width = 360;
            _textBox.Hint = prompt;

            _okButton = new MaterialButton();
            _okButton.Text = "确定";
            _okButton.Location = new Point(225, 160);
            _okButton.AutoSize = false;
            _okButton.Size = new Size(75, 36);
            _okButton.Click += (s, e) => { DialogResult = DialogResult.OK; Close(); };

            _cancelButton = new MaterialButton();
            _cancelButton.Text = "取消";
            _cancelButton.Location = new Point(305, 160);
            _cancelButton.AutoSize = false;
            _cancelButton.Size = new Size(75, 36);
            _cancelButton.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

            this.Controls.Add(promptLabel);
            this.Controls.Add(_textBox);
            this.Controls.Add(_okButton);
            this.Controls.Add(_cancelButton);

            _textBox.Focus();
            _textBox.SelectAll();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            _textBox.Focus();
            _textBox.SelectAll();
        }
    }
}