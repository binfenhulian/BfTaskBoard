using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;
using TaskBoard.Models;

namespace TaskBoard.Controls
{
    public class TodoListControl : UserControl
    {
        private FlowLayoutPanel _flowPanel;
        private List<TodoItem> _todos = new List<TodoItem>();
        private Action _onDataChanged;
        private Color _backgroundColor;

        public List<TodoItem> Todos 
        { 
            get => _todos;
            set 
            {
                _todos = value ?? new List<TodoItem>();
                RefreshDisplay();
            }
        }

        public TodoListControl(Color backgroundColor, Action onDataChanged)
        {
            _backgroundColor = backgroundColor;
            _onDataChanged = onDataChanged;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Size = new Size(300, 100);
            this.BackColor = _backgroundColor;
            this.Padding = new Padding(5);

            _flowPanel = new FlowLayoutPanel();
            _flowPanel.Dock = DockStyle.Fill;
            _flowPanel.FlowDirection = FlowDirection.TopDown;
            _flowPanel.WrapContents = false;
            _flowPanel.AutoScroll = true;
            _flowPanel.BackColor = _backgroundColor;

            var addButton = new Label();
            addButton.Text = "+ 添加待办";
            addButton.AutoSize = true;
            addButton.ForeColor = Color.FromArgb(100, 100, 100);
            addButton.Cursor = Cursors.Hand;
            addButton.Margin = new Padding(2);
            addButton.Click += AddButton_Click;

            _flowPanel.Controls.Add(addButton);
            this.Controls.Add(_flowPanel);
        }

        private void AddButton_Click(object sender, EventArgs e)
        {
            var newTodo = new TodoItem { Text = "新待办事项" };
            _todos.Add(newTodo);
            RefreshDisplay();
            _onDataChanged?.Invoke();
        }

        public void RefreshDisplay()
        {
            _flowPanel.SuspendLayout();
            _flowPanel.Controls.Clear();

            foreach (var todo in _todos)
            {
                var todoPanel = CreateTodoItemPanel(todo);
                _flowPanel.Controls.Add(todoPanel);
            }

            var addButton = new Label();
            addButton.Text = "+ 添加待办";
            addButton.AutoSize = true;
            addButton.ForeColor = Color.FromArgb(100, 100, 100);
            addButton.Cursor = Cursors.Hand;
            addButton.Margin = new Padding(2);
            addButton.Click += AddButton_Click;
            _flowPanel.Controls.Add(addButton);

            _flowPanel.ResumeLayout();
        }

        private Panel CreateTodoItemPanel(TodoItem todo)
        {
            var panel = new Panel();
            panel.Height = 25;
            panel.Width = _flowPanel.ClientSize.Width - 10;
            panel.Margin = new Padding(0, 2, 0, 2);

            var checkBox = new CheckBox();
            checkBox.Location = new Point(5, 3);
            checkBox.Size = new Size(20, 20);
            checkBox.Checked = todo.IsCompleted;
            checkBox.FlatStyle = FlatStyle.Flat;
            checkBox.FlatAppearance.BorderSize = 0;
            checkBox.Paint += (s, e) => DrawCustomCheckBox(e.Graphics, checkBox);

            var textBox = new TextBox();
            textBox.Location = new Point(30, 3);
            textBox.Width = panel.Width - 60;
            textBox.BorderStyle = BorderStyle.None;
            textBox.BackColor = _backgroundColor;
            textBox.ForeColor = todo.IsCompleted ? Color.Gray : Color.White;
            textBox.Text = todo.Text;
            textBox.Font = new Font(textBox.Font, todo.IsCompleted ? FontStyle.Strikeout : FontStyle.Regular);
            textBox.TextChanged += (s, e) =>
            {
                todo.Text = textBox.Text;
                _onDataChanged?.Invoke();
            };

            checkBox.CheckedChanged += (s, e) =>
            {
                var message = checkBox.Checked ? "确认完成此待办事项？" : "确认取消完成状态？";
                var result = MessageBox.Show(message, "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    todo.IsCompleted = checkBox.Checked;
                    textBox.Font = new Font(textBox.Font, todo.IsCompleted ? FontStyle.Strikeout : FontStyle.Regular);
                    textBox.ForeColor = todo.IsCompleted ? Color.Gray : Color.White;
                    _onDataChanged?.Invoke();
                }
                else
                {
                    checkBox.Checked = !checkBox.Checked;
                }
            };

            var deleteButton = new Label();
            deleteButton.Text = "×";
            deleteButton.Location = new Point(panel.Width - 25, 3);
            deleteButton.Size = new Size(20, 20);
            deleteButton.ForeColor = Color.FromArgb(150, 150, 150);
            deleteButton.TextAlign = ContentAlignment.MiddleCenter;
            deleteButton.Cursor = Cursors.Hand;
            deleteButton.Click += (s, e) =>
            {
                _todos.Remove(todo);
                RefreshDisplay();
                _onDataChanged?.Invoke();
            };

            panel.Controls.Add(checkBox);
            panel.Controls.Add(textBox);
            panel.Controls.Add(deleteButton);

            return panel;
        }

        private void DrawCustomCheckBox(Graphics g, CheckBox checkBox)
        {
            g.Clear(_backgroundColor);
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var rect = new Rectangle(2, 2, 16, 16);
            var dotRect = new Rectangle(6, 6, 8, 8);

            using (var pen = new Pen(Color.FromArgb(100, 100, 100), 1))
            {
                g.DrawEllipse(pen, rect);
            }

            if (checkBox.Checked)
            {
                using (var brush = new SolidBrush(Color.FromArgb(0, 200, 0)))
                {
                    g.FillEllipse(brush, dotRect);
                }
            }
            else
            {
                using (var brush = new SolidBrush(Color.FromArgb(200, 200, 200)))
                {
                    g.FillEllipse(brush, dotRect);
                }
            }
        }
    }
}