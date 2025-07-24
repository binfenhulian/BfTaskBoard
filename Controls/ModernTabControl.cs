using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace TaskBoard.Controls
{
    public class ModernTabControl : TabControl
    {
        private const int TAB_HEIGHT = 40;
        private const int ADD_BUTTON_WIDTH = 40;
        private Rectangle _addButtonRect;
        private bool _addButtonHovered = false;

        public event EventHandler AddTabClicked;
        public event EventHandler<TabRenameEventArgs> TabRenameRequested;

        public ModernTabControl()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | 
                    ControlStyles.ResizeRedraw | ControlStyles.OptimizedDoubleBuffer, true);
            
            ItemSize = new Size(0, TAB_HEIGHT);
            DrawMode = TabDrawMode.OwnerDrawFixed;
            SizeMode = TabSizeMode.Fixed;
            
            MouseMove += OnMouseMove;
            MouseLeave += OnMouseLeave;
            MouseClick += OnMouseClick;
            MouseDoubleClick += OnMouseDoubleClick;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.Clear(Color.FromArgb(30, 30, 30));

            var tabsWidth = 0;
            
            for (int i = 0; i < TabCount; i++)
            {
                var tabRect = GetTabRect(i);
                var isSelected = SelectedIndex == i;
                
                using (var brush = new SolidBrush(isSelected ? Color.FromArgb(45, 45, 45) : Color.FromArgb(35, 35, 35)))
                {
                    g.FillRectangle(brush, tabRect);
                }

                if (isSelected)
                {
                    using (var pen = new Pen(Color.FromArgb(0, 122, 204), 2))
                    {
                        g.DrawLine(pen, tabRect.Left, tabRect.Bottom - 1, tabRect.Right, tabRect.Bottom - 1);
                    }
                }

                using (var brush = new SolidBrush(Color.White))
                {
                    var textRect = new Rectangle(tabRect.X + 10, tabRect.Y, tabRect.Width - 20, tabRect.Height);
                    var format = new StringFormat
                    {
                        Alignment = StringAlignment.Near,
                        LineAlignment = StringAlignment.Center,
                        Trimming = StringTrimming.EllipsisCharacter
                    };
                    
                    g.DrawString(TabPages[i].Text, Font, brush, textRect, format);
                }

                tabsWidth = tabRect.Right;
            }

            _addButtonRect = new Rectangle(tabsWidth + 5, 5, ADD_BUTTON_WIDTH, TAB_HEIGHT - 10);
            
            using (var brush = new SolidBrush(_addButtonHovered ? Color.FromArgb(60, 60, 60) : Color.FromArgb(45, 45, 45)))
            {
                g.FillRectangle(brush, _addButtonRect);
            }

            using (var pen = new Pen(Color.White, 2))
            {
                var centerX = _addButtonRect.X + _addButtonRect.Width / 2;
                var centerY = _addButtonRect.Y + _addButtonRect.Height / 2;
                g.DrawLine(pen, centerX - 8, centerY, centerX + 8, centerY);
                g.DrawLine(pen, centerX, centerY - 8, centerX, centerY + 8);
            }

            var contentRect = new Rectangle(0, TAB_HEIGHT, Width, Height - TAB_HEIGHT);
            using (var brush = new SolidBrush(Color.FromArgb(37, 37, 38)))
            {
                g.FillRectangle(brush, contentRect);
            }
        }

        protected override void OnPaintBackground(PaintEventArgs pevent)
        {
        }

        private void OnMouseMove(object sender, MouseEventArgs e)
        {
            var wasHovered = _addButtonHovered;
            _addButtonHovered = _addButtonRect.Contains(e.Location);
            
            if (wasHovered != _addButtonHovered)
            {
                Invalidate();
            }
        }

        private void OnMouseLeave(object sender, EventArgs e)
        {
            if (_addButtonHovered)
            {
                _addButtonHovered = false;
                Invalidate();
            }
        }

        private void OnMouseClick(object sender, MouseEventArgs e)
        {
            if (_addButtonRect.Contains(e.Location))
            {
                AddTabClicked?.Invoke(this, EventArgs.Empty);
            }
        }

        private void OnMouseDoubleClick(object sender, MouseEventArgs e)
        {
            for (int i = 0; i < TabCount; i++)
            {
                if (GetTabRect(i).Contains(e.Location))
                {
                    TabRenameRequested?.Invoke(this, new TabRenameEventArgs { TabIndex = i, CurrentName = TabPages[i].Text });
                    break;
                }
            }
        }
    }

    public class TabRenameEventArgs : EventArgs
    {
        public int TabIndex { get; set; }
        public string CurrentName { get; set; }
    }
}