using System;
using System.Drawing;
using System.Windows.Forms;
using TaskBoard.Services;

namespace TaskBoard.Controls
{
    public class ModernDataGridView : DataGridView
    {
        public ModernDataGridView()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
            
            ApplyTheme();
            
            // 订阅主题变化事件
            ThemeService.Instance.ThemeChanged += (s, e) => ApplyTheme();
            
            EnableHeadersVisualStyles = false;
            ColumnHeadersHeight = 35;
            RowTemplate.Height = 30;
            
            BorderStyle = BorderStyle.None;
            CellBorderStyle = DataGridViewCellBorderStyle.Single;
            ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            
            AllowUserToResizeRows = false;
            AllowUserToAddRows = false;
            SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            MultiSelect = false;
            
            Font = new Font("Segoe UI", 9F);
            ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
        }
        
        private void ApplyTheme()
        {
            var colors = ThemeService.Instance.GetColors();
            
            BackgroundColor = colors.GridBackground;
            GridColor = colors.GridLines;
            DefaultCellStyle.BackColor = colors.GridBackground;
            DefaultCellStyle.ForeColor = colors.Text;
            DefaultCellStyle.SelectionBackColor = colors.Selection;
            DefaultCellStyle.SelectionForeColor = colors.SelectionText;
            
            ColumnHeadersDefaultCellStyle.BackColor = colors.GridHeader;
            ColumnHeadersDefaultCellStyle.ForeColor = colors.Text;
            ColumnHeadersDefaultCellStyle.SelectionBackColor = colors.GridHeader;
            ColumnHeadersDefaultCellStyle.SelectionForeColor = colors.Text;
            
            RowHeadersDefaultCellStyle.BackColor = colors.GridHeader;
            RowHeadersDefaultCellStyle.ForeColor = colors.Text;
            RowHeadersDefaultCellStyle.SelectionBackColor = colors.Selection;
            RowHeadersDefaultCellStyle.SelectionForeColor = colors.SelectionText;
        }

        protected override void OnCellPainting(DataGridViewCellPaintingEventArgs e)
        {
            base.OnCellPainting(e);
            
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                if (e.RowIndex % 2 == 0)
                {
                    var colors = ThemeService.Instance.GetColors();
                    e.CellStyle.BackColor = colors.GridAlternateRow;
                }
            }
        }
    }
}