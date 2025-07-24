using System;
using System.Windows.Forms;

namespace TaskBoard
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                ApplicationConfiguration.Initialize();
                // 使用不依赖 MaterialSkin 的版本
                Application.Run(new MainFormNoMaterial());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"程序启动失败：\n{ex.Message}\n\n详细信息：\n{ex.StackTrace}", 
                    "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}