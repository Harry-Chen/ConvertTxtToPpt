using System;
using System.Windows.Forms;

namespace ConvertTxtToPpt
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "文本文件|*.txt";
            dialog.Title = "请选择要转换的文本";
            DialogResult result = DialogResult.No;
            while (result != DialogResult.OK)
            {
                result = dialog.ShowDialog();
            }
            string txt = dialog.FileName;

            SaveFileDialog dialog2 = new SaveFileDialog();
            dialog2.Filter = "幻灯片|*.ppt";
            dialog2.Title = "请选择保存的文件名";
            result = DialogResult.No;
            while (result != DialogResult.OK)
            {
                result = dialog2.ShowDialog();
            }
            string ppt = dialog2.FileName;

            dialog.Filter = "图片文件|*.jpg;*.png;*.bmp";
            dialog.Title = "请选择红色对勾文件";
            result = DialogResult.No;
            while (result != DialogResult.OK)
            {
                result = dialog.ShowDialog();
            }
            string tick = dialog.FileName;

            Converter.createPpt(txt, ppt, tick);
        }
    }
}
