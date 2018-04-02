using Aspose.Cells;
using PresentationControls;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Excel汇总
{
    public class Helper
    {
        public static string SelectPath()
        {
            string path = string.Empty;
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                path = fbd.SelectedPath;
            }
            return path;
        }
        public static string SelectFile()
        {
            string path = string.Empty;
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Files (*.*)|*.*"//如果需要筛选txt文件（"Files (*.txt)|*.txt"）
            };
            var result = openFileDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                path = openFileDialog.FileName;
            }
            return path;
        }
        public static void CheckAll(PresentationControls.CheckBoxComboBox cbb)
        {
            foreach (CheckBoxComboBoxItem item in cbb.CheckBoxItems)
            {
                item.Checked = true;
            }

        }
        public static void UnCheckAll(PresentationControls.CheckBoxComboBox cbb)
        {
            foreach (CheckBoxComboBoxItem item in cbb.CheckBoxItems)
            {
                item.Checked = false;
            }
        }
        public static void OpenUrl(string url)
        {
            url = System.IO.Path.GetDirectoryName(url);
            Process.Start("explorer.exe", url);
        }
        public static void OpenHelpFile(int i)
        {
            string url = Application.StartupPath + "\\help\\Help-00" + i + ".htm";
            System.Diagnostics.Process.Start(url);
        }
    }
}
