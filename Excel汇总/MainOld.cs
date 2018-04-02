using CCWin;
using CCWin.SkinControl;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Excel汇总
{
    public partial class MainOld : CCSkinMain
    {
        public MainOld()
        {
            InitializeComponent();
        }
        private void AddWindows(UserControl f, SkinButton button)
        {
            if (button != null)
            {
                button.BaseColor = Color.GreenYellow;
            }
            HideAllControlInPanel();
            Control c = IsExitWindowsInPanel(f);
            if (c == null)
            {
                panelMain.Controls.Add(f);
                f.Dock = DockStyle.Fill;
            }
            else
            {
                c.Show();
            }


        }
        private void HideAllControlInPanel()
        {
            foreach (Control item in panelMain.Controls)
            {

                item.Hide();


            }
        }
        private Control IsExitWindowsInPanel(UserControl f)
        {
            Control[] arr = panelMain.Controls.Find(f.Name, true);
            if (arr.Length == 0)
            {
                return null;
            }
            return arr[0];
        }
        /// <summary>
        /// 汇总相同格式的工作表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void skinButton1_Click(object sender, EventArgs e)
        {
            SkinButton button = sender as SkinButton;
          
        }
    }
}
