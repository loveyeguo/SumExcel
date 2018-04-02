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
    public partial class Main : CCSkinMain
    {
        public Main()
        {
            InitializeComponent();
        }
        private void OpenWindow(Form form)
        {
            form.StartPosition = FormStartPosition.CenterScreen;
            form.ShowDialog();
        }

        private void Form_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Show();
        }

        /// <summary>
        /// 汇总格式相同的工作表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnxueshuSearch_Click(object sender, EventArgs e)
        {
            OpenWindow(new F_SumSameSheet());
        }

        private void skinButton1_Click(object sender, EventArgs e)
        {
            OpenWindow(new F_TiquFormatSame());
        }

        private void skinButton2_Click(object sender, EventArgs e)
        {
            OpenWindow(new F_Huizongduoge());
        }

        private void skinButton3_Click(object sender, EventArgs e)
        {
            OpenWindow(new F_HuizongRow());
        }

        private void skinButton4_Click(object sender, EventArgs e)
        {
            OpenWindow(new F_Fenleihuizong());
        }

        private void skinButton5_Click(object sender, EventArgs e)
        {
            OpenWindow(new F_Hebinglie());
        }

        private void skinButton6_Click(object sender, EventArgs e)
        {
            OpenWindow(new F_TiquGongyou());
        }

        private void skinButton7_Click(object sender, EventArgs e)
        {
            OpenWindow(new F_DeleteSame());
        }

        private void skinButton8_Click(object sender, EventArgs e)
        {
            OpenWindow(new F_Check());
        }

        private void Main_Load(object sender, EventArgs e)
        {
            txtInfo.Hide();
            //InitToolMsg();
            ButtonMouseLeaveEvent();
        }
        void toolTip1_Draw(object sender, DrawToolTipEventArgs e)
        {
            Font f = new Font("Arial", 16.0f);
            e.DrawBackground();
            e.DrawBorder();

            e.Graphics.DrawString(e.ToolTipText, f, Brushes.Black, new PointF(2, 2));
        }
        //private void InitToolMsg()
        //{
        //    toolMsg.SetToolTip(btnxueshuSearch, "将多个行、列格式相同的工作表的内容汇总到单个工作表对应单元格中。参加汇总的工作表可以在一个Excel文件中，也可以在不同Excel文件中");
        //    toolMsg.SetToolTip(skinButton1, "提取多个格式相同工作的表数据，可以提取一个Excel文件的多个工作表，也可以提取多个Excel文件的多个工作表的数据");
        //    toolMsg.SetToolTip(skinButton2, "");

        //    toolMsg.SetToolTip(skinButton3, "");
        //    toolMsg.SetToolTip(skinButton4, "");
        //    toolMsg.SetToolTip(skinButton5, "");
        //    toolMsg.SetToolTip(skinButton6, "");
        //    toolMsg.SetToolTip(skinButton7, "");
        //    toolMsg.SetToolTip(skinButton8, "");
        //}

        private void btnxueshuSearch_MouseHover(object sender, EventArgs e)
        {
            txtInfo.Show();
            txtInfo.Text = "将多个行、列格式相同的工作表的内容汇总到单个工作表对应单元格中。参加汇总的工作表可以在一个Excel文件中，也可以在不同Excel文件中";
        }

        private void ButtonMouseLeaveEvent()
        {
            foreach (Control item in this.Controls)
            {
                if (item is Button && item.Tag != null)
                {
                    item.MouseLeave += Item_MouseLeave;
                }
            }
        }

        private void Item_MouseLeave(object sender, EventArgs e)
        {
            txtInfo.Hide();
            txtInfo.Text = "";
        }

        private void skinButton1_MouseHover(object sender, EventArgs e)
        {
            txtInfo.Show();
            txtInfo.Text = "提取多个格式相同工作的表数据，可以提取一个Excel文件的多个工作表，也可以提取多个Excel文件的多个工作表的数据";
        }

        private void skinButton2_MouseHover(object sender, EventArgs e)
        {
            txtInfo.Show();
            txtInfo.Text = "本项功能可以将多个Excel文件的工作表快速拷贝到一个Excel文件中";
        }

        private void skinButton3_MouseHover(object sender, EventArgs e)
        {
            txtInfo.Show();
            txtInfo.Text = "本项功能可以把多个Excel文件内工作表的行数据复制到指定的单一工作表中";
        }

        private void skinButton4_MouseHover(object sender, EventArgs e)
        {
            txtInfo.Show();
            txtInfo.Text = "本项功能和“汇总工作表行数据”类似的是，同样可以把多个Excel文件内工作表的行数据复制到指定的单一工作表中。所不同的是，可以指定一个关键列，所有这一列的值相同的行，都会被汇总成一行";
        }

        private void skinButton5_MouseHover(object sender, EventArgs e)
        {
            txtInfo.Show();
            txtInfo.Text = "本项功能可以根据关键列合并两个工作表的列，并且不要求两个工作表的行按照顺序一一对应，程序可以根据关键列的值自动匹配对应行，然后把两个工作表中的行拼接成一行";
        }

        private void skinButton6_MouseHover(object sender, EventArgs e)
        {
            txtInfo.Show();
            txtInfo.Text = "本项功能可以将2个Excel工作表中具有相同关键列值的数据行输出到目的Excel文件中。这两个工作表可以在同一个Excel文件中，也可以在不同Excel文件中";
        }

        private void skinButton7_MouseHover(object sender, EventArgs e)
        {
            txtInfo.Show();
            txtInfo.Text = "本项功能将工作表1除去与工作表2具有相同关键列值的数据行";
        }

        private void skinButton8_MouseHover(object sender, EventArgs e)
        {
            txtInfo.Show();
            txtInfo.Text = "根据模版工作表审核其他工作表。这些工作表可以在同一个Excel文件中，也可以在不同Excel文件中。审核内容： 1）单元格的数据是否满足公式。 2）单元格的数据类型是否和模版工作表相同。";
        }
    }
}
