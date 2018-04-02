using Aspose.Cells;
using CCWin;
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
    public partial class F_Hebinglie : BaseControl
    {
        private int key1 = 0;
        private int key2 = 0;
        private string excuteSingle1;
        private string excuteSingle2;
        public F_Hebinglie()
        {
            InitializeComponent();
            workCheck.WorkerReportsProgress = true;
            workCheck.WorkerSupportsCancellation = true;
            workCheck.DoWork += new DoWorkEventHandler(workCheck_DoWork);
            workCheck.ProgressChanged += new ProgressChangedEventHandler(workCheck_ProgressChanged);

            workCheck.RunWorkerCompleted += new RunWorkerCompletedEventHandler(workCheck_RunWorkerCompleted);
        }
        void workCheck_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            txtLog.AppendText(e.UserState.ToString() + Environment.NewLine);
        }
        void workCheck_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                workCheck.ReportProgress(0, "开始处理...");
                Workbook w1 = new Workbook(txt1.Text);
                Workbook w2 = new Workbook(txt2.Text);
                Cells c1 = w1.Worksheets[excuteSingle1].Cells;
                Cells c2 = w2.Worksheets[excuteSingle2].Cells;
                DataTable table1 = c1.ExportDataTable(0, 0, c1.MaxDataRow + 1, c1.MaxDataColumn + 1,true);
                DataTable table2 = c2.ExportDataTable(0, 0, c2.MaxDataRow + 1, c2.MaxDataColumn + 1,true);
                table1.PrimaryKey = new DataColumn[] { table1.Columns[key1 - 1] };
                table2.PrimaryKey = new DataColumn[] { table2.Columns[key2 - 1] };
                table1.Merge(table2, false, MissingSchemaAction.AddWithKey);
                c1.Clear();
                c1.ImportDataTable(table1, true, 0, 0, table1.Rows.Count, table1.Columns.Count);
                w1.Save(txtResult.Text);
                workCheck.ReportProgress(0, "处理完成");
            }
            catch (Exception ex)
            {
                workCheck.ReportProgress(0, ex.Message);
            }

        }
      
        void workCheck_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnStart.Text = "开始";
            btnStart.Enabled = true;
            btnClose.Enabled = true;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            txt1.Text = Helper.SelectFile();
            InitSingleCheckCbb(txt1.Text, cbbSheet1);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            txt2.Text = Helper.SelectFile();
            InitSingleCheckCbb(txt2.Text, cbbSheet2);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            txtResult.Text = Helper.SelectFile();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txt1.Text) || string.IsNullOrEmpty(txt2.Text))
            {
                MessageBox.Show("请选择要处理的Excel");
                return;
            }
            if (cbbSheet1.SelectedItem == null || cbbSheet2.SelectedItem == null)
            {
                MessageBox.Show("请选择要处理的表");
                return;
            }
            int i = 0;
            int j = 0;
            if (!int.TryParse(txtKey1.Text, out i) || !int.TryParse(txtKey2.Text, out j))
            {
                MessageBox.Show("关键列格式不正确，必须为数字");
                return;
            }
            key1 = i;
            key2 = j;
            excuteSingle1 = cbbSheet1.SelectedItem.ToString();
            excuteSingle2 = cbbSheet2.SelectedItem.ToString();
            txtLog.Clear();
            btnStart.Text = "正在处理...";
            btnStart.Enabled = false;
            btnClose.Enabled = false;
            workCheck.RunWorkerAsync();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Helper.OpenHelpFile(5);
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            Helper.OpenUrl(txtResult.Text);
        }
    }
}
