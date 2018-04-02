using Aspose.Cells;
using CCWin;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Excel汇总
{
    public partial class F_Huizongduoge : BaseControl
    {
        public F_Huizongduoge()
        {
            InitializeComponent();
            workCheck.WorkerReportsProgress = true;
            workCheck.WorkerSupportsCancellation = true;
            workCheck.DoWork += new DoWorkEventHandler(workCheck_DoWork);
            workCheck.ProgressChanged += new ProgressChangedEventHandler(workCheck_ProgressChanged);

            workCheck.RunWorkerCompleted += new RunWorkerCompletedEventHandler(workCheck_RunWorkerCompleted);
            cbbExcuteTable.CheckBoxCheckedChanged += CbbExcuteTable_CheckBoxCheckedChanged;
        }

        private void CbbExcuteTable_CheckBoxCheckedChanged(object sender, EventArgs e)
        {
            radioButton1.Checked = true;
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
                string[] listAllExcel = DirFileHelper.GetFileNames(txtSource.Text, "*.xls", true);
                Workbook template = new Workbook(listAllExcel[0]);
                foreach (var item in listAllExcel)
                {
                    if (item == listAllExcel[0])
                    {
                        continue;
                    }
                    Workbook wb = new Workbook(item);
                    template.Combine(wb);

                }
                for (int i = 0; i < template.Worksheets.Count; i++)
                {
                    if (!IsExcuteSheet(template.Worksheets[i].Name))
                    {
                        template.Worksheets.RemoveAt(i);
                        i--;
                    }
                }
                template.Save(txtResult.Text);
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
        private void button1_Click(object sender, EventArgs e)
        {
            txtSource.Text = Helper.SelectPath();
            InitExcuteTable();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            txtResult.Text = Helper.SelectFile();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtSource.Text) || string.IsNullOrEmpty(cbbExcuteTable.Text))
            {
                MessageBox.Show("请选择要处理的Excel,并选择要处理的表");
                return;
            }
            txtLog.Clear();
            btnStart.Text = "正在处理...";
            btnStart.Enabled = false;
            btnClose.Enabled = false;
            listExcuteTable = new List<string>(cbbExcuteTable.Text.Split(','));
            IsMohu = rbMohu.Checked;
            workCheck.RunWorkerAsync();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Helper.OpenHelpFile(2);
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            Helper.OpenUrl(txtResult.Text);
        }

        private void btnCheckAll_Click(object sender, EventArgs e)
        {
            Helper.CheckAll(cbbExcuteTable);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Helper.UnCheckAll(cbbExcuteTable);
        }
        private void InitExcuteTable()
        {
            cbbExcuteTable.Items.Clear();
            if (string.IsNullOrEmpty(txtSource.Text))
            {
                return;
            }
            string[] listAllExcel = DirFileHelper.GetFileNames(txtSource.Text, "*.xls", true);
            Hashtable ht = new Hashtable();
            foreach (var item in listAllExcel)
            {
                Workbook wb = new Workbook(item);
                foreach (var sheet in wb.Worksheets)
                {
                    if (!ht.ContainsKey(sheet.Name))
                    {
                        ht.Add(sheet.Name, null);
                    }
                }
            }
            foreach (var key in ht.Keys)
            {
                cbbExcuteTable.Items.Add(key);
            }
        }
    }
}
