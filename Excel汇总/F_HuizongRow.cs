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
    public partial class F_HuizongRow : BaseControl
    {
        private int StartRow = 0;
        private int EndRow = 0;
        private int pdKeyColumn = 0;
        public F_HuizongRow()
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
                Workbook template = new Workbook();
                Cells TemplateCells = template.Worksheets[0].Cells;
                int k = 0;
                // DataTable TemplateTable = TemplateCells.ExportDataTable(StartRow - 1, 0, TemplateCells.MaxDataRow - EndRow + 1, TemplateCells.MaxDataColumn + 1);
                foreach (var item in listAllExcel)
                {
                    Workbook workbook = new Workbook(item);
                    ExportTableOptions opts = new ExportTableOptions();
                    opts.ExportAsString = true;
                    foreach (var sheet in workbook.Worksheets)
                    {

                        if (!IsExcuteSheet(sheet.Name))
                        {
                            workCheck.ReportProgress(0, "此sheet名称不在待处理表中，将忽略...---" + sheet.Name);
                            continue;
                        }
                        workCheck.ReportProgress(0, "开始处理:" + item + "/" + sheet.Name);
                        Cells cells = sheet.Cells;
                        DataTable table;
                        k++;
                        if (k == listAllExcel.Length)
                        {
                            //最后一个表保留表头
                            table = cells.ExportDataTable(0, 0, cells.MaxDataRow + 1 - EndRow + 1, cells.MaxDataColumn + 1, opts);
                        }
                        else
                        {
                            table = cells.ExportDataTable(StartRow - 1, 0, cells.MaxDataRow - EndRow + 1, cells.MaxDataColumn + 1, opts);
                        }


                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            //去除无效数据行
                            if (table.Rows[i][pdKeyColumn - 1] == null || string.IsNullOrEmpty(table.Rows[i][pdKeyColumn - 1].ToString()))
                            {
                                table.Rows.RemoveAt(i);
                                //调整索引减1  
                                i--;
                                continue;
                            }
                        }
                        //TemplateCells.ImportDataTable(table, false, 0, 0, table.Rows.Count, table.Columns.Count, true, "yyyy/MM/dd");
                        TemplateCells.ImportDataTable(table, false, 0, 0, table.Rows.Count, table.Columns.Count, true, "yyyy/MM/dd");
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
            int i = 0;
            int j = 0;
            int l = 0;
            if (!int.TryParse(txtStart.Text, out i) || !int.TryParse(txtEnd.Text, out j) || !int.TryParse(txtPDKeyColumn.Text, out l))
            {
                MessageBox.Show("要处理的行格式不正确，必须为数字");
                return;
            }
            StartRow = i;
            EndRow = j;
            pdKeyColumn = l;
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
            Helper.OpenHelpFile(3);
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
    }
}
