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
    public partial class F_DeleteSame : BaseControl
    {
        private int key1 = 0;
        private int key2 = 0;
        private string excuteSingle1;
        private string excuteSingle2;
        public F_DeleteSame()
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
                DataTable table1 = c1.ExportDataTable(0, 0, c1.MaxDataRow + 1, c1.MaxDataColumn + 1, true);
                DataTable table2 = c2.ExportDataTable(0, 0, c2.MaxDataRow + 1, c2.MaxDataColumn + 1, true);
                table1.PrimaryKey = new DataColumn[] { table1.Columns[key1 - 1] };
                table2.PrimaryKey = new DataColumn[] { table2.Columns[key2 - 1] };
                DataTable MaxDatatable;
                DataTable MinDatatable;
                int minKey = 0;
                if (table1.Rows.Count > table2.Rows.Count)
                {
                    MaxDatatable = table1;
                    MinDatatable = table2;
                    minKey = key2;
                }
                else
                {
                    MaxDatatable = table2;
                    MinDatatable = table1;
                    minKey = key1;
                }

                for (int i = 0; i < MinDatatable.Rows.Count; i++)
                {
                    DataRow result = MaxDatatable.Rows.Find(MinDatatable.Rows[i][minKey - 1]);
                    if (result != null)
                    {
                        MaxDatatable.Rows.RemoveAt(i);
                        
                    }
                }
                c2.Clear();
                c2.ImportDataTable(MaxDatatable, true, 0, 0, MaxDatatable.Rows.Count, MaxDatatable.Columns.Count);
                w2.Save(txtResult.Text);
                workCheck.ReportProgress(0, "处理完成");
            }
            catch (Exception ex)
            {
                workCheck.ReportProgress(0, ex.Message);
            }

        }
        public DataTable Merge(DataTable sourceDataTable, DataTable targetDataTable)
        {
            if (sourceDataTable != null || targetDataTable != null || !sourceDataTable.Equals(targetDataTable))
            {
              //  sourceDataTable.PrimaryKey = new DataColumn[] { sourceDataTable.Columns[primaryKey] };
                DataTable dt = targetDataTable.Copy();
                foreach (DataRow tRow in dt.Rows)
                {
                    //拒绝自上次调用 System.Data.DataRow.AcceptChanges() 以来对该行进行的所有更改。
                    //因为行状态为DataRowState.Deleted时无法访问ItemArray的值
                    tRow.RejectChanges();
                    //在加载数据时关闭通知、索引维护和约束。
                    sourceDataTable.BeginLoadData();
                    //查找和更新特定行。如果找不到任何匹配行，则使用给定值创建新行。
                    DataRow temp = sourceDataTable.LoadDataRow(tRow.ItemArray, true);
                    sourceDataTable.EndLoadData();
                    sourceDataTable.Rows.Remove(temp);
                }
            }
            sourceDataTable.AcceptChanges();
            return sourceDataTable;
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

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txt1.Text) || string.IsNullOrEmpty(txt2.Text))
            {
                MessageBox.Show("请选择要处理的Excel");
                return;
            }
            int i = 0;
            int j = 0;
            if (!int.TryParse(txtKey1.Text, out i) || !int.TryParse(txtKey2.Text, out j))
            {
                MessageBox.Show("关键列格式不正确，必须为数字");
                return;
            }
            if (cbbSheet1.SelectedItem == null || cbbSheet2.SelectedItem == null)
            {
                MessageBox.Show("请选择要处理的表");
                return;
            }
            excuteSingle1 = cbbSheet1.SelectedItem.ToString();
            excuteSingle2 = cbbSheet2.SelectedItem.ToString();
            key1 = i;
            key2 = j;
            txtLog.Clear();
            btnStart.Text = "正在处理...";
            btnStart.Enabled = false;
            btnClose.Enabled = false;
            workCheck.RunWorkerAsync();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Helper.OpenHelpFile(7);
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            Helper.OpenUrl(txtResult.Text);
        }
    }
}
