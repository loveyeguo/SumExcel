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
    public partial class F_TiquFormatSame : BaseControl
    {
        public F_TiquFormatSame()
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
                bool isUseTemplate = false;
                workCheck.ReportProgress(0, "开始处理...");
                List<ModelStore> list = new List<ModelStore>();
                string[] listAllExcel = DirFileHelper.GetFileNames(txtSource.Text, "*.xls", true);
                Workbook workbookTemplate;
                workbookTemplate = new Workbook(listAllExcel[0]);
                Cells cellsTemplate = workbookTemplate.Worksheets[0].Cells;
                string str = string.Empty;
                //  workCheck.ReportProgress(0, "模版处理...");
                for (int i = 0; i < cellsTemplate.MaxDataRow + 1; i++)
                {
                    for (int j = 0; j < cellsTemplate.MaxDataColumn + 1; j++)
                    {
                        if (cellsTemplate[i, j].IsFormula)
                        {
                            continue;
                        }
                        if (isUseTemplate)
                        {
                            if (cellsTemplate[i, j].Value != null && cellsTemplate[i, j].StringValue.ToLower() == TemplateTag)
                            {
                                list.Add(new ModelStore { row = i, column = j });
                            }
                        }
                        else
                        {

                            list.Add(new ModelStore { row = i, column = j });

                        }

                    }
                }
                if (e.Argument.ToString() == "按行排列")
                {
                    SumByRow(listAllExcel, list);
                }
                else if (e.Argument.ToString() == "按列排列")
                {
                    SumByColumn(listAllExcel, list);
                }
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
            workCheck.RunWorkerAsync(cbbSort.SelectedItem.ToString());
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void SumByRow(string[] listExcel, List<ModelStore> listStore)
        {
            Workbook wb = new Workbook(FileFormatType.Excel97To2003);
            Worksheet sheet = wb.Worksheets[0];
            sheet.Name = "sheet";
            Cells cellsResult = wb.Worksheets[0].Cells;
           
            int row = 0;
            foreach (var item in listExcel)
            {
                Workbook workbook = new Workbook(item);
                foreach (var s in workbook.Worksheets)
                {

                    if (!IsExcuteSheet(s.Name))
                    {
                        workCheck.ReportProgress(0, "此sheet名称不在待处理表中，将忽略...---" + s.Name);
                        continue;
                    }
                    Cells cells = s.Cells;
                    int column = 0;
                    cellsResult[row, column].PutValue(item);
                    column++;
                    for (int i = 0; i < cells.MaxDataRow + 1; i++)
                    {
                        for (int j = 0; j < cells.MaxDataColumn + 1; j++)
                        {
                            if (listStore.FindIndex(x => x.row == i && x.column == j) >= 0)
                            {
                                cellsResult[row, column].PutValue(cells[i, j].Value);
                                column++;
                            }

                        }
                    }
                    row++;
                }
            }
            wb.Worksheets[0].AutoFitColumns();
            wb.Save(txtResult.Text);//默认支持xls版，需要修改指定版本

        }

        private void SumByColumn(string[] listExcel, List<ModelStore> listStore)
        {
            Workbook wb = new Workbook(FileFormatType.Excel97To2003);
            Worksheet sheet = wb.Worksheets[0];
            sheet.Name = "sheet";
            Cells cellsResult = wb.Worksheets[0].Cells;
            int column = 0;
           
            foreach (var item in listExcel)
            {
                Workbook workbook = new Workbook(item);
                Cells cells = workbook.Worksheets[0].Cells;
                int row = 0;
                cellsResult[row, column].PutValue(item);
                row++;
                for (int i = 0; i < cells.MaxDataRow + 1; i++)
                {
                    for (int j = 0; j < cells.MaxDataColumn + 1; j++)
                    {
                        if (listStore.FindIndex(x => x.row == i && x.column == j) >= 0)
                        {
                            cellsResult[row, column].PutValue(cells[i, j].Value);
                            row++;
                        }

                    }
                }
                column++;
            }
            wb.Worksheets[0].AutoFitColumns();
            wb.Save(txtResult.Text);//默认支持xls版，需要修改指定版本
        }

        private void F_TiquFormatSame_Load(object sender, EventArgs e)
        {
            cbbSort.SelectedIndex = 0;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Helper.OpenHelpFile(9);
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
