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
    public partial class F_SumSameSheet : BaseControl
    {

        public F_SumSameSheet()
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
        private void InitCbbTemplateSheet()
        {
            if (string.IsNullOrEmpty(txtTemplate.Text))
            {
                return;
            }

            Workbook wb = new Workbook(txtTemplate.Text);
            foreach (var sheet in wb.Worksheets)
            {
                CbbTemplateSheet.Items.Add(sheet.Name);
            }

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
                Cells cellsTemplate;
                if (!string.IsNullOrEmpty(txtTemplate.Text))
                {
                    workbookTemplate = new Workbook(txtTemplate.Text);
                    isUseTemplate = true;
                    cellsTemplate = workbookTemplate.Worksheets[TemplateSheetName].Cells;
                }
                else
                {
                    workbookTemplate = new Workbook(listAllExcel[0]);
                    cellsTemplate = workbookTemplate.Worksheets[0].Cells;
                }
                
                string str = string.Empty;
                workCheck.ReportProgress(0, "模版处理...");
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
                            if (cellsTemplate[i, j].Type == CellValueType.IsNumeric)
                            {
                                list.Add(new ModelStore { row = i, column = j });
                            }
                        }

                    }
                }
                foreach (var item in listAllExcel)
                {
                    if (!isUseTemplate && item == listAllExcel[0])
                    {
                        continue;
                    }
                    Workbook workbook = new Workbook(item);
                    foreach (var sheet in workbook.Worksheets)
                    {

                        if (!IsExcuteSheet(sheet.Name))
                        {
                            workCheck.ReportProgress(0, "此sheet名称不在待处理表中，将忽略...---" + sheet.Name);
                            continue;
                        }

                        Cells cells = sheet.Cells;
                        workCheck.ReportProgress(0, "开始处理数据...---" + workbook.FileName + "sheetName:" + sheet.Name);
                        for (int i = 0; i < cells.MaxDataRow + 1; i++)
                        {
                            for (int j = 0; j < cells.MaxDataColumn + 1; j++)
                            {
                                if (list.FindIndex(x => x.row == i && x.column == j) >= 0 && cells[i, j].Type == CellValueType.IsNumeric)
                                {
                                    if (isUseTemplate && cellsTemplate[i, j].Type != CellValueType.IsNumeric)
                                    {
                                        cellsTemplate[i, j].PutValue(cells[i, j].DoubleValue);
                                    }
                                    else
                                    {
                                        cellsTemplate[i, j].PutValue(cellsTemplate[i, j].DoubleValue + cells[i, j].DoubleValue);
                                    }

                                }

                            }
                        }
                    }
                }


                if (isUseTemplate)
                {
                    workCheck.ReportProgress(0, "开始将含有(#TAG#)的单元格变为0");
                    for (int i = 0; i < cellsTemplate.MaxDataRow + 1; i++)
                    {
                        for (int j = 0; j < cellsTemplate.MaxDataColumn + 1; j++)
                        {
                            if (cellsTemplate[i, j].StringValue.ToLower() == TemplateTag)
                            {
                                cellsTemplate[i, j].PutValue(0);
                            }
                        }
                    }
                }

                workbookTemplate.Save(txtResult.Text);
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
        private void button4_Click(object sender, EventArgs e)
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
            if (!string.IsNullOrEmpty(txtTemplate.Text) && CbbTemplateSheet.SelectedItem == null)
            {
                MessageBox.Show("请选择模版的sheet");
                return;
            }
            else if (CbbTemplateSheet.SelectedItem!=null)
            {
                TemplateSheetName = CbbTemplateSheet.SelectedItem.ToString();
            }
          
            txtLog.Clear();
            btnStart.Text = "正在处理...";
            btnStart.Enabled = false;
            btnClose.Enabled = false;
            listExcuteTable = new List<string>(cbbExcuteTable.Text.Split(','));
            IsMohu = rbMohu.Checked;
            workCheck.RunWorkerAsync();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            txtSource.Text = Helper.SelectPath();
            InitExcuteTable();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            txtTemplate.Text = Helper.SelectFile();
            InitCbbTemplateSheet();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            txtResult.Text = Helper.SelectFile();
        }

        private void F_SumSameSheet_Load(object sender, EventArgs e)
        {
            InitExcuteTable();
        }

        private void btnCheckAll_Click(object sender, EventArgs e)
        {
            Helper.CheckAll(cbbExcuteTable);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Helper.UnCheckAll(cbbExcuteTable);
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            Helper.OpenUrl(txtResult.Text);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Helper.OpenHelpFile(1);
        }
    }
}
