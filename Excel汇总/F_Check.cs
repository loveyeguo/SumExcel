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
    public partial class F_Check : BaseControl
    {
        DataTable tableResult = new DataTable();
        public F_Check()
        {
            InitializeComponent();

            tableResult.Columns.Add("位置");
            tableResult.Columns.Add("错误信息");
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

        private void AddError(string location, string info)
        {
            DataRow row = tableResult.NewRow();
            row["位置"] = location;
            row["错误信息"] = info;
            tableResult.Rows.Add(row);
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
                List<ModelStore> list = new List<ModelStore>();
                string[] listAllExcel = DirFileHelper.GetFileNames(txtSource.Text, "*.xls", true);
                Workbook workbookTemplate;
                workbookTemplate = new Workbook(txtTemplate.Text);
                Cells cellsTemplate = workbookTemplate.Worksheets[TemplateSheetName].Cells;
                string str = string.Empty;
                workCheck.ReportProgress(0, "模版处理...");
                for (int i = 0; i < cellsTemplate.MaxDataRow + 1; i++)
                {
                    for (int j = 0; j < cellsTemplate.MaxDataColumn + 1; j++)
                    {
                        if (cellsTemplate[i, j] == null || cellsTemplate[i, j].Value == null)
                        {
                            continue;
                        }
                        if (cellsTemplate[i, j].IsFormula)
                        {
                            list.Add(new ModelStore { row = i, column = j, IsFormula = true, cellType = cellsTemplate[i, j].Type, Formula = cellsTemplate[i, j].Formula });
                        }
                        else
                        {
                            list.Add(new ModelStore { row = i, column = j, IsFormula = false, cellType = cellsTemplate[i, j].Type, Formula = cellsTemplate[i, j].Formula });
                        }

                    }
                }
                foreach (var item in listAllExcel)
                {
                    Workbook workbookOld = new Workbook(item);
                    Workbook workbook = new Workbook(item);
                    foreach (var sheet in workbook.Worksheets)
                    {
                        bool isPassGongshi = true;
                        bool isPassGeshi = true;
                        if (!IsExcuteSheet(sheet.Name))
                        {
                            workCheck.ReportProgress(0, "此sheet名称不在待处理表中，将忽略...---" + sheet.Name);
                            continue;
                        }

                        Cells cells = sheet.Cells;
                        Cells cellsOld = workbookOld.Worksheets[sheet.Name].Cells;
                        workCheck.ReportProgress(0, "开始处理数据...---" + workbook.FileName + "sheetName:" + sheet.Name);
                        for (int i = 0; i < cells.MaxDataRow + 1; i++)
                        {
                            for (int j = 0; j < cells.MaxDataColumn + 1; j++)
                            {
                                ModelStore ms = list.Find(x => x.row == i && x.column == j);
                                if (ms!=null && ms.IsFormula)
                                {
                                    cells[i, j].Formula = ms.Formula;
                                }

                            }
                        }
                        workbook.CalculateFormula();
                        for (int i = 0; i < cells.MaxDataRow + 1; i++)
                        {
                            for (int j = 0; j < cells.MaxDataColumn + 1; j++)
                            {
                                ModelStore ms = list.Find(x => x.row == i && x.column == j);
                                if (ms == null)
                                {
                                    continue;
                                }
                                if (ms!=null && ms.IsFormula)
                                {
                                    if (cells[i, j].StringValue != cellsOld[i, j].StringValue)
                                    {
                                        isPassGongshi = false;
                                        AddError(workbook.FileName + ",sheetName:" + sheet.Name, CellsHelper.CellIndexToName(i, j) + "公式有误");
                                    }
                                }
                                else
                                {
                                    
                                    if (cells[i, j].Type != ms.cellType)
                                    {
                                        isPassGeshi = false;
                                        AddError(workbook.FileName + ",sheetName:" + sheet.Name, CellsHelper.CellIndexToName(i, j) + "格式有误");
                                    }
                                }

                            }
                        }
                        if (isPassGeshi)
                        {
                            AddError(workbook.FileName + ",sheetName:" + sheet.Name, "格式审核通过");
                        }
                        if (isPassGongshi)
                        {
                            AddError(workbook.FileName + ",sheetName:" + sheet.Name, "公式审核通过");
                        }

                    }
                }
                cellsTemplate.Clear();
                cellsTemplate.ImportDataTable(tableResult, true, 0, 0, tableResult.Rows.Count, tableResult.Columns.Count);
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
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
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

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtSource.Text) || string.IsNullOrEmpty(cbbExcuteTable.Text))
            {
                MessageBox.Show("请选择要处理的Excel,并选择要处理的表");
                return;
            }
            if (!string.IsNullOrEmpty(txtTemplate.Text) && CbbTemplateSheet.SelectedItem == null)
            {
                workCheck.ReportProgress(0, "请选择模版的sheet");
                return;
            }
            else if(CbbTemplateSheet.SelectedItem!=null)
            {
                TemplateSheetName = CbbTemplateSheet.SelectedItem.ToString();
            }
            txtLog.Clear();
            btnStart.Text = "正在处理...";
            btnStart.Enabled = false;
            btnClose.Enabled = false;
            listExcuteTable = new List<string>(cbbExcuteTable.Text.Split(','));
            IsMohu = rbMohu.Checked;
            tableResult.Clear();
            workCheck.RunWorkerAsync();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Helper.OpenHelpFile(8);
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
    }
}
