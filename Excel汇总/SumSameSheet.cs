using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using Aspose.Cells;
using System.Collections;

namespace Excel汇总
{
    public partial class SumSameSheet : BaseControl
    {
        public SumSameSheet()
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
                bool isUseTemplate = false;
                workCheck.ReportProgress(0, "开始处理...");
                List<ModelStore> list = new List<ModelStore>();
                string[] listAllExcel = DirFileHelper.GetFileNames(txtSource.Text, "*.xls", true);
                Workbook workbookTemplate;
                if (!string.IsNullOrEmpty(txtTemplate.Text))
                {
                    workbookTemplate = new Workbook(txtTemplate.Text);
                    isUseTemplate = true;
                }
                else
                {
                    workbookTemplate = new Workbook(listAllExcel[0]);
                }
                Cells cellsTemplate = workbookTemplate.Worksheets[0].Cells;
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
                    Cells cells = workbook.Worksheets[0].Cells;
                    workCheck.ReportProgress(0, "开始处理数据...---" + workbook.FileName);
                    for (int i = 0; i < cells.MaxDataRow + 1; i++)
                    {
                        for (int j = 0; j < cells.MaxDataColumn + 1; j++)
                        {
                            if (list.FindIndex(x => x.row == i && x.column == j) >= 0 && cells[i, j].Type == CellValueType.IsNumeric)
                            {
                                if (isUseTemplate && cellsTemplate[i, j].Type != CellValueType.IsNumeric)
                                {
                                    cellsTemplate[i, j].PutValue(cells[i, j].IntValue);
                                }
                                else
                                {
                                    cellsTemplate[i, j].PutValue(cellsTemplate[i, j].IntValue + cells[i, j].IntValue);
                                }

                            }

                        }
                    }
                }

                workbookTemplate.Save(txtResult.Text);
                workCheck.ReportProgress(0, "处理完成");
            }
            catch (Exception ex)
            {
                txtLog.AppendText(ex.Message);
            }

        }
        void workCheck_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnStart.Text = "开始";
            btnStart.Enabled = true;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            txtSource.Text = Helper.SelectPath();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            txtTemplate.Text = Helper.SelectFile();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            txtResult.Text = Helper.SelectFile();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtSource.Text))
            {
                MessageBox.Show("请选择要处理的Excel");
                return;
            }
            txtLog.Clear();
            btnStart.Text = "正在处理...";
            btnStart.Enabled = false;
            workCheck.RunWorkerAsync();

        }
    }
}
