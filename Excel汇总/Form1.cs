using Aspose.Cells;
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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public class ModelStore
        {
            public int row { get; set; }
            public int column { get; set; }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<ModelStore> list = new List<Excel汇总.Form1.ModelStore>();
            int s = 0;
            Workbook workbook = new Workbook();
            workbook.Open(Application.StartupPath + "\\北京分公司2006年度销售、成本报表.xls");
            Workbook workbook2 = new Workbook();
            workbook2.Open(Application.StartupPath + "\\上海分公司2006年度销售、成本报表.xls");
            Cells cells = workbook.Worksheets[0].Cells;
            Cells cells2 = workbook2.Worksheets[0].Cells;
            string str = string.Empty;
            for (int i = 0; i < cells.MaxDataRow + 1; i++)
            {
                for (int j = 0; j < cells.MaxDataColumn + 1; j++)
                {
                    if (cells[i, j].IsFormula)
                    {
                        str += i + "-" + j;
                    }
                    if (cells[i, j].Type == CellValueType.IsNumeric)
                    {
                        list.Add(new ModelStore { row = i, column = j });
                      

                    }

                }
            }

            for (int i = 0; i < cells2.MaxDataRow + 1; i++)
            {
                for (int j = 0; j < cells2.MaxDataColumn + 1; j++)
                {
                    if (list.FindIndex(x => x.row == i && x.column == j) >= 0 && cells2[i, j].Type == CellValueType.IsNumeric)
                    {
                          cells[i, j].PutValue(cells[i, j].IntValue + cells2[i, j].IntValue);
                      //  cells[i, j].PutValue(11111);
                    }

                }
            }
            workbook.Save(Application.StartupPath + "\\result.xls");
            MessageBox.Show("操作成功"+str);
        }
    }
}
