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
    public partial class Test : Form
    {
        public Test()
        {
            InitializeComponent();
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Workbook w1 = new Workbook(@"D:\coding\Excel汇总\Desktop\复件 复件 苏州----附件 保费支出的财务影响测算数据采集表--苏州上报.xls");
            Cells c1 = w1.Worksheets[0].Cells;
            DataTable table1 = c1.ExportDataTable(0, 0, c1.MaxDataRow + 1, c1.MaxDataColumn + 1, true);
            dataGridView1.DataSource = table1;
        }
    }
}
