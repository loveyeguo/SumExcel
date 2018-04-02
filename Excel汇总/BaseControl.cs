using CCWin;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Windows.Forms;
using Aspose.Cells;

namespace Excel汇总
{
    public class BaseControl : CCSkinMain
    {
        protected string TemplateTag = "#tag#";
        protected List<string> listExcuteTable = new List<string>();
        protected string TemplateSheetName;
        protected bool IsMohu = false;
        protected bool IsExcuteSheet(string sheetName)
        {
            foreach (var item in listExcuteTable)
            {
                if (IsMohu)
                {
                    if (sheetName.Contains(item))
                    {
                        return true;
                    }
                }
                else
                {
                    if (sheetName==(item))
                    {
                        return true;
                    }
                }
                
            }
            return false;
        }
        protected void InitSingleCheckCbb(string txtExcleUrl, ComboBox cbb)
        {
            cbb.Items.Clear();
            if (string.IsNullOrEmpty(txtExcleUrl))
            {
                return;
            }

            Workbook wb = new Workbook(txtExcleUrl);
            foreach (var sheet in wb.Worksheets)
            {
                cbb.Items.Add(sheet.Name);
            }
        }
    }
}
