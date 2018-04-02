using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Excel汇总
{
    public class ModelStore
    {
        public int row { get; set; }
        public int column { get; set; }

        public CellValueType cellType { get; set; }

        public bool IsFormula { get; set; }
        public string Formula { get; set; }
    }
}
