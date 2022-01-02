using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateInsertScripFromExcelSheet.cs
{
    public class BusinessRule
    {
        public int ColumnNumber { get; set; }

        public int[] Rule { get; set; }

        public string BusinessDerivedColumnValue { get; set; }
    }
}
