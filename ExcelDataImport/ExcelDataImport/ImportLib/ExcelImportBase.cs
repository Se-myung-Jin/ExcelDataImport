using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportLib
{
    public abstract class ExcelImportBase
    {
        public string FileName { get; protected set; }
        public string TableName { get; protected set; }
        public string[] ColumnNames { get; protected set; }
    }
}
