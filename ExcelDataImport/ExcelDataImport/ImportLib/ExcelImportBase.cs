using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RowsRetType = System.Collections.Generic.List<System.Collections.Generic.Dictionary<string, string>>;
using RowRetType = System.Collections.Generic.Dictionary<string, string>;

namespace ImportLib
{
    public abstract class ExcelImportBase
    {
        public string FileName { get; protected set; }
        public string TableName { get; protected set; }
        public string[] ColumnNames { get; protected set; }
        protected RowsRetType Rows = new RowsRetType();
        public IEnumerable<RowRetType> GetRows() => Rows;

        protected int id = 0;

        public abstract object[] GetValues(Dictionary<string, string> Row);

        public virtual void ResetId()
        {
            id = 0;
        }
    }
}
