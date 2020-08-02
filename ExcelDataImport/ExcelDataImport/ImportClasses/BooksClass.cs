using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ImportLib;
using RowRetType = System.Collections.Generic.Dictionary<string, string>;

namespace ImportClasses
{
    class BooksClass : ExcelImportBase
    {
        public BooksClass() : base(ExcelName.Books)
        {
            SetTableInfo("books", new String[] { "bookid", "author"});
        }

        public override object[] GetValues(RowRetType Row)
        {
            return new object[]
            {
                GetValue<int>(Row, "bookid"),
                GetValue<string>(Row, "author"),
            };
        }
    }
}
