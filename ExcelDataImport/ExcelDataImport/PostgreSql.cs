using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Npgsql;

namespace ExcelDataImport
{
    class PostgreSql
    {
        private NpgsqlConnection conn = null;

        public PostgreSql(string connectionStringRw)
        {
            try
            {
                NpgsqlConnection conn = new NpgsqlConnection(connectionStringRw);
                conn.Open();
                this.conn = conn;
            }
            catch (Exception e)
            {
                Console.WriteLine($"DB 연결 실패 : {e.Message}");
                return;
            }
        }
    }
}
