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
        public enum eIdType : sbyte
        {
            Default     = 0, // 기본 테이블
            Generate    = 1, // id가 1씩 증가하는 테이블
            Overlap     = 2, // id가 중복인 테이블
        }

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

        #region ExecuteQuery
        public bool ExecuteQuery(string query, out NpgsqlDataReader results)
        {
            NpgsqlCommand cmd = null;
            results = null;

            try
            {
                cmd = new NpgsqlCommand(query, conn);
                results = cmd.ExecuteReader();
            }
            catch (PostgresException e)
            {
                Console.WriteLine($"query: {cmd.CommandText}");
                Console.WriteLine($"Msg: {e.Message}");
                Console.WriteLine($"Detail: {e.Detail}");
                return false;
            }
            catch (NpgsqlOperationInProgressException e)
            {
                Console.WriteLine($"error : {e.Message}");
                return false;
            }
            return true;
        }

        public bool ExecuteQuery(string query, ref List<Dictionary<String, object>> Rows)
        {
            if (ExecuteQuery(query, out var results))
            {
                if (results != null && results.IsClosed == false)
                {
                    while (results.Read())
                    {
                        var row = new Dictionary<String, object>();

                        for (int i = 0; i < results.FieldCount; i++)
                        {
                            string key = results.GetName(i);
                            object value = results[i];

                            row.Add(results.GetName(i), value);
                        }
                        Rows.Add(row);
                    }
                    results.Close();
                    return true;
                }
            }
            return true;
        }

        public bool ExecuteQuery (string query)
        {
            var cmd = new NpgsqlCommand(query);
            return ExecuteQuery(cmd);
        }

        public bool ExecuteQuery (NpgsqlCommand cmd)
        {
            try
            {
                cmd.Connection = conn;
                cmd.ExecuteNonQuery();
            }
            catch (PostgresException e)
            {
                Console.WriteLine($"query: {cmd.CommandText}");
                Console.WriteLine($"Msg: {e.Message}");
                Console.WriteLine($"Detail: {e.Detail}");
                return false;
            }
            return true;
        }
        #endregion
    }
}
