using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ImportLib;
using System.Collections.Concurrent;
using System.Configuration;

namespace ExcelDataImport
{
    class Program
    {
        public static ConcurrentQueue<ExcelImportBase> excelSheetQ = new ConcurrentQueue<ExcelImportBase>();
        public static ConcurrentQueue<Action> loadSheetTaskQ = new ConcurrentQueue<Action>();
        public static ConcurrentQueue<string> errorMsgQ = new ConcurrentQueue<string>();

        [STAThread]
        static int Main(string[] args)
        {
            String dirPathName;
            if (args.Length < 1)
            {
                FolderBrowserDialog dlg = new FolderBrowserDialog();

                if (dlg.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(dlg.SelectedPath))
                {
                    dirPathName = dlg.SelectedPath;
                }
                else
                {
                    MessageBox.Show("파일이 존재하는 폴더를 선택하세요.");
                    return -1;
                }
            }
            else
            {
                dirPathName = System.IO.Path.GetFullPath(args[0]);
            }

            if (!Directory.Exists(dirPathName))
            {
                MessageBox.Show("존재하는 폴더를 선택해 주세요.");
                return -1;
            }

            int retVal = LoadAllExcelSheet(dirPathName);

            while (true)
            {
                Console.WriteLine("0 : 종료, 1 : 테이블 검사 후 임포트");
                var input = Console.ReadLine();

                if (input == "0")
                    break;
                else if (input == "1")
                {
                    // import 추가
                    ExecuteImport(dirPathName);
                    break;
                }
            }

            Console.WriteLine("Press [Enter] to exit.");
            Console.ReadLine();

            return retVal;
        }

        #region Load all excel files
        public static int LoadAllExcelSheet(string dirPathName)
        {
            // 모든 엑셀 시트 로드

            return 0;
        }

        public static void LoadExcelSheetAsync<T>(string dirPath, string sheetName)
        {
            var act = new Action(() =>
            {
                bool load = LoadExcelSheet<T>(dirPath, sheetName, out var import);

                if (load)
                    excelSheetQ.Enqueue(import);
            });

            loadSheetTaskQ.Enqueue(act);
        }

        public static bool LoadExcelSheet<T>(string dirPath, string sheetName, out ExcelImportBase import)
        {
            import = Activator.CreateInstance(typeof(T)) as ExcelImportBase;

            return true;
        }
        #endregion

        public static int ExecuteImport(string dirPathName)
        {
            AppSettingsReader settingsReader = new AppSettingsReader();
            string connStr = settingsReader.GetValue("PostgreSQLConnectionString", typeof(string)) as string;

            Console.WriteLine(connStr);

            PostgreSql db = new PostgreSql(connStr);

            Task[] tasks = new Task[8];
            for (int i = 0; i < tasks.Length; i++)
            {
                tasks[i] = Task.Run(() =>
                {
                    PostgreSql taskdb = new PostgreSql(connStr);

                    while (true)
                    {
                        if (!excelSheetQ.TryDequeue(out var import))
                            return;

                        // db로 카피 추가
                    }
                });
            }

            foreach (var task in tasks)
                task.Wait();

            if (PrintErrorMsg())
                return 1;

            return 0;
        }

        public static bool PrintErrorMsg()
        {
            bool error = errorMsgQ.Count > 0;
            Console.ForegroundColor = ConsoleColor.Red;

            while (true)
            {
                if (!errorMsgQ.TryDequeue(out var errorMsg))
                    break;

                Console.WriteLine(errorMsg);
            }
            Console.ResetColor();

            return error;
        }

        public static bool TryCopyToDB(PostgreSql db, ExcelImportBase import)
        {
            var dbRows = new List<Dictionary<String, object>>();

            if (!db.ExecuteQuery($"select * from {import.TableName};", ref dbRows))
            {
                Console.WriteLine($"DB query failed, {import.TableName}");
                return false;
            }

            var dbValues = new List<object[]>();

            if (!GetAllValues(dbRows, import.ColumnNames, ref dbValues))
            {
                Console.WriteLine($"Get db query failed, {import.TableName}");
                return false;
            }

            var tableValues = new List<object[]>();

            import.ResetId();
            foreach (var row in import.GetRows())
            {
                try
                {
                    var tableValue = import.GetValues(row);
                    tableValues.Add(tableValue);
                }
                catch (Exception e)
                {
                    errorMsgQ.Enqueue($"{import.FileName} 테이블에서 예외가 발생하였습니다.");
                    return false;
                }
            }

            var addedTable = new List<object[]>();
            var missingTable = new List<object[]>();

            if (!GetAllChanges(ref tableValues, ref dbValues, ref addedTable, ref missingTable))
            {
                Console.WriteLine($"{import.TableName} 테이블 데이터의 바뀐 내용이 없습니다.");
                return true;
            }
            else
            {
                var modifiedTable = new List<object[]>();
                // 수정 작업
            }

            return true;
        }

        public static bool GetAllValues(List<Dictionary<String, object>> Rows, string[] ColumnNames, ref List<object[]> allValues)
        {
            for (int i = 0; i < Rows.Count; i++)
            {
                var row = Rows[i];
                object[] values = new object[ColumnNames.Length];

                for (int  j = 0; j < ColumnNames.Length; j++)
                {
                    var columnName = ColumnNames[j];
                    if (row.TryGetValue(columnName, out var value))
                    {
                        values[j] = value;
                    }
                    else
                    {
                        return false;
                    }
                }
                allValues.Add(values);
            }

            return true;
        }

        public static bool GetAllChanges(ref List<object[]> tableValues, ref List<object[]> dbValues, ref List<object[]> addedTable, ref List<object[]> missingTable)
        {
            tableValues = tableValues.OrderBy((val) => val?[0]).ToList();
            dbValues = dbValues.OrderBy((val) => val?[0]).ToList();

            bool isChange = false;

            for (int i = 0; i < tableValues.Count; i++)
            {
                object[] tableValue = tableValues[i];

                if (tableValue == null)
                    continue;

                bool find = false;
                int remove = -1;

                int tableNumberValue = -1;

                try
                {
                    tableNumberValue = Convert.ToInt32(tableValue[0]);
                }
                catch (Exception e)
                {

                }

                for (int j = 0; j < dbValues.Count; j++)
                {
                    object[] dbValue = dbValues[j];

                    if (dbValue == null)
                        continue;

                    int DBNumberValue = -1;

                    try
                    {
                        DBNumberValue = Convert.ToInt32(dbValue[0]);
                    }
                    catch (Exception e)
                    {

                    }

                    if (DBNumberValue > tableNumberValue)
                    {
                        find = false;

                        break;
                    }
                    else if (DBNumberValue != tableNumberValue)
                    {
                        remove = j;
                    }

                    for (int k = 0; k < dbValue.Length; k++)
                    {
                        var tableValueString = tableValue[k]?.ToString() ?? "null";
                        var dbValueString = dbValue[k]?.ToString() ?? "null";

                        dbValueString = dbValueString == "" ? "null" : dbValueString;

                        if (tableValueString != dbValueString)
                        {
                            find = false;
                            break;
                        }

                        if (k == dbValue.Length - 1)
                        {
                            find = true;
                        }
                    }

                    if (find)
                    {
                        dbValues.RemoveAt(j);
                        break;
                    }
                }

                if (!find)
                {
                    addedTable.Add(tableValue);

                    isChange = true;
                }

                if (remove != -1)
                {
                    for (int k = 0; k < remove; k++)
                    {
                        var removeValue = dbValues[0];
                        missingTable.Add(removeValue);
                        dbValues.RemoveAt(0);
                    }
                }
            }

            foreach (var dbValue in dbValues)
            {
                missingTable.Add(dbValue);
                isChange = true;
            }

            return isChange;
        }
    }
}
