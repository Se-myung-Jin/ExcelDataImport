﻿using System;
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

            return 0;
        }
    }
}
