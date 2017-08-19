
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Data.SQLite;
using System.Data;
using System.IO;


namespace SQLiteToExcel
{
    class Program
    {
        static string MsgText = "";
        static void Main(string[] args)
        {
            //获取程序所在的位置
            string nowpath = AppDomain.CurrentDomain.BaseDirectory;
            //args = new string[1];
            //args[0] = "D:\\WordSpace\\SqliteToExcel\\SQLiteToExcel\\SQLiteToExcel\\bin\\Debug\\test2";
            int count = 0;

            Console.WriteLine("请选择导出的方式：\n1:导出txt文件(速度快)\n2:导出excel文件(速度慢)");
            string type = Console.ReadLine();
            if (type == "2")
                ShowMsg("=====================================\n已选择 2 :\t导出excel文件\n=====================================");
            else
                ShowMsg("=====================================\n已选择 1 :\t导出txt文件\n=====================================");

            if (args.Length >= 1)
            {
                int index = 1;
                //获取拖入的文件
                foreach (string f in args)
                {
                    string path = Path.GetDirectoryName(f);
                    string filename = Path.GetFileName(f);
                    string Extension = Path.GetExtension(f);
                    string outpath = CreatBackDic(path);
                    Console.Title = string.Format("正在读取... {0}", filename);
                    try
                    {
                        //filepath = nowpath + filename;
                        string connStr = "Data Source=" + f;
                        SQLiteConnection conn = new SQLiteConnection(connStr);
                        conn.Open();
                        System.Data.DataTable schemaTable = conn.GetSchema("TABLES");
                        string[] tableNames = GetTableName(conn);

                        if (tableNames == null)
                        {
                            ShowMsg(string.Format("({0}/{1})没有数据... {2}.", index, args.Length, filename));
                            index++;
                            continue;
                        }
                        string sql;
                        ShowMsg(string.Format("({0}/{1})正在读取... {2}.", index, args.Length, filename));
                        ArrayList tableText = new ArrayList();
                        foreach (string table in tableNames)
                        {
                            sql = string.Format("SELECT * FROM '{0}'", table);
                            tableText = GetSqliteTable(table, conn);

                            if (tableText.Count <= 0)
                                break;
                            else
                            {
                                if (type == "2")
                                {
                                    CreateExcel(filename + "_" + table, outpath, tableText);
                                }
                                else
                                {
                                    CreateFille(filename + "_" + table + ".txt", outpath, tableText);
                                }
                                count++;
                            }
                        }
                        conn.Close();
                        index++;
                    }
                    catch
                    {
                        ShowMsg(string.Format("({0}/{1})已忽略非数据库文件: {2}", index, args.Length, filename));
                        index++;
                        continue;
                    }
                }
            }
            else
            {
                ShowMsg("\n没有找到文件，请把需要导出文件拖到程序中..");
            }

            ShowMsg(string.Format("\n\n导出完成!\n本次共导出 {0} 个文件，忽略了 {1} 个文件。\n点击任意键关闭窗口。", count, (args.Length - count).ToString()));
            Console.Read();//防止闪退  
        }


        static ArrayList GetSqliteTable(string table, SQLiteConnection conn)
        {
            ArrayList tableText = new ArrayList();

            string sql = string.Format("PRAGMA table_info({0})", table);
            //查询字段名
            SQLiteDataAdapter dataAdapter = new SQLiteDataAdapter(sql, conn);
            DataSet dataSet = new DataSet();
            dataSet.EnforceConstraints = false;
            dataAdapter.Fill(dataSet, "newtable");
            System.Data.DataTable dataTable = dataSet.Tables["newtable"];
        
            string[] info;
            if (dataTable != null && dataTable.Rows.Count > 0)
            {
                info = new string[dataTable.Rows.Count];
                int index = 0;
                foreach (DataRow dataRow in dataTable.Rows)
                {
                    info[index] = dataRow[1].ToString();
                    index++;
                }
                tableText.Add(info);
            }

            //查询信息
            sql = string.Format("SELECT * FROM '{0}'", table);
            dataAdapter = new SQLiteDataAdapter(sql, conn);
            dataSet = new DataSet();
            dataSet.EnforceConstraints = false;
            dataAdapter.Fill(dataSet, "newtable");
            dataTable = dataSet.Tables["newtable"];

            if (dataTable != null && dataTable.Rows.Count > 0)
            {
                foreach (DataRow dataRow in dataTable.Rows)
                {
                    info = new string[dataTable.Columns.Count];
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        info[i] = dataRow[i].ToString();
                    }
                    tableText.Add(info);
                }
            }
            else
                tableText = new ArrayList();

            return tableText;
        }

        static void ShowMsg()
        {
            Console.Clear();
            Console.Write(MsgText);
        }
        static void ShowMsg(string text)
        {
            Console.Clear();
            MsgText += text+"\n";
            Console.WriteLine(MsgText);

        }

        static string[] GetTableName(SQLiteConnection conn)
        {
            System.Data.DataTable schemaTable = conn.GetSchema("TABLES");

            if (schemaTable == null || schemaTable.Rows.Count == 0)
                return null;

            //获取数据表
            string[] tableNames = new string[schemaTable.Rows.Count];
            int index = 0;
            foreach (DataRow dataRow in schemaTable.Rows)
            {
                tableNames[index] = dataRow["TABLE_NAME"].ToString();
                index++;
            }

            return tableNames;
        }


        static void CreateExcel(string filename, string path, ArrayList list)
        {
            Application excel = new Application();
            Workbooks wbks = excel.Workbooks;
            Workbook wb = wbks.Add(true);
            Worksheet wsh = wb.Sheets[1];
            
            Console.Title = string.Format("正在写入... {0}", filename);

            //开始写入值
            for (int i = 0; i < list.Count; i++)
            {
                string[] info = (string[])list[i];
                for (int j = 0;j< info.Length;j++)
                {
                    //string newstr = info[j].Replace("\r\n", "\\r\\n");
                    if (i == 0)
                    {
                        wsh.Cells[i + 1, j + 1].Font.Bold = true;
                        wsh.Cells[i + 1, j + 1].Interior.ColorIndex = 15;
                    }
                    wsh.Cells[i+1, j+1].Value = info[j];
                    if(i%10==0&&j==0)
                    {
                        ShowMsg();
                        Console.Write("正在写入:" + filename + ".xlsx");
                        Console.WriteLine("({0}/{1})", i,list.Count);
                    }
                }
            }
            
            excel.DisplayAlerts = false;
            excel.AlertBeforeOverwriting = false;
            wb.SaveAs(path + filename + ".xlsx");
            wb.Close();
            wbks.Close();
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            System.GC.Collect();
            ShowMsg(string.Format("\t写入成功！{0}", filename));
        }

        static void CreateFille(string file, string path, ArrayList notes)
        {
            string filename = path + file;
            if (!File.Exists(filename))
            {
                FileStream fs1 = new FileStream(filename, FileMode.Create, FileAccess.Write);//创建写入文件 
                StreamWriter sw = new StreamWriter(fs1);

                //开始写入值
                for (int i = 0; i < notes.Count; i++)
                {
                    string[] info = (string[])notes[i];
                    foreach (string @str in info)
                    {
                        string newstr = str.Replace("\r\n", "\\r\\n");
                        sw.Write(newstr + "\t");
                    }
                    sw.Write("\n");
                }
                sw.Close();
                fs1.Close();
            }
            else
            {
                FileStream fs = new FileStream(filename, FileMode.Truncate, FileAccess.Write);
                StreamWriter sr = new StreamWriter(fs);

                //开始写入值
                for (int i = 0; i < notes.Count; i++)
                {
                    string[] info = (string[])notes[i];
                    foreach (string str in info)
                    {
                        string newstr = str.Replace("\r\n", "\\r\\n");
                        sr.Write(newstr + ",");
                    }
                    sr.Write("\n");
                }
                sr.Close();
                fs.Close();
            }
            ShowMsg(string.Format("\t写入成功！{0}", file));
        }

        static string CreatBackDic(string path)
        {
            string filename = "SQLiteOutput";
            //创建backup文件夹
            if (!Directory.Exists(path + "//"+filename))
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(path + "/" + filename);
                directoryInfo.Create();
            }
            return path + "\\" + filename + "\\";
        }
    }
}