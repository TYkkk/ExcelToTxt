using Excel;
using System.Data;
using System.IO;
using System;
using System.Text;

namespace ExcelToTxt
{
    class Program
    {
        private static FileStream stream;
        private static IExcelDataReader excelReader;

        private static string SavePath = null;

        static void Main(string[] args)
        {
            var path = @"C:\Users\zhaolingzhu\Documents\work\svn\04.Config\MyShop\myshop.seasons.xlsx";
            WriteToTxt(OpenExcel(path));

            Console.ReadLine();
        }

        private static DataSet OpenExcel(string strFileName)
        {
            if (!strFileName.EndsWith(".xlsx") && !strFileName.EndsWith(".xls"))
            {
                Console.WriteLine("发生错误，需选择正确的.xlsx||.xls文件");
                return null;
            }

            DataSet result = null;
            try
            {
                stream = File.Open(strFileName, FileMode.Open, FileAccess.Read);
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                result = excelReader.AsDataSet();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (excelReader != null) excelReader.Dispose();
                if (stream != null) stream.Dispose();
            }

            return result;
        }

        private static void WriteToTxt(DataSet data)
        {
            while (string.IsNullOrEmpty(SavePath) || !Directory.Exists(SavePath))
            {
                Console.WriteLine("输入保存配置路径");
                SavePath = Console.ReadLine();
            }

            for (int i = 0; i < data.Tables.Count; i++)
            {
                string tabName = data.Tables[i].TableName;

                if (tabName.StartsWith("#"))
                {
                    string saveName = tabName.Substring(tabName.IndexOf('#') + 1, tabName.Length - 1);
                    string filePath = SavePath + saveName + ".txt";
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }

                    StringBuilder sb = new StringBuilder();

                    for (int row = 1; row < data.Tables[i].Rows.Count; row++)
                    {
                        for (int col = 0; col < data.Tables[i].Columns.Count; col++)
                        {
                            sb.Append(data.Tables[i].Rows[row][col].ToString());
                            Console.Write(data.Tables[i].Rows[row][col].ToString());
                            if (col != data.Tables[i].Columns.Count - 1)
                            {
                                sb.Append(",");
                                Console.Write(",");
                            }
                        }

                        sb.Append("\n");
                        Console.WriteLine();
                    }

                    File.WriteAllText(filePath, sb.ToString());
                }
            }
        }

        public static int ToInt(string text, int defaultValue = 0)
        {
            int result;

            if (!int.TryParse(text, out result))
            {
                result = defaultValue;
            }

            return result;
        }
    }
}
