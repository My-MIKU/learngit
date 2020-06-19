using CLTabtoyV2.ExcelData;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using LitJson;
using System.Reflection;
using System.CodeDom.Compiler;

namespace CLTabtoyV2.Builder
{
    class JSONBuilder : IBuilder
    {
        private List<IExcelContext> _excelContexts;//Excel上下文
        private int _startRow;
        private Assembly _assembly;
        public JSONBuilder(Assembly assembly, List<IExcelContext> excelContexts)
        {
            Debug.Assert(excelContexts != null);
            _excelContexts = excelContexts;
            _assembly = assembly;
            _startRow = 0;
        }

        string innerBuilder(IExcelContext excelContext)
        {
            //Type[] types = _assembly.GetTypes();
            //foreach (var i in types) {
            //    Console.WriteLine(i.Name);

            //    MemberInfo[] temp = i.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);
            //    for (int j = 0; j < temp.Length - 1; j++) {

            //    }
            //    Console.WriteLine();
            //}

            try
            {
                foreach (DataTable dt in excelContext.DataSet.Tables)
                {
                    if (dt.Rows.Count <= 0) continue;
                    Table2JSON(dt, _startRow, dt.TableName, excelContext.FileName);
                }
            }
            catch (Exception e)
            {
                Debug.LogError($"{excelContext.FileName} 导出JSON时发生错误 错误信息:{e.Message} ");
            }
            return "";
        }

        public async Task<BuildResult> Build(Func<int, int> dispalyProgress = null)
        {
            if (Globals.FullExport)
            {
                Utils.ClearDirectory(Globals.JSONFileOutputPath);
            }
            StringBuilder sb = new StringBuilder();

            var allTasks = new List<Task<string>>();
            int count = 0;
            foreach (var tempExcelContext in _excelContexts)
            {
                IExcelContext excelContext = tempExcelContext;
                allTasks.Add(Task<string>.Run(
                    () => innerBuilder(excelContext)));
            }

            while (allTasks.Any())
            {
                var finished = await Task.WhenAny(allTasks);
                sb.AppendLine(finished.Result);
                allTasks.Remove(finished);
                dispalyProgress?.Invoke((int)(++count * 1f / _excelContexts.Count * 100));
            }

            var buildResult = new BuildResult(sb.ToString(), "");
            dispalyProgress?.Invoke(100);
            return buildResult;
        }

        private void Table2JSON(DataTable table, int startRow, string tableName, string fileName)
        {
            if (tableName.Contains('#')) return;

            string path = $"{Globals.JSONFileOutputPath}{fileName}_{tableName}.json";
            if (File.Exists(path))
            {
                File.Delete(path);
            }
            FileStream fs = new FileStream(path, FileMode.Create);

            using (StreamWriter sw = new StreamWriter(fs))
            {
                int rows = table.Rows.Count;
                int columns = table.Columns.Count;

                StringBuilder sb = new StringBuilder();
                for (int i = startRow; i < rows; i++)
                {
                    sb.Clear();
                    for (int j = 0; j < columns; j++)
                    {
                        sb.Append(table.Rows[i][j].ToString());

                        if (j < columns - 1)
                            sb.Append(',');
                    }
                    sw.WriteLine(sb);
                }
            }
            fs.Close();
            fs.Dispose();
        }
    }
}
