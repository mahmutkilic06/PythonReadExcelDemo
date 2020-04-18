using System;
using System.Diagnostics;
using System.IO;

namespace PythonReadExcelDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Merhaba Pyton!!");
            var result = ReadFromPython("Detay Raporu Aralık.xlsx","Detay Raporu");

            Console.WriteLine(result);
            Console.ReadLine();
        }

        public static string ReadFromPython(string ExcelPath, string SheetName)
        {
            var standardError = string.Empty;

            var pythonPath = GetPythonPath();

            var path = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).Replace("file:\\", ""); ;
            var pyPath = string.Concat(path, "\\ExcelReader.py");

            pyPath = "\"" + pyPath + "\"";
            ExcelPath = "\"" + ExcelPath + "\"";
            SheetName = "\"" + SheetName + "\"";

            var parameter = string.Format("{0} {1} {2}", pyPath, ExcelPath, SheetName);

            Process p = new Process();
            p.StartInfo = new ProcessStartInfo(pythonPath, parameter)
            {
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardError = true,
            };
            p.Start();

            string output = p.StandardOutput.ReadToEnd();
            standardError = p.StandardError.ReadToEnd();
            p.WaitForExit();

            if (!string.IsNullOrEmpty(standardError))
            {
                throw new Exception(standardError);
            }
            //var result = JsonConvert.DeserializeObject<List<T>>(output);

            return output;
        }

        static string GetPythonPath()
        {
            var result = string.Empty;
            var environmentVariables = Environment.GetEnvironmentVariables();
            string pathVariable = environmentVariables["Path"] as string;
            if (pathVariable != null)
            {
                string[] allPaths = pathVariable.Split(';');
                foreach (var path in allPaths)
                {
                    string pythonPathFromEnv = path + "python.exe";
                    if (File.Exists(pythonPathFromEnv))
                        result = pythonPathFromEnv;
                }
            }
            return result;
        }

    }
}
