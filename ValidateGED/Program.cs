using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel;

namespace TestApp
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
       
        static void Main(string[] args)
        {
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());

            string fileName = "";
            if (args.Length == 1)
                fileName = args[0];

            if (String.IsNullOrWhiteSpace(fileName) || !File.Exists(fileName))
            {
                Console.WriteLine("Please enter the valid filename with full path.");
                return;
            }

            //fileName = "C:\shared\temp\1crq8_LFeh9vr7R7-thRf3GLtXHnqRd4dD6Bt2RzALyw-16686.xlsx";
            Stream stream = File.OpenRead(fileName);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            try
            {
                DataSet result = excelReader.AsDataSet();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
                return;
            }

            Console.WriteLine("GED excel file is valid.");
            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }
    }
}
