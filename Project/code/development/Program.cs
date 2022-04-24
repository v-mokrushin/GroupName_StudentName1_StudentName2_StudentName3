using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Console;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp1
{

    class Program
    {

        static void Main(string[] args)
        {

            XExcelTable tableDefaultHorizontal = new XExcelTable();
            tableDefaultHorizontal.ReadDefaultHorizonatExcelFile();
            Console.WriteLine("|XExcelTable tableDefaultHorizontal| reading complete.");

            XExcelTable tableDefaultVertical = new XExcelTable();
            //tableDefaultVertical.ReadDefaultVerticalExcelFile();
            //tableDefaultVertical.PrintMeaningsTable();
            //tableDefaultVertical.PrintTableParameters();

            XExcelTable tableTest = new XExcelTable();
            tableTest.ReadTestExcelFile();
            Console.WriteLine("|XExcelTable tableTest| reading complete.");

            tableTest.SetDefaultTable(tableDefaultHorizontal);
            tableTest.CreateModsAndBoolsTables();
            //tableTest.CreateExcelReport();

            //Console.ReadKey();
        }

    }
}