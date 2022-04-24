using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Project
{
    public static class LogLibrary
    {

        private static string logFilePath = @"C:/Users/TheIm/Desktop/log.txt";

        public static void Log(string info)
        {
            using (StreamWriter file = new StreamWriter(logFilePath, true, System.Text.Encoding.Default))
            {
                file.WriteLine(DateTime.Now + " | " + info);
            }
        }

        public static void Log(int info)
        {
            using (StreamWriter file = new StreamWriter(logFilePath, true, System.Text.Encoding.Default))
            {
                file.WriteLine(DateTime.Now + " | " + info);
            }
        }

    }
}
