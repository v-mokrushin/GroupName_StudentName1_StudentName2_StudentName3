using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace Project
{

    public static class StringLibrary
    {

        public static void SplitLine(string str, out string parameter, out string meaning)
        {

            // Разделить строку (например, из файла конфигурации) на параметр и значение параметра

            parameter = "";
            meaning = "";

            str = str.Trim();
            ushort spacePosition = 0;

            for (ushort i = 0; i < str.Length; i++)
            {
                if (str[i] != ' ')
                {
                    parameter += str[i];
                }

                else
                {
                    spacePosition = i;
                    break;
                }
            }

            for (ushort i = (ushort)(spacePosition + 1); i < str.Length; i++)
            {
                meaning += str[i];
            }

        }

        public static int GetMaxLength(List<string> list)
        {
            int maxLength = 0;
            for (int i = 0; i < list.Count; i++)
            {
                maxLength = Math.Max(maxLength, list[i].Length);
            }
            return maxLength;
        }

        public static string IncreaseStringToMaxLength(string inputString, int maxStringLenght)
        {
            int tt = maxStringLenght - inputString.Length;

            if (inputString.Length <= maxStringLenght)
            {
                for (int i = 0; i < tt; i++) inputString = inputString + '.';
                return inputString;
            }
            else
            {
                string newString = "";
                for (int i = 0; i < maxStringLenght - 5; i++) newString += inputString[i];
                newString += ".....";
                return newString;
            }

        }

        public static string AppendChar(string textBoxString)
        {
            if (textBoxString.IndexOf("_") != -1)
            {
                string newString = "";
                for (int i = 0; i < textBoxString.Length; i++)
                {
                    newString += textBoxString.ElementAt(i);
                    if (textBoxString.ElementAt(i) == '_') newString += textBoxString.ElementAt(i);
                }
                return newString;
            }
            else
            {
                return textBoxString;
            }
        }

    }

}