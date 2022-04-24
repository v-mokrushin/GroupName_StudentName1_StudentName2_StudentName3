using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project
{
    public static class MathLibrary
    {

        public static double GetListElementsSum(List<double> list)
        {
            double sum = 0;
            for (int i = 0; i < list.Count; i++) sum += list.ElementAt(i);
            return sum;
        }

        public static double GetListElementsSum(List<double> list, int indexFrom, int indexTo)
        {
            double sum = 0;
            for (int i = indexFrom; i <= indexTo; i++) sum += list.ElementAt(i);
            return sum;
        }

        public static double RoundTo3(double value)
        {
            return Math.Round(value, 3);
        }

    }
}
