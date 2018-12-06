using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SalesReport
{
    public struct Globals
    {
        private static string userName = Environment.UserName;
        public static string resourcesPath { get { return @"C:\Users\" + userName + @"\OneDrive\Documents\Excel Sheets\"; } }
        public static string savePath { get { return @"C:\Users\" + userName + @"\OneDrive\Documents\Excel Sheets\"; } }
    }

    public struct Data
    {
        internal string compName;
        internal int bMonth;
        internal int eMonth;
        internal int _period { get{ return (eMonth - bMonth) + 1; } }
        internal double[] sales;
    }
}
