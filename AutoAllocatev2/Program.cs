using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
/* To work eith EPPlus library */
using OfficeOpenXml;
using OfficeOpenXml.Drawing;

/* For I/O purpose */
using System.IO;

/* For Diagnostics */
using System.Diagnostics;
using System.Data;

using System.Drawing;
using System.Runtime.Serialization.Formatters.Binary;

namespace AutoAllocatev2
{
    class Program
    {

        static void Main(string[] args)
        {

            if (args.Length <= 0 || string.IsNullOrEmpty(args[0]))
            {
                args = new string[1];
                args[0] = @"..\..\..\Resource Allocation Sheet.xlsx";
                System.IO.File.Copy(@"..\..\..\Resource Allocation Sheet - Original.xlsx", args[0], true);
            }
            if (!System.IO.File.Exists(args[0]))
            {
                Console.WriteLine("{0} doesn't exists. Please provide a valid File Location.", args[0]);
                System.Environment.Exit(1000);
            }
            AutoAllocator.Allocate(args[0]);
        }
    }
}

