using System;
using System.Collections.Generic;
using System.Text;

using Power_Flow.类文件;
using System.IO;

namespace PQ_PF
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Please input diskette data-file name:\r");
            string fileName = Console.ReadLine();
            if (!File.Exists(fileName))
            {
                Console.WriteLine("File name:" + System.IO.Path.GetFileName(fileName)+" could not be found in this path!"+
                    " Please check the file name!");
                Environment.Exit(1);
            }
            PF_Calc pf = new PF_Calc();
            pf.PF_Main(System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase+fileName);

        }
    }
}
