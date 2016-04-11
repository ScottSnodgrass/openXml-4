using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ElancoPimsDdsParser
{
    class Program
    {
        static void Main(string[] args)
        {           
            DDSParser batchDataParser = new DDSParser();
            string fullPath = args[0];
            string path = System.IO.Path.GetDirectoryName(fullPath);
            //string filename = System.IO.Path.GetFileName(fullPath);
            //string extension = System.IO.Path.GetExtension(filename);
            string filenameMinusExtension = System.IO.Path.GetFileNameWithoutExtension(fullPath);
            string outputFile = path + @"\" + filenameMinusExtension + ".xlsx";

            bool success = batchDataParser.processWordDoc(fullPath);
            
            if (success)
            {
                success = batchDataParser.buildExcelWorkbook(outputFile);
            }

            Console.WriteLine("Success: {0}", success);
            Console.ReadKey();
        }
    }
}
