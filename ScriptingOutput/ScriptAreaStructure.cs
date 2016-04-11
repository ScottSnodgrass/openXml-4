using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ElancoPimsDdsParser.ScriptingOutput
{
    class ScriptAreaStructure
    {
        public static bool generateAreaStructure(List<Operation> listOp, WorkbookPart workbookPart)
        {
            Console.WriteLine("Creating AreaStructure tab of Excel");

            string sheetName = "AreaStructure";
            uint row, col;
            row = 3;  // starting row position
            col = 1;
            foreach (var op in listOp)
            {
               // Console.WriteLine("[{0:D}, {1:D}]", row, col);
                ExcelGenerator.addCellData(workbookPart, sheetName, row++, col, op.Name);

                foreach (var phase in op.getPhases())
                {
                 //   Console.WriteLine("[{0:D}, {1:D}]", row, col);
                    ExcelGenerator.addCellData(workbookPart, sheetName, row, col, op.Name);
                //    Console.WriteLine("[{0:D}, {1:D}]", row, col + 1);
                    ExcelGenerator.addCellData(workbookPart, sheetName, row++, col + 1, phase.Name);
                }
            }


            return true;
        }
    }
}
