using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ElancoPimsDdsParser
{
    public class AspenScript
    {
        XNamespace w = Constants.XmlSchemas.wordmlNamespace;
        public string OperationName { get; set; }
        public string PhaseName { get; set; }
        public string ScriptName { get; set; }
        public string Trigger { get; set; }

        public string ScriptText { get; set; }

        List<AspenScriptVariable> variables = new List<AspenScriptVariable>();

        public AspenScript(string[] identifiers)
        {
            OperationName = String.Copy(identifiers[0]);
            PhaseName = String.Copy(identifiers[1]);
            ScriptName = String.Copy(identifiers[2]);
        }

        public bool parseTable(XElement tableNode)
        {
            var allRows = from x in tableNode.Elements(w + "tr")
                        select x;
            int numRows = allRows.Count();

            bool foundTrigger = false, foundVars = false, foundScript = false;
            
            for (int rowCounter = 0; rowCounter < numRows; rowCounter++)            
            {
                var tr = allRows.ElementAt(rowCounter);
                var cells = tr.Elements(w + "tc");

                if (cells.Count() == 1)
                {
                    string cellTxt = XUtils.WordProcessingMLUtils.getParagraphTextFromCell(cells.FirstOrDefault());
                    if (cellTxt.Equals(Constants.ElancoDocConstants.TriggerTitle))
                    {
                        tr = allRows.ElementAt(++rowCounter);
                        cells = tr.Elements(w + "tc");
                        Trigger = XUtils.WordProcessingMLUtils.getParagraphTextFromCell(cells.FirstOrDefault());
                        foundTrigger = true;
                        continue;
                    }
                    if (cellTxt.Equals(Constants.ElancoDocConstants.VariableBindingTitle))
                    {
                        rowCounter++;
                        if (!parseScriptVars(allRows, ref rowCounter))
                        {
                            return false;
                        }
                        foundVars = true;
                        continue;
                    }
                    if (cellTxt.Equals(Constants.ElancoDocConstants.ScriptTitle))
                    {
                        tr = allRows.ElementAt(++rowCounter);
                        cells = tr.Elements(w + "tc");
                        ScriptText = XUtils.WordProcessingMLUtils.getParagraphTextFromCell(cells.FirstOrDefault());
                        foundScript = true;
                    }
                }
                
                rowCounter++; 
            } // end of for loop

            if (!foundTrigger || !foundVars || !foundScript)
            {
                Console.WriteLine("ERROR - improperly formatted Characteristics Table!");
                return false;
            }
            return true;
        }

        private bool parseScriptVars(IEnumerable<XElement> allRows,ref int rowIndex)
        {            
            var tr = allRows.ElementAt(rowIndex);
            int numTotalRows = allRows.Count(); // note that remaing rows includes
            // those not processed by this method.
            var cells = tr.Elements(w + "tc");
            int numCells = cells.Count();
            if (!verifyVarsRowsHeader(cells)) // first row
            {
                return false;
            }
            // skip the header row
            rowIndex++;
            tr = allRows.ElementAt(rowIndex);
            cells = tr.Elements(w + "tc");
            while ((rowIndex < numTotalRows) && (numCells == AspenScriptVariable.NumScriptVarProperties))
            {                
                //Console.WriteLine("num cells: {0:D}", numCells);
                variables.Add(new AspenScriptVariable(cells));

                tr = allRows.ElementAt(++rowIndex);
                cells = tr.Elements(w + "tc");
                numCells = cells.Count();

            }// end of while

            return true;
        }
        // if the script text is spread out over multiple tables - this 
        // model will not work.
        // TODO: add check for tables that do no conform in the table list 

        private bool verifyVarsRowsHeader(IEnumerable<XElement> cells)
        {
            if (cells.Count() != AspenScriptVariable.NumScriptVarProperties)
            {
                return false;
            }
            /* this is quick and fugly */
            string txt0 = XUtils.WordProcessingMLUtils.
                        getParagraphTextFromCell(cells.ElementAt(0));
            string txt1 = XUtils.WordProcessingMLUtils.
                getParagraphTextFromCell(cells.ElementAt(1));
            string txt2 = XUtils.WordProcessingMLUtils.
                getParagraphTextFromCell(cells.ElementAt(2));
            string txt3 = XUtils.WordProcessingMLUtils.
                getParagraphTextFromCell(cells.ElementAt(3));

            if (!(txt0.Equals(AspenScriptVariable.ScriptVarPropertyNames[0])) ||
                !(txt1.Equals(AspenScriptVariable.ScriptVarPropertyNames[1])) ||
                !(txt2.Equals(AspenScriptVariable.ScriptVarPropertyNames[2])) ||
                !(txt3.Equals(AspenScriptVariable.ScriptVarPropertyNames[3])) 
                )
            {
                return false;
            }
            return true;
        }
        
    }

    public class AspenScriptVariable
    {
        public const int NumScriptVarProperties = 4;
        public readonly static string[] ScriptVarPropertyNames = { "Name", "Mode", "Binding Type", "Binding" };

        public string Name { get; set; }
        public string Mode { get; set; }
        public string BindingType { get; set; }
        public string BindingText { get; set; }

        public AspenScriptVariable(IEnumerable<XElement> cells)
        {
            Name = XUtils.WordProcessingMLUtils.
                    getParagraphTextFromCell(cells.ElementAt(0));
            Mode = XUtils.WordProcessingMLUtils.
                    getParagraphTextFromCell(cells.ElementAt(1));
            BindingType = XUtils.WordProcessingMLUtils.
                    getParagraphTextFromCell(cells.ElementAt(2));
            BindingText = XUtils.WordProcessingMLUtils.
                    getParagraphTextFromCell(cells.ElementAt(3));

        }
    }
}
