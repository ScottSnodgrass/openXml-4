using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ElancoPimsDdsParser
{
    class BatchUnit
    {
        XNamespace w = Constants.XmlSchemas.wordmlNamespace;

        List<Operation> operationList = new List<Operation>();

        public string Name { get; private set; }
        List<Operation> operations = new List<Operation>();

        public BatchUnit(string name)
        {
            Name = name;
        }

        public List<Operation> getOperationList()
        {
            return operationList;
        }

        public bool parseReportLayout(XElement reportTableRootNode)
        {
            bool success = true;

            XElement rowOfColHdrsElem = findTableRowWithNumCells(reportTableRootNode, Constants.ElancoDocConstants.NumCellsInLayoutTableRow);
            var nextSibling = rowOfColHdrsElem.NextNode;
            if (nextSibling.NodeType == System.Xml.XmlNodeType.Element)
            {
                XElement firstRowElem = (XElement)nextSibling;
                success = processLayoutTableRows(firstRowElem);
            }
            else
            {
                success = false;
            }

            if (!success)
            {
                Console.WriteLine("ERROR - There will be trouble");
            }

            return success;            
        }


        private bool processLayoutTableRows(XElement firstRowElem)
        {
            XElement nextSibling = firstRowElem; // this row follows the column headers row

            while (nextSibling != null)
            {
                if (nextSibling.NodeType == System.Xml.XmlNodeType.Element)
                {
                    bool rowProcessed = processOneRowLayoutTable(nextSibling);
                    if(!rowProcessed)
                    {
                        Console.WriteLine("ERROR - MORE trouble");
                        return false;
                    }
                    nextSibling = (XElement)nextSibling.NextNode;
                }
                else
                {
                    Console.WriteLine("ERROR - MORE trouble");
                    return false;
                }
            }
            return true;
        }

        Operation currentOp = null;
        Phase currentPhase = null;
        private bool processOneRowLayoutTable(XElement rowElem)
        {
            var cells = rowElem.Elements(w + "tc");
            if (cells.Count() != Constants.ElancoDocConstants.NumCellsInLayoutTableRow)
            {
                Console.WriteLine("ERROR - parsing layout table row, num cells: {0:D}", cells.Count());
                return false;
            }

            string trimmedContent;
            int columnCounter = 1;
            
            foreach (var cell in cells)
            {
                trimmedContent = XUtils.WordProcessingMLUtils.getParagraphTextFromCell(cell);

                if (columnCounter == 1)
                {
                    if (trimmedContent.Length > 0)
                    {
                        Operation newOp = new Operation(trimmedContent);
                        operationList.Add(newOp);
                        currentOp = newOp;
                    }
                }
                else if (columnCounter == 2)
                {
                    if (trimmedContent.Length > 0)
                    {
                        Phase newPhase = new Phase(trimmedContent);
                        currentOp.addPhase(newPhase);
                        currentPhase = newPhase;
                    }
                }
                else if (columnCounter == 3)
                {
                    if (trimmedContent.Length > 0)
                    {
                        if (trimmedContent.Substring(0, Constants.ElancoDocConstants.ProcessVarSkipName.Length)
                            .Equals(Constants.ElancoDocConstants.ProcessVarSkipName))
                        {
                            // don't process - (this is admittedly hackish)
                        }
                        else
                        {
                            // here is where we can add a reference for later adding a Units field to ProcessVar
                            ProcessVar newProcessVar = new ProcessVar(trimmedContent);
                            currentPhase.addProcessVar(newProcessVar);
                        }
                        
                    }
                }
                else
                {
                    // at this time we don't care about other columns
                    // TODO: (maybe) grab the units and associate them
                    //   with the Process variables
                    break;
                }

                columnCounter++;
            }

            return true;
        }


        public bool parseCharacteristicsSection(XElement startNode)
        {
            // check if number of level 3 headers match the number of tables

            var tables = XUtils.WordProcessingMLUtils.getChildrenOfType(startNode, "tbl");
      //      XUtils.WordProcessingMLUtils.DebugWriteCollectionLocalNames(tables, 
      //          "DebugTableNames.xml");

            var parasWithHeading3 = XUtils.WordProcessingMLUtils.getNonEmptyParagraphsOfHeaderType(startNode, 
                Constants.WordProcessingMLDefines.HeadingLevel.Heading3);
      //      XUtils.WordProcessingMLUtils.DebugWriteCollectionLocalNames(parasWithHeading3, 
      //          "DebugHeading3Names.xml");

            int numTables;
            Console.WriteLine("DBG- num Heading3: {0:D}", parasWithHeading3.Count());
            Console.WriteLine("DBG- num Tables: {0:D}", tables.Count());
                      
            if (tables == null || parasWithHeading3 == null)
            {
                Console.WriteLine("ERROR - in parseCharacteristicsSection, missing tables or headers");
                return false;
            }
            else
            {
                numTables = tables.Count();
                if (numTables != parasWithHeading3.Count())
                {
                    Console.WriteLine("ERROR - in parseCharacteristicsSection, num tables: {0:D}, num hdrs: {1:D}",
                    tables.Count(), parasWithHeading3.Count());
                    Console.WriteLine("more tables than headers usually indicates split tables!");
                    return false;
                }
            }

            for (int i = 0; i < numTables; i++)
            {
                XElement hdrNode = parasWithHeading3.ElementAt(i);
                XElement tblNode = tables.ElementAt(i);
                if (processScriptTable(hdrNode, tblNode) == false)
                {
                    return false;
                }
            }           

            return true;
        }

        bool processScriptTable(XElement hdrNode, XElement tblNode)
        {
            string hdrTitle = XUtils.WordProcessingMLUtils.getTextFromParagraphNode(hdrNode);
            bool success = hdrTitle.Length > 0  ? true : false;
            //Console.WriteLine("len: {0:D}", hdrTitle.Length);
            if (success)
            {
                string[] parts = hdrTitle.Split('.');
                // pass the data to the appropriate operation
                // there is one 'BATCH.END_TIME' process that belongs to no operation
                // therefore we need to hack a test for it.

                bool foundOp = false;
                foreach (var op in operationList)
                {
                    if (parts.Count() != 3)
                    {
                        break;
                    }
                    if (parts[0].Equals(op.Name))
                    {
                        foundOp = true;
                        op.addAspenScript(parts, tblNode);
                        break;
                    }
                }

                if(!foundOp)
                {
                    if (parts[0].Contains(Constants.ElancoDocConstants.BatchEndTimeString.ToUpper()))
                    {
                        // TODO: add it to special list at BatchUnit level
                        Console.WriteLine("TODO: add it to special list at BatchUnit level");
                    }
                    else
                    {
                        Console.WriteLine("ERROR - malformed script name: {0}", parts[0]);
                        success = false;
                    }
                }
               
            }
            return success;
        }
            
        /// <summary>
        /// return XElement of first row with matching number of cells
        /// </summary>
        /// <param name="tableRootNode"></param>
        /// <param name="numCells"></param>
        /// <returns></returns>
        private XElement findTableRowWithNumCells(XElement tableRootNode, int numCells)
        {
            var query = from x in tableRootNode.Elements(w + "tr")
                        select x;

            foreach (var tr in query)
            {
                var cells = tr.Elements(w + "tc");
                Console.WriteLine("num cells: {0:D}", cells.Count());   

                if (cells.Count() == numCells)
                {
                    return tr;
                }
            }

            return null;
        }

        public void parseReportLayoutSim(XElement reportTableRootNode)
        {

        }
    }
}
