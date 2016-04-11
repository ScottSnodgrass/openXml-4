using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace ElancoPimsDdsParser
{
    class DDSParser
    {        
        XNamespace w = Constants.XmlSchemas.wordmlNamespace;          
        XNamespace w14 = Constants.XmlSchemas.wordmlNamespace2010;

        BatchUnit batchUnit;
        XDocument xDocument;
        bool wordDocLoaded;

       // ExcelGenerator excelGenerator = new ExcelGenerator();
        ExcelGenerator excelGenerator;

        #region Properties
        // these properties can be adjusted with the config file
        private string characteristicsHeader = "Characteristics Header";

        public string CharacteristicsHeader
        {
            get
            {
                return characteristicsHeader;
            }
            protected set
            {
                characteristicsHeader = value;
            }
        }

        private string characteristicsHeader_SIM = "Characteristics Header - SIM";

        public string CharacteristicsHeader_SIM
        {
            get
            {
                return characteristicsHeader_SIM;
            }
            protected set
            {
                characteristicsHeader_SIM = value;
            }
        }

        private string operationReportLayoutHeader = "Operation Report Layout Header";

        public string OperationReportLayoutHeader
        {
            get
            {
                return operationReportLayoutHeader;
            }
            protected set
            {
                operationReportLayoutHeader = value;
            }
        }

        private string operationReportLayoutHeader_SIM = "Operation Report Layout Header - SIM";

        public string OperationReportLayoutHeader_SIM
        {
            get
            {
                return operationReportLayoutHeader_SIM;
            }
            protected set
            {
                operationReportLayoutHeader_SIM = value;
            }
        }

        #endregion - End Properties

        public int MyProperty { get; set; }

        public DDSParser()
        {
            initializeConfigsFromFile();

            // TODO: set batchUnit name from config file or command line
            batchUnit = new BatchUnit("7060");
        }

        public bool processWordDoc(string filename)
        {
            bool success = openWordDocFile(filename);

            if (success)
            {
             //   XUtils.WordProcessingMLUtils.DebugWriteXElementSiblings(xDocument.Root, "DebugTreeOrig.xml");
                wordDocLoaded = parseDoc();
            }

            return wordDocLoaded;
        }

        public bool buildExcelWorkbook(string filename)
        {
            bool success = false;

            if (wordDocLoaded)
            {
                excelGenerator = new ExcelGenerator(batchUnit);
                success = excelGenerator.generateElancoExcel(filename);
            }
            else
            {
                Console.WriteLine("No Doc loaded in memory!");
            }

            return success;
        }



        /// <summary>
        /// Some docs will not have all 4 headings.
        /// TODO: revisit this method's design.
        /// </summary>
        /// <returns></returns>
        private bool parseDoc()
        {
            bool success = processReportLayoutTable();

         //   success = extractReportLayoutSimTable();

            if (success)
            {
                success = processScriptTables();
            }
            

        //    success = extractScriptTablesSim();

            return success;
        }

        private bool processReportLayoutTable()
        {
            XElement firstNode = findNodeStartPosition(OperationReportLayoutHeader,
                Constants.WordProcessingMLDefines.AttributeHeading1);
            if (firstNode == null)
            {
                Console.WriteLine("ERROR! Did not find header for {0}!", this.OperationReportLayoutHeader);
                return false;
            }
            XElement lastNode = findPrecursorOfNextMatchingSibling(firstNode, Constants.WordProcessingMLDefines.HeadingLevel.Heading1);

            if (lastNode == null)
            {
                return false;
            }   

            XElement copiedNodeSegmentRoot = copySegmentElements(firstNode, lastNode);

          //  XUtils.WordProcessingMLUtils.DebugWriteXElementSiblings(copiedNodeSegmentRoot, "DebugReportLayoutSegment.xml");

            return parseReportLayoutSection(false, copiedNodeSegmentRoot);
        }
        private bool processScriptTables()
        {
            XElement firstNode = findNodeStartPosition(this.CharacteristicsHeader,
                Constants.WordProcessingMLDefines.AttributeHeading2);
            if (firstNode == null)
            {
                Console.WriteLine("ERROR! Did not find header for {0}!", this.CharacteristicsHeader);
                return false;
            }

            XElement lastNode = findPrecursorOfNextMatchingSibling(firstNode, Constants.WordProcessingMLDefines.HeadingLevel.Heading1);

            if (lastNode == null)
            {
                return false;
            }
            XElement copiedNodeSegmentRoot = copySegmentElements(firstNode, lastNode);

          //  XUtils.WordProcessingMLUtils.DebugWriteXElementSiblings(copiedNodeSegmentRoot, "DebugCharacteristicsSegment.xml");

            return parseCharacteristicsSection(false, copiedNodeSegmentRoot);
        }

        bool parseReportLayoutSection(bool isSim, XElement segmentRoot)
        {
            XElement reportTableRootNode = findReportTableElement(segmentRoot);
            if (reportTableRootNode == null)
            {
                return false;
            }
            
            return batchUnit.parseReportLayout(reportTableRootNode);

        }

        bool parseCharacteristicsSection(bool isSim, XElement segmentRoot)
        {

            return batchUnit.parseCharacteristicsSection(segmentRoot);
        }


        XElement findReportTableElement(XElement tableNodes)
        {
            XElement reportTableRootNode = null;
           
            Console.WriteLine("tableNodeStart name: {0}", tableNodes.Name);
            Console.WriteLine("tableNodeStart local name: {0}", tableNodes.Name.LocalName);
            if (tableNodes.FirstNode.NodeType == XmlNodeType.Element)
            {
                XElement child = (XElement)tableNodes.FirstNode;
                Console.WriteLine("child name: {0}", child.Name);
                Console.WriteLine("child local name: {0}", child.Name.LocalName);
            }
            else
            {
                Console.WriteLine("child node is NOT an XElement! - ERROR");
            }
            //Console.WriteLine("NAME: {0}", ((XElement)(tableNodeStart.FirstNode)).Name);

            var allP = from x in tableNodes.Elements()
                    where x.Name.LocalName == "p"
                       select x;
            Console.WriteLine("p count: {0:D}", allP.Count());

            var allTbl = from x in tableNodes.Elements()
                       where x.Name.LocalName == "tbl"
                       select x;
            Console.WriteLine("tbl count: {0:D}", allTbl.Count());

            foreach (XElement tbl in allTbl)
            {
                var numRows = from x in tbl.Elements()
                              where x.Name.LocalName == "tr"
                              select x;
                Console.WriteLine("num rows: {0:D}", numRows.Count());

                // following test could be... better but will work for now
                if (numRows.Count() > 5)
                {
                    reportTableRootNode = tbl;
                }
            }

            return reportTableRootNode;
        }


        /// <summary>
        /// TODO: merge with above method during refactoring cycle.
        /// </summary>
        /// <returns></returns>
        private bool extractReportLayoutSimTable()
        {
            XElement firstNode = findNodeStartPosition(this.OperationReportLayoutHeader_SIM,
                Constants.WordProcessingMLDefines.AttributeHeading1);
            if (firstNode == null)
            {
                Console.WriteLine("ERROR! Did not find header for {0}!", this.OperationReportLayoutHeader_SIM);
                return false;
            }

            return true;
        }

        
        private XElement copySegmentElements(XElement startNode, XElement lastNode)
        {
            // make a new "root" node
            XElement newRoot = new XElement("newroot");
            

            // copy the input element as the first child of new root
            XElement firstCopyChild = XUtils.WordProcessingMLUtils.CustomCopyElement_ver2(startNode);
            //XUtils.WordProcessingMLUtils.DebugWriteXElementSiblings(firstCopyChild, "DebugTreeCopy.xml");

            newRoot.Add(firstCopyChild);

            XNode copyRunner = firstCopyChild;
            XNode nextSibling = startNode.NextNode; // nextSibling starts at top child's sibling

            XElement nextSiblingCopyXElem = null;
            XNode nextSiblingCopyXNode = null;
            do
            {                
                if (nextSibling == null)
                {
                    continue;
                }

                if (nextSibling.NodeType != XmlNodeType.Element)
                {
                    nextSiblingCopyXNode = XUtils.WordProcessingMLUtils.getXNodeCopy(nextSibling);
                    copyRunner.AddAfterSelf(nextSiblingCopyXNode);
                }
                else
                {
                    nextSiblingCopyXElem = XUtils.WordProcessingMLUtils.
                        CustomCopyElement_ver2((XElement)nextSibling);
                    copyRunner.AddAfterSelf(nextSiblingCopyXElem);
                }

                if (nextSibling == lastNode)
                {
                    break;
                }
                copyRunner = copyRunner.NextNode; // adjust copy runner
                nextSibling = nextSibling.NextNode; // adjust original runner
            }
            while (nextSibling != null) ;

            return newRoot;
        }



        /// <summary>
        /// TODO: merge with above method during refactoring cycle.
        /// </summary>
        /// <returns></returns>
        private bool extractScriptTablesSim()
        {
            XElement firstNode = findNodeStartPosition(this.CharacteristicsHeader_SIM,
                            Constants.WordProcessingMLDefines.AttributeHeading2);

            if (firstNode == null)
            {
                Console.WriteLine("ERROR! Did not find header for {0}!", this.CharacteristicsHeader_SIM);
                return false;
            }
            return true;
        }

        /// <summary>
        /// Finds first XElement within XDocument with matching paragraph
        /// title and style of input parameters
        /// </summary>
        /// <param name="headerTitle"></param>
        /// <param name="headerLevel"></param>
        /// <returns></returns>
        private XElement findNodeStartPosition(string headerTitle, string headerLevel)
        {
            
            var paragraphs =
                from para in this.xDocument
                             .Root
                             .Element(w + "body")
                             .Descendants(w + "p")

                let styleNodes = para
                                .Elements(w + "pPr")
                                .Elements(w + "pStyle")

                from styleNode in styleNodes

                let sty = styleNode
                where sty.Attribute(w + "val").Value.Equals(headerLevel)

                select sty;

            foreach (var paraStyleNode in paragraphs)
            {
                //Console.WriteLine("Name: {0}", paraStyleNode.Name.LocalName);
                var grandParentNode = paraStyleNode.Parent.Parent;
                //Console.WriteLine("gp Node ID: {0}", grandParentNode.FirstAttribute.Value);

                //var theRunNodes = grandParentNode.Elements(w + "r");
                // some of the runs are lower down the heirarchy than the child
                // e.g. w:ins
                var theRunNodes = grandParentNode.Descendants(w + "r");

                StringBuilder sb = new StringBuilder();
                foreach (var runNode in theRunNodes)
                {
                    XElement theTextNode = null;
                    if (runNode != null)
                    {
                        theTextNode = runNode.Elements(w + "t").FirstOrDefault();
                    }
                    if (theTextNode != null)
                    {
                        sb.Append(theTextNode.Value);
                    }
                }
                if (sb.Length > 0)
                {
                    Console.WriteLine("text: {0}", sb);
                    if (sb.ToString().Trim().Equals(headerTitle))
                    {
                        return grandParentNode;
                    }
                }
            }

            return null;
        }
        
        /// <summary>
        /// Finds the following sibling of type paragraph with a header level
        /// matching the input parameter.
        /// Returns the node that is the previous sibling of the matching
        /// node.
        /// Returns null on the error condition where there is no sibling
        /// nodes after the input Element.
        /// </summary>
        /// <param name="startElement"></param>
        /// <param name="headerLevel"></param>
        /// <returns></returns>
        private XElement findPrecursorOfNextMatchingSibling(XElement startElement,
            Constants.WordProcessingMLDefines.HeadingLevel headerLevel)
        {
            var elem = XUtils.WordProcessingMLUtils.findNextParagraphHeaderAtLevel(startElement, headerLevel);
            if (elem == null)
            {
                return null;
            }
            var prev = elem.PreviousNode;
            return (XElement)prev;
        } // end of findNextMatchingSiblingV2

        private bool openWordDocFile(string filename)
        {

            try
            {
                using (Package wdPackage = Package.Open(filename, FileMode.Open, FileAccess.Read))
                {
                    PackageRelationship docPackageRelationship =
                      wdPackage.GetRelationshipsByType(Constants.XmlSchemas.documentRelationshipType).FirstOrDefault();
                    if (docPackageRelationship != null)
                    {
                        Uri documentUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative),
                          docPackageRelationship.TargetUri);
                        PackagePart documentPart = wdPackage.GetPart(documentUri);

                        //  Load the document XML in the part into an XDocument instance.
                        xDocument = XDocument.Load(XmlReader.Create(documentPart.GetStream()));

                        Console.WriteLine("TargetUri:{0}", docPackageRelationship.TargetUri);
                        Console.WriteLine("==================================================================");
                        //       Console.WriteLine(xdoc.Root);  // too long/large
                        Console.WriteLine();

                        //  Find the styles part. There will only be one.
                        PackageRelationship styleRelation =
                          documentPart.GetRelationshipsByType(Constants.XmlSchemas.stylesRelationshipType).FirstOrDefault();
                        if (styleRelation != null)
                        {
                            Uri styleUri = PackUriHelper.ResolvePartUri(documentUri, styleRelation.TargetUri);
                            PackagePart stylePart = wdPackage.GetPart(styleUri);

                            //  Load the style XML in the part into an XDocument instance.
                            XDocument styleDoc = XDocument.Load(XmlReader.Create(stylePart.GetStream()));

                            // testFunc4(styleDoc, w, xdoc, filepath);
                            return true;
                        }

                        return false;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception caught: {0}",e.Message);
                return false;
            }
            return false;
        } // end of openAndParseFile

        #region Excel generation routines

        #endregion - End - Excel generation

        List<Tuple<string, string>> propertiesList;
        private void initializeConfigsFromFile()
        {
            string filename = XmlSerializerHelper.XmlConfigDirectory + "\\" + Properties.Settings.Default.AppPropertiesFile;
            if (File.Exists(filename))
            {
                this.propertiesList = (XmlSerializerHelper.DeserializeTuples(filename));
            }
            else
            {
                // TODO: handle error
                Console.WriteLine("Error opening file");
                return;
            }


            foreach (var entry in this.propertiesList)
            {
                string item1Trimmed = entry.Item1.Trim();
                if (item1Trimmed.Equals(Constants.ElancoDocConstants.CharacteristicsHeader))
                {
                    CharacteristicsHeader = entry.Item2;
                }
                else if (item1Trimmed.Equals(Constants.ElancoDocConstants.CharacteristicsHeader_SIM))
                {
                    CharacteristicsHeader_SIM = entry.Item2;
                }
                else if (item1Trimmed.Equals(Constants.ElancoDocConstants.OperationReportLayoutHeader))
                {
                    OperationReportLayoutHeader = entry.Item2;
                }
                else if (item1Trimmed.Equals(Constants.ElancoDocConstants.OperationReportLayoutHeader_SIM))
                {
                    OperationReportLayoutHeader_SIM = entry.Item2;
                }
            }

            //if (!dataDirectoryWasSet)
            //{
                //setDataDirectory(Properties.Settings.Default.DataOutputDirectory);
            //}
        }

        // kept for the RESET reference only
        private void setDataDirectory(string directory)
        {
            //Properties.Settings.Default.Reset(); /* USED ONLY TO RESET - COMMENT OUT AFTER DEBUGGING */
            /* create output dir for data and video files */
            //System.IO.FileInfo fileInfo = new System.IO.FileInfo(directory);
            //fileInfo.Directory.Create(); // If the directory already exists, this method does nothing.

            //Properties.Settings.Default.DataOutputDirectory = directory;
            //Properties.Settings.Default.Save();

            //string logFilename = MyExtensions.appendTimeStampToFilename(
            //    directory + "\\" +
            //    Properties.Settings.Default.LogFile);

            //Logger.self.setLogFileName(logFilename);

            ////Debug.WriteLine("LaunchMain - starting logger. thread: {0:D} ", Thread.CurrentThread.ManagedThreadId);
            //Logger.self.Log(String.Format("LaunchMain - starting logger. thread: {0:D} ",
            //    System.Threading.Thread.CurrentThread.ManagedThreadId));
        }
    } // end of class DDSParser

    
}
