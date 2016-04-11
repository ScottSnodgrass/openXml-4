using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace ElancoPimsDdsParser.XUtils
{
    class WordProcessingMLUtils
    {
        public static XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        
        public static string getParagraphTextFromCell(XElement para)
        {
            XNamespace w = Constants.XmlSchemas.wordmlNamespace;
            //Console.WriteLine("in getText from cell, parent name: {0}", para.Name.LocalName);
            var query1 = from x in para.Elements(w + "p")
                        select x;         
            return (getTextFromMultipleParagraphs(query1));            
        }
        /*
         * Each cell can have N paragraphs.
         * Each paragraph can have N runs.
         * Each run will have 1 at most 'text' node.
         * So the question is what do we want to handle and
         * how do we want to add newlines.
         * Currently I'm only adding newlines to multiple paragraphs
         * and not to the last paragraph text. And concatenating all
         * text from within a paragraph (all of its 'run' nodes,
         * meaning all of the 'text' nodes of the run nodes).
         */
        public static string getTextFromMultipleParagraphs(IEnumerable<XElement> paras)
        {
            StringBuilder sb = new StringBuilder();
            int numParas = paras.Count();


            for (int i = 0; i < numParas; i++)
            {
                var query = from x in paras.ElementAt(i).Descendants(w + "t")
                            select x;
                foreach (var tNode in query)
                {
                    sb.Append(tNode.Value);
                }

                // if another paragraph needs to be processed
                // then add a newline. However the next paragraph
                // may have no text...
                if (i < numParas - 1)
                {
                    sb.AppendLine();
                }
            }

            //       if (trimIt)
            {
                return sb.ToString().Trim();
            }
    //        return sb.ToString();
        }

        public static string getTextFromParagraphNode(XElement para)
        {
            List<XElement> wrapper = new List<XElement>();
            wrapper.Add(para);
            return getTextFromMultipleParagraphs(wrapper);
        }

        public static XNode getXNodeCopy(XNode org)
        {
            if (org.NodeType == XmlNodeType.Comment)
            {
                return new XComment((XComment)org);
            }

            Console.WriteLine("ERROR! Unhandled XNode type: {0}", org.NodeType.ToString());
            return null;
        }

        public static XElement CustomCopyElement_ver1(XElement element)
        {
            return new XElement(element.Name,
                element.Attributes().Where(a => !a.IsNamespaceDeclaration),
                element.Nodes().Select(n =>
                {
                    if (n is XComment)
                        return null;
                    XElement e = n as XElement;
                    if (e != null)
                        return CustomCopyElement_ver1(e);
                    return n;
                }
                )
            );
        }

        public static XElement findNextParagraphHeaderAtLevel(XElement startNode,
            Constants.WordProcessingMLDefines.HeadingLevel headerLevel)
        {
            var restOfElems = startNode.ElementsAfterSelf();

            foreach (var elem in restOfElems)
            {
                var styleNodes = elem.Elements(w + "pPr")
                            .Elements(w + "pStyle");
                var checkSyleNode =
                    from styleNode in styleNodes
                    where styleNode.Attribute(w + "val").Value.Trim().
                    Equals(Constants.WordProcessingMLDefines.Headings[headerLevel])
                    select styleNode;

                if (checkSyleNode.Count() > 0)
                {
                    return elem;
                }
            }

            return null;
        }

        public static IEnumerable<XElement> getChildrenOfType(XElement root, string t)
        {
            return (from el in root.Elements()
                         where el.Name.LocalName.Equals(t)
                         select el);
        }

        public static IEnumerable<XElement> getParagraphsOfHeaderType(XElement root,
            Constants.WordProcessingMLDefines.HeadingLevel headerLevel)
        {

            var allParas = 
                from el in root.Elements()
                where el.Name.LocalName.Equals("p")
                select el;

            var hdrParas =
                from el in allParas.Elements(w + "pPr").Elements(w + "pStyle")
                where el.Attribute(w + "val").Value.Trim().
                    Equals(Constants.WordProcessingMLDefines.Headings[headerLevel])
                select el.Parent.Parent; // we need the paragraph node, not its grandchild style property node

            return hdrParas;
        }

        public static IEnumerable<XElement> getNonEmptyParagraphsOfHeaderType(XElement root,
            Constants.WordProcessingMLDefines.HeadingLevel headerLevel)
        {
            var parasWithHeading = WordProcessingMLUtils.getParagraphsOfHeaderType(root,
                headerLevel);
            
            var query =
              from el in parasWithHeading
              where el.Descendants(w + "t").Any()
              select el;

            return query;
        }

        public static XElement CustomCopyElement_ver2(XElement element)
        {
            return new XElement(element.Name,
                //element.Attributes().Where(a => !a.IsNamespaceDeclaration),
                element.Attributes(),
                element.Nodes().Select(n =>
                {
                    if (n is XComment)
                    {
                        return null;
                    }
                    XElement e = n as XElement;
                    
                    if (e != null)
                    {
                        // Although it breaks my syntax model of always using braces,
                        // I did not include them in the following IF statements because
                        // there are so many and I wanted it cleaner looking.

                        // remove extraneous noise
                        if (e.Name.Equals(w + "pageBreakBefore"))
                            return null;
                        if (e.Name.Equals(w + "bookmarkStart"))
                            return null;
                        if (e.Name.Equals(w + "bookmarkEnd"))
                            return null;
                        if (e.Name.Equals(w + "rPr"))
                            return null;

                        if (e.Name.Equals(w + "tblGrid"))
                            return null;
                        if (e.Name.Equals(w + "tblPrChange"))
                            return null;
                        if (e.Name.Equals(w + "tblBorders"))
                            return null;
                        if (e.Name.Equals(w + "trPrChange"))
                            return null;
                        if (e.Name.Equals(w + "tcPrChange"))
                            return null;

                        if (e.Name.Equals(w + "tblPrEx"))  // don't know what this is
                            return null;

                        if (e.Name.Equals(w + "delText"))
                            return null;

                        // Console.WriteLine("node-name: {0}", e.Name);
                        //  Console.WriteLine("compare>>: {0}", testnamspace + "pR");
                        return CustomCopyElement_ver2(e);
                    }
                    return n;
                }
                )
            );
        }

        public static void DebugWriteXElementSiblings(XElement root, string outfileName)
        {
            string outputFileName = @"C:\scratchpad\MS_Office_related\ElancoPimsDdsParser\" + outfileName;
            using (StreamWriter outputFile = new StreamWriter(outputFileName))
            {
                outputFile.WriteLine("<!-- Tree Debugging -->");
                outputFile.WriteLine("<!-- ================================================================== -->");
                outputFile.WriteLine(root);                
                var followingSiblings = root.ElementsAfterSelf();
                foreach (var sib in followingSiblings)
                {
                    outputFile.WriteLine(sib);
                }
            }
        }
        public static void DebugWriteCollectionLocalNames(IEnumerable<XElement> collection, string outfileName)
        {
            string outputFileName = @"C:\scratchpad\MS_Office_related\ElancoPimsDdsParser\" + outfileName;
            using (StreamWriter outputFile = new StreamWriter(outputFileName))
            {
                outputFile.WriteLine("<!-- Listing -->");
                outputFile.WriteLine("<!-- ================================================================== -->");
                foreach(var elem in collection)
                {
                    if (elem.Name.LocalName.Equals("p"))
                    {
                        outputFile.WriteLine(getTextFromParagraphNode(elem));
                    }
                    else
                    {
                        outputFile.WriteLine(elem.Name.LocalName);
                    }
                }
            }
        }

        private static Cell createTextCell(int columnIndex, int rowIndex, object cellValue)
        {
            // https://code.msdn.microsoft.com/How-to-convert-Word-table-0cb4c9c3
            Cell cell = new Cell();

            cell.DataType = CellValues.InlineString;
        //    cell.CellReference = getColumnName(columnIndex) + rowIndex;

            InlineString inlineString = new InlineString();
            DocumentFormat.OpenXml.Spreadsheet.Text t = new DocumentFormat.OpenXml.Spreadsheet.Text();

            t.Text = cellValue.ToString();
            inlineString.AppendChild(t);
            cell.AppendChild(inlineString);

            return cell;
        }
        /**************************/
        // TODO: refactor using Eric White's code:
        //  https://msdn.microsoft.com/en-us/library/office/ff686712(v=office.14).aspx#Mastering_Retrieving
        //    public static string StringConcatenate(this IEnumerable<string> source)
        //    {
        //        StringBuilder sb = new StringBuilder();
        //        foreach (string s in source)
        //            sb.Append(s);
        //        return sb.ToString();
        //    }
        //    public static IEnumerable<XElement> LogicalChildrenContent(this IEnumerable<XElement> source)
        //    {
        //        foreach (XElement e1 in source)
        //            foreach (XElement e2 in e1.LogicalChildrenContent())
        //                yield return e2;
        //    }

        //    public static IEnumerable<XElement> LogicalChildrenContent(this XElement element,
        //XName name)
        //    {
        //        return element.LogicalChildrenContent().Where(e => e.Name == name);
        //    }

        //    public static IEnumerable<XElement> LogicalChildrenContent(
        //        this IEnumerable<XElement> source, XName name)
        //    {
        //        foreach (XElement e1 in source)
        //            foreach (XElement e2 in e1.LogicalChildrenContent(name))
        //                yield return e2;
        //    }
        /**************************/
    }
}
