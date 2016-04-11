using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ElancoPimsDdsParser
{
    class Constants
    {
        public static class XmlSchemas
        {
            public const string documentRelationshipType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

            public const string stylesRelationshipType =
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

            public const string wordmlNamespace =
              "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            // xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            public const string wordmlNamespace2010 = 
                "http://schemas.microsoft.com/office/word/2010/wordml";
        }

        public static class ElancoDocConstants
        {
            // the following strings MUST match the "keys" (Item1 of tuples)
            // within the file DdsParserConfig.xml
            public const string CharacteristicsHeader = "Characteristics Header";
            public const string CharacteristicsHeader_SIM = "Characteristics Header - SIM";
            public const string OperationReportLayoutHeader = "Operation Report Layout Header";
            public const string OperationReportLayoutHeader_SIM = "Operation Report Layout Header - SIM";

            public const int NumCellsInLayoutTableRow = 9;

            // within the Operation Report Layout table there are some
            // entries to demarcate 'stages'. These entries do NOT represent
            // any batch report variables.
            public const string ProcessVarSkipName = "Stage";

            // process without the Operation or Phase identifiers
            public const string BatchEndTimeString = "Batch";

            // titles in Batch Area Characteristic Processing tables
            public const string TriggerTitle = "Trigger";
            public const string VariableBindingTitle = "Variable Binding";
            public const string ScriptTitle = "Script";

 
        }

        
        public static class WordProcessingMLDefines
        {
            public enum HeadingLevel { Heading1 = 1, Heading2, Heading3};

            public static readonly Dictionary<HeadingLevel, string> Headings = new Dictionary<HeadingLevel, string>
            {
                {HeadingLevel.Heading1, "Heading1" },
                {HeadingLevel.Heading2, "Heading2" },
                {HeadingLevel.Heading3, "Heading3" }
            };

            public const string AttributeHeading1 = "Heading1";
            public const string AttributeHeading2 = "Heading2";

            // is the below overkill?

            public const string ElementP = "p";
            public const string ElementBody = "body";
            public const string ElementPStyle = "pStyle";
            public const string ElementPpr = "pPr";
            public const string AttributeVal = "val";
            
        }

    }
}
