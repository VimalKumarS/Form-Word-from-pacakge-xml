using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;

using System.Collections.Generic;

using System.Linq;

using System.IO;

using System.Xml.Linq;

using DocumentFormat.OpenXml.Packaging;
namespace UnitTest
{
    class listPartOOxml
    {
        private static void AddPart(HashSet<OpenXmlPart> partList, OpenXmlPart part)
        {

            if (partList.Contains(part))

                return;

            partList.Add(part);

            foreach (IdPartPair p in part.Parts)

                AddPart(partList, p.OpenXmlPart);

        }

        public static List<OpenXmlPart> GetAllParts(WordprocessingDocument doc)
        {

            // use the following so that parts are processed only once

            HashSet<OpenXmlPart> partList = new HashSet<OpenXmlPart>();

            foreach (IdPartPair p in doc.Parts)

                AddPart(partList, p.OpenXmlPart);

            return partList.OrderBy(p => p.ContentType).ThenBy(p => p.Uri.ToString()).ToList();

        }
        public static List<OpenXmlPart> GetAllParts(SpreadsheetDocument doc)
        {

            // use the following so that parts are processed only once

            HashSet<OpenXmlPart> partList = new HashSet<OpenXmlPart>();

            foreach (IdPartPair p in doc.Parts)

                AddPart(partList, p.OpenXmlPart);

            return partList.OrderBy(p => p.ContentType).ThenBy(p => p.Uri.ToString()).ToList();

        }
        public static List<OpenXmlPart> GetAllParts(PresentationDocument doc)
        {

            // use the following so that parts are processed only once

            HashSet<OpenXmlPart> partList = new HashSet<OpenXmlPart>();

            foreach (IdPartPair p in doc.Parts)

                AddPart(partList, p.OpenXmlPart);

            return partList.OrderBy(p => p.ContentType).ThenBy(p => p.Uri.ToString()).ToList();

        }

        public static void PrintParts(List<OpenXmlPart> partList)
        {

            int[] tabs = new[] { 25 };

            Console.WriteLine("{0}{1}", "URI".PadRight(tabs[0]), "Content Type");

            Console.WriteLine("{0}{1}", "===".PadRight(tabs[0]), "============");

            foreach (var p in partList)

                Console.WriteLine("{0}{1}", p.Uri.ToString().PadRight(tabs[0]), p.ContentType);

        }

        public static void MainCall()
        {
            string file = "Test.docx";

            if (!File.Exists(file))
            {

                Console.WriteLine("File ‘{0}’ doesn’t exist.", file);

                Environment.Exit(1);

            }

            FileInfo fi = new FileInfo(file);

            switch (fi.Extension.ToLower())
            {

                case ".docx":

                    using (WordprocessingDocument wp1 = WordprocessingDocument.Open(file, true))
                    {

                        List<OpenXmlPart> partList = GetAllParts(wp1);

                        PrintParts(partList);

                    }

                    break;

                case ".xlsx":

                    using (SpreadsheetDocument s1 = SpreadsheetDocument.Open(file, true))
                    {

                        List<OpenXmlPart> partList = GetAllParts(s1);

                        PrintParts(partList);

                    }

                    break;

                case ".pptx":

                    using (PresentationDocument p1 = PresentationDocument.Open(file, true))
                    {

                        List<OpenXmlPart> partList = GetAllParts(p1);

                        PrintParts(partList);

                    }

                    break;

            }


        }
    }
}
