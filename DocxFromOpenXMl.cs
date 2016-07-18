using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ParsingHTML;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace UnitTest
{
    public class DocxFromOpenXMl
    {

        public  void CreateWordfromPackage()
        {
            string path = @"C:\VimalKumar\test\OfficeApp1\ContentControlAppWeb\OpenXML\Simple.xml";

            XmlDocument document = new XmlDocument();
            document.Load(path);
            var nav = document.CreateNavigator();
            XmlNamespaceManager xnsManager = getNamespace();
            XmlNodeList xmlNodeLst = document.SelectSingleNode("//pkg:package/pkg:part/pkg:xmlData", xnsManager).ChildNodes;  // relationship node
            foreach (XmlNode node in xmlNodeLst[0].ChildNodes)
            {
                string Xpath = "//pkg:package/pkg:part[@pkg:name='/" + node.Attributes["Target"].Value + "']";
                /*XPathExpression expr = nav.Compile(Xpath);
                expr.SetContext(xnsManager);
                XPathNodeIterator rnode = nav.Select(expr);*/

                XmlNode appnode = document.SelectSingleNode(Xpath, xnsManager);
            }

            // package.AddNewPart<OpenXmlPart>();

            // OpenXmlPart openXMLPart = package.AddPart<OpenXmlPart>(;
            /*  using (Stream stream = mainPart.GetStream())
              {

                 stream.Write(fstream, 0, fstream.Length);
             }*/

            //byte[] fstream = Encoding.ASCII.GetBytes( document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/document.xml']/pkg:xmlData", xnsManager).InnerXml);

            using (MemoryStream zipStream = new MemoryStream())
            {
                using (WordprocessingDocument package = WordprocessingDocument.Create(zipStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
                {
                    MainDocumentPart mainPart = package.AddMainDocumentPart();
                    CoreFilePropertiesPart coreFileProperty = package.AddCoreFilePropertiesPart();
                    ExtendedFilePropertiesPart fileProperties = package.AddExtendedFilePropertiesPart();

                    AddDocumentToMainDocumentPart(fileProperties, document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/docProps/app.xml']/pkg:xmlData", xnsManager).InnerXml, "rId3");
                    AddDocumentToMainDocumentPart(coreFileProperty, document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/docProps/core.xml']/pkg:xmlData", xnsManager).InnerXml, "rId2");
                    AddDocumentToMainDocumentPart(mainPart, document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/document.xml']/pkg:xmlData", xnsManager).InnerXml, "rId1");



                    AddSettingsToMainDocumentPart(mainPart, mainPart.AddNewPart<DocumentSettingsPart>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/settings.xml']/pkg:xmlData", xnsManager).InnerXml, "rId3");

                    AddSettingsToMainDocumentPart(mainPart, mainPart.AddNewPart<HeaderPart>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/header2.xml']/pkg:xmlData", xnsManager).InnerXml, "rId8");
                    AddSettingsToMainDocumentPart(mainPart, mainPart.AddNewPart<HeaderPart>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/header1.xml']/pkg:xmlData", xnsManager).InnerXml, "rId7");
                    AddSettingsToMainDocumentPart(mainPart, mainPart.AddNewPart<FontTablePart>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/fontTable.xml']/pkg:xmlData", xnsManager).InnerXml, "rId13");
                    AddSettingsToMainDocumentPart(mainPart, mainPart.AddNewPart<FooterPart>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/footer3.xml']/pkg:xmlData", xnsManager).InnerXml, "rId12");
                    AddSettingsToMainDocumentPart(mainPart, mainPart.AddNewPart<StyleDefinitionsPart>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/styles.xml']/pkg:xmlData", xnsManager).InnerXml, "rId2");

                    AddSettingsToMainDocumentPart(mainPart, mainPart.AddNewPart<EndnotesPart>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/endnotes.xml']/pkg:xmlData", xnsManager).InnerXml, "rId6");
                    AddSettingsToMainDocumentPart(mainPart, mainPart.AddNewPart<HeaderPart>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/header3.xml']/pkg:xmlData", xnsManager).InnerXml, "rId11");
                    AddSettingsToMainDocumentPart(mainPart, mainPart.AddNewPart<FootnotesPart>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/footnotes.xml']/pkg:xmlData", xnsManager).InnerXml, "rId5");

                    AddSettingsToMainDocumentPart(mainPart, mainPart.AddNewPart<ThemePart>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/theme/theme1.xml']/pkg:xmlData", xnsManager).InnerXml, "rId15");

                    AddSettingsToMainDocumentPart(mainPart, mainPart.AddNewPart<FooterPart>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/footer2.xml']/pkg:xmlData", xnsManager).InnerXml, "rId10");

                    AddSettingsToMainDocumentPart(mainPart, mainPart.AddNewPart<WebSettingsPart>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/webSettings.xml']/pkg:xmlData", xnsManager).InnerXml, "rId4");

                    AddSettingsToMainDocumentPart(mainPart, mainPart.AddNewPart<FooterPart>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/footer1.xml']/pkg:xmlData", xnsManager).InnerXml, "rId9");

                    AddSettingsToMainDocumentPart(mainPart, mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/customXml/item1.xml']/pkg:xmlData", xnsManager).InnerXml, "rId1");
                    // AddSettingsToMainDocumentPart(mainPart, mainPart.AddNewPart<CustomXmlProperties>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/customXml/itemProps1.xml']/pkg:xmlData", xnsManager).InnerXml, "rId1");


                    AddSettingsToMainDocumentPart(mainPart, mainPart.AddNewPart<GlossaryDocumentPart>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/glossary/document.xml']/pkg:xmlData", xnsManager).InnerXml, "rId14");


                    GlossaryDocumentPart glossarypart = mainPart.GlossaryDocumentPart;

                    AddSettingsToMainDocumentPart(glossarypart, glossarypart.AddNewPart<DocumentSettingsPart>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/glossary/settings.xml']/pkg:xmlData", xnsManager).InnerXml, "rId2");
                    AddSettingsToMainDocumentPart(glossarypart, glossarypart.AddNewPart<FontTablePart>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/glossary/fontTable.xml']/pkg:xmlData", xnsManager).InnerXml, "rId4");
                    AddSettingsToMainDocumentPart(glossarypart, glossarypart.AddNewPart<WebSettingsPart>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/glossary/webSettings.xml']/pkg:xmlData", xnsManager).InnerXml, "rId3");
                    AddSettingsToMainDocumentPart(glossarypart, glossarypart.AddNewPart<StyleDefinitionsPart>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/glossary/styles.xml']/pkg:xmlData", xnsManager).InnerXml, "rId1");

                    // AddSettingsToMainDocumentPart(glossarypart, glossarypart.AddNewPart<Relationship>(), document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/word/glossary/_rels/document.xml.rels']/pkg:xmlData", xnsManager).InnerXml);

                    foreach (IdPartPair partPair in package.Parts)
                    {
                        // partPair.RelationshipId= document.SelectSingleNode("//pkg:package/pkg:part[@pkg:name='/_rels/.rels']/pkg:xmlData/Relationships/Relationship", xnsManager)
                    }
                    foreach (IdPartPair partPair in glossarypart.Parts)
                    {
                    }

                    package.MainDocumentPart.Document.Save();
                    foreach (var header in package.MainDocumentPart.HeaderParts)
                        header.Header.Save();
                    foreach (var footer in package.MainDocumentPart.FooterParts)
                        footer.Footer.Save();
                    if (package.MainDocumentPart.FootnotesPart != null)
                        package.MainDocumentPart.FootnotesPart.Footnotes.Save();
                    if (package.MainDocumentPart.EndnotesPart != null)
                        package.MainDocumentPart.EndnotesPart.Endnotes.Save();
                    
                }
                zipStream.Position = 0;
                //using (WordprocessingDocument package = WordprocessingDocument.Open(zipStream, false))
                //{
                //    MainDocumentPart mainPart = package.MainDocumentPart;

                //}
                using (FileStream fileStream = new FileStream("Test2.docx",
                System.IO.FileMode.CreateNew))
                {
                    zipStream.WriteTo(fileStream);
                }
            }
        }

       
        public  void AddDocumentToMainDocumentPart(OpenXmlPart part, string innerXmlStr, string id)
        {
            MemoryStream stream = new MemoryStream();
            StreamWriter writer = new StreamWriter(stream);
            writer.Write(innerXmlStr);
            writer.Flush();
            stream.Position = 0;
            part.FeedData(stream);
            part.OpenXmlPackage.ChangeIdOfPart(part, id);
        }

        public  void AddSettingsToMainDocumentPart(MainDocumentPart part, OpenXmlPart settingsPart, string innerXmlStr, string id)
        {
            //DocumentSettingsPart settingsPart = part.AddNewPart<DocumentSettingsPart>();

            // OpenXmlPart settingsPart = part.AddNewPart<T>();
            //FileStream settingsTemplate = new FileStream("settings.xml", FileMode.Open, FileAccess.Read);
            MemoryStream stream = new MemoryStream();
            StreamWriter writer = new StreamWriter(stream);
            writer.Write(innerXmlStr);
            writer.Flush();
            stream.Position = 0;

            settingsPart.FeedData(stream);
            part.ChangeIdOfPart(settingsPart, id);
            //settingsPart.Settings.Save();
        }


        public  void AddSettingsToMainDocumentPart(GlossaryDocumentPart part, OpenXmlPart settingsPart, string innerXmlStr, string id)
        {
            //DocumentSettingsPart settingsPart = part.AddNewPart<DocumentSettingsPart>();

            // OpenXmlPart settingsPart = part.AddNewPart<T>();
            //FileStream settingsTemplate = new FileStream("settings.xml", FileMode.Open, FileAccess.Read);
            MemoryStream stream = new MemoryStream();
            StreamWriter writer = new StreamWriter(stream);
            writer.Write(innerXmlStr);
            writer.Flush();
            stream.Position = 0;

            settingsPart.FeedData(stream);
            part.ChangeIdOfPart(settingsPart, id);
            //settingsPart.Settings.Save();
        }

        public  XmlNamespaceManager getNamespace()
        {
            XmlNameTable xnt = new NameTable();
            XmlNamespaceManager xnManager = new XmlNamespaceManager(xnt);
            xnManager.AddNamespace("", "http://schemas.openxmlformats.org/package/2006/relationships");
            xnManager.AddNamespace("pkg", "http://schemas.microsoft.com/office/2006/xmlPackage");
            xnManager.AddNamespace("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            xnManager.AddNamespace("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            xnManager.AddNamespace("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            xnManager.AddNamespace("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            xnManager.AddNamespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            xnManager.AddNamespace("o", "urn:schemas-microsoft-com:office:office");




            xnManager.AddNamespace("p0", "http://schemas.openxmlformats.org/markup-compatibility/2006");


            xnManager.AddNamespace("p1", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");


            xnManager.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");


            xnManager.AddNamespace("v", "urn:schemas-microsoft-com:vml");


            xnManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");


            xnManager.AddNamespace("w10", "urn:schemas-microsoft-com:office:word");


            xnManager.AddNamespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml");


            xnManager.AddNamespace("w15", "http://schemas.microsoft.com/office/word/2012/wordml");


            xnManager.AddNamespace("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");


            xnManager.AddNamespace("wne", "http://schemas.microsoft.com/office/word/2006/wordml");


            xnManager.AddNamespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");


            xnManager.AddNamespace("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");


            xnManager.AddNamespace("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");


            xnManager.AddNamespace("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");


            xnManager.AddNamespace("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");


            xnManager.AddNamespace("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");


            return xnManager;

        }
    }
}
