using Apttus.XAuthor.Integration.Model;
using Apttus.XAuthor.Integration.Service.Document;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.WebExtension;
using DocumentFormat.OpenXml.Office2013.WebExtentionPane;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using ParsingHTML;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using WebAuthorIntegration;

//C:\VimalKumar\XAuthor\main\main\Apttus.XAuthor.Integration
namespace UnitTest
{
    public static class IntegrationTest
    {
        static void Main(string[] args)
        {
            //XDocument doc;
            // doc = OpenXmlToOpenFlat.OpcToFlatOpc(@"C:\VimalKumar\test\OfficeApp1\ContentControlApp\SimpleOffline.docx");
            // doc.Save(“Test.xml”, SaveOptions.DisableFormatting);
            // readOpenXml();
            //var x = Convert.ToDouble("");
            //var DateRev = DateTime.Parse("", CultureInfo.InvariantCulture);
            //checkinDoc();
            //UpdatePlacementOfTag();
            //AddWebSettingTaskPane();
            //GetBase64OfDocx();
            //TestHttpClient();
            // WordPackageXMLOnly(@"C:\VimalKumar\test\OfficeApp1\ContentControlApp\Normal1.docx");
            // CheckDocPartCategory();
            MarkDocumentAsOffline();
            /*
            UpdatePlacementOfTag();
            DateTime.Now.ToString("M/dd/yyyy h:mm:ss tt");
            Console.WriteLine(DateTime.UtcNow.ToString("s") + "Z");
           DateTime dt= new DateTime(DateTime.Parse("2016-01-21T11:11:00.000Z", CultureInfo.InstalledUICulture).Ticks);
           dt.ToString("M/dd/yyyy h:mm:ss tt");
           DateTime parsedDateTime;
           DateTime.TryParseExact("2016-01-21T11:11:00.000Z", "yyyy-MM-ddTHH:mm:ss.fffZ", null, System.Globalization.DateTimeStyles.None, out parsedDateTime);
           // parsedDateTime;
            WordPackageXML(@"C:\VimalKumar\Dropbox (Apttus)\MAC Code\template\XAuthor.dotm");
             * */
            //CreateWordfromPackage();
            string openxml = "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";
            //ParseDocumentData(openxml);
        }


        public static void ParseDocumentData(string openXml)
        {
            byte[] byteArray = Convert.FromBase64String(openXml);
            using (var memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                //We got our document on server. Yay!
                using (var wordDoc = WordprocessingDocument.Open(memoryStream, true))
                {

                    wordDoc.SaveAll();

                    using (FileStream fileStream = new FileStream("Test2.docx",
                     System.IO.FileMode.CreateNew))
                    {
                        memoryStream.WriteTo(fileStream);
                    }
                }
            }

        }


        private static void CreateWordfromPackage()
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
                XPathExpression expr = nav.Compile(Xpath);
                expr.SetContext(xnsManager);
                XPathNodeIterator rnode = nav.Select(expr);

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

                    package.SaveAll();
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


        private static void AddDocumentToMainDocumentPart(OpenXmlPart part, string innerXmlStr, string id)
        {
            MemoryStream stream = new MemoryStream();
            StreamWriter writer = new StreamWriter(stream);
            writer.Write(innerXmlStr);
            writer.Flush();
            stream.Position = 0;
            part.FeedData(stream);
            part.OpenXmlPackage.ChangeIdOfPart(part, id);
        }

        private static void AddSettingsToMainDocumentPart(MainDocumentPart part, OpenXmlPart settingsPart, string innerXmlStr, string id)
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


        private static void AddSettingsToMainDocumentPart(GlossaryDocumentPart part, OpenXmlPart settingsPart, string innerXmlStr, string id)
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

        public static XmlNamespaceManager getNamespace()
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
        private static void readOpenXml()
        {
            byte[] fstream = File.ReadAllBytes(@"C:\VimalKumar\test\OpenXMLTest.xml");

            using (MemoryStream zipStream = new MemoryStream())
            {
                using (WordprocessingDocument package = WordprocessingDocument.Create(zipStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
                {
                    MainDocumentPart mainPart = package.AddMainDocumentPart();
                    using (Stream stream = mainPart.GetStream())
                    {

                        stream.Write(fstream, 0, fstream.Length);
                    }
                }
                zipStream.Position = 0;
                using (WordprocessingDocument package = WordprocessingDocument.Open(zipStream, false))
                {
                    MainDocumentPart mainPart = package.MainDocumentPart;

                }
            }


            /*byte[] data = File.ReadAllBytes(@"C:\VimalKumar\test\OpenXMLTest.xml");
                //Encoding.UTF8.GetBytes(String.Join("\n", new string[1000].Select(s => "Something to zip.").ToArray()));
            byte[] zippedBytes;
            using (MemoryStream zipStream = new MemoryStream())
            {
                using (Package package = Package.Open(zipStream, FileMode.Create))
                {
                    PackagePart document = package.CreatePart(new Uri(@"/OpenXMLTest.docx", UriKind.Relative), "");
                    using (MemoryStream dataStream = new MemoryStream(data))
                    {
                        document.GetStream().WriteAll(dataStream);
                    }
                }
                zippedBytes = zipStream.ToArray();
            }
            File.WriteAllBytes("test.zip", zippedBytes);*/
        }

        private static void WriteAll(this Stream target, Stream source)
        {
            const int bufSize = 0x1000;
            byte[] buf = new byte[bufSize];
            int bytesRead = 0;
            while ((bytesRead = source.Read(buf, 0, bufSize)) > 0)
                target.Write(buf, 0, bytesRead);
        }

        private static void GetBase64OfDocx()
        {
            string path = @"C:\VimalKumar\test-officeapi\AJ_March28_1_Original_XA_Web_Reconcile_4_2016-03-28.docx";
            byte[] binarydata = File.ReadAllBytes(path);
            var base64 = System.Convert.ToBase64String(binarydata, 0, binarydata.Length);

            XDocument _xmlDoc = OpenXmlToOpenFlat.OpcToFlatOpc(binarydata);
        }
        public class CheckInDocument
        {
            
            public string Base64Dcoument
            {
                get;
                set;
            }

            public string CheckInFileName
            {
                get;
                set;
            }
        }
        private static void TestHttpClient1()
        {
            CheckInDocument checkin = new CheckInDocument { Base64Dcoument = "Dhananjay", CheckInFileName = "9" };
            WebClient proxy = new WebClient();
            proxy.Headers["Content-type"] = "application/json";
            MemoryStream ms = new MemoryStream();
            DataContractJsonSerializer serializerToUplaod = new DataContractJsonSerializer(typeof(CheckInDocument));
            serializerToUplaod.WriteObject(ms, checkin);
            byte[] data = proxy.UploadData("https://localhost:44307/ServiceXAuhtor.svc/api/CheckInDocumentSF", "POST", ms.ToArray());
            var stream = new MemoryStream(data);

        }


        private static async void TestHttpClient()
        {
            try
            {
                HttpClient httpClient = new HttpClient();
                CheckInDocument checkin = new CheckInDocument { Base64Dcoument = "Dhananjay", CheckInFileName = "9" };
                string postBody = JsonSerializer(checkin);
                httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                HttpResponseMessage wcfResponse = await httpClient.PostAsync("https://localhost:44307/ServiceXAuhtor.svc/api/CheckInDocumentSF", new StringContent(postBody, Encoding.UTF8, "application/json"));
                string responJsonText = await wcfResponse.Content.ReadAsStringAsync();
            }
            catch (System.Exception exp)
            {

            }
        }

        public static string JsonSerializer(CheckInDocument objectToSerialize)
        {
            if (objectToSerialize == null)
            {
                throw new ArgumentException("objectToSerialize must not be null");
            }
            MemoryStream ms = null;

            DataContractJsonSerializer serializer = new DataContractJsonSerializer(objectToSerialize.GetType());
            ms = new MemoryStream();
            serializer.WriteObject(ms, objectToSerialize);
            ms.Seek(0, SeekOrigin.Begin);
            StreamReader sr = new StreamReader(ms);
            return sr.ReadToEnd();
        } 

        private static void AddWebSettingTaskPane()
        {
            using (WordprocessingDocument package = WordprocessingDocument.Open(@"C:\VimalKumar\test-officeapi\ExistingPaneAdd.docx", true))
            {
                WebExTaskpanesPart webExTaskpanesPart1 = package.GetPartsOfType<WebExTaskpanesPart>().FirstOrDefault();

                IEnumerable<IdPartPair> parts= package.Parts;
                if (webExTaskpanesPart1 != null)
                {
                    WebExtensionPart webExtensionPart1 = webExTaskpanesPart1.GetPartsOfType<WebExtensionPart>().FirstOrDefault();
                    WebExtensionPart ApttusWebExtension = webExTaskpanesPart1.GetPartsOfType<WebExtensionPart>().Where(x => x.WebExtension.WebExtensionStoreReference.Id == "2c4c8a5c-7da3-46f1-9f49-6b262c2376f8").FirstOrDefault();
                    if (ApttusWebExtension == null)
                    {

                        WebExTaskpanesPart newWebTaskPane = package.GetPartsOfType<WebExTaskpanesPart>().FirstOrDefault();
                        GenerateWebExTaskpanesPart1Content(newWebTaskPane, "rId11");


                        WebExtensionPart webExtensionPart11 = newWebTaskPane.AddNewPart<WebExtensionPart>("rId11");
                        GenerateWebExtensionPart1Content(webExtensionPart11);
                    }
                }
                else
                {
                    WebExTaskpanesPart newWebTaskPane = package.AddNewPart<WebExTaskpanesPart>("rId12"); ;
                    GenerateWebExTaskpanesPart1Content(newWebTaskPane, "rId11");


                    WebExtensionPart webExtensionPart1 = newWebTaskPane.AddNewPart<WebExtensionPart>("rId11");
                    GenerateWebExtensionPart1Content(webExtensionPart1);
                }

                
                package.SaveAll();

            }
        }

      
        private static void GenerateWebExtensionPart1Content(WebExtensionPart webExtensionPart1)
        {

            WebExtension webExtension1 = new WebExtension() { Id = "{"+Guid.NewGuid().ToString()+"}" };
            webExtension1.AddNamespaceDeclaration("we", "http://schemas.microsoft.com/office/webextensions/webextension/2010/11");
            WebExtensionStoreReference webExtensionStoreReference1 = new WebExtensionStoreReference() { Id = "2c4c8a5c-7da3-46f1-9f49-6b262c2376f8", Version = "1.0.0.0", Store = "developer", StoreType = "Registry" };
            WebExtensionReferenceList webExtensionReferenceList1 = new WebExtensionReferenceList();
            WebExtensionPropertyBag webExtensionPropertyBag1 = new WebExtensionPropertyBag();
            WebExtensionBindingList webExtensionBindingList1 = new WebExtensionBindingList();

            Snapshot snapshot1 = new Snapshot();
            snapshot1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            webExtension1.Append(webExtensionStoreReference1);
            webExtension1.Append(webExtensionReferenceList1);
            webExtension1.Append(webExtensionPropertyBag1);
            webExtension1.Append(webExtensionBindingList1);
            webExtension1.Append(snapshot1);

            webExtensionPart1.WebExtension = webExtension1;
        }

        private static void GenerateWebExTaskpanesPart1Content(WebExTaskpanesPart part,string partID)
        {

            Taskpanes taskpanes1 = new Taskpanes();
            taskpanes1.AddNamespaceDeclaration("wetp", "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11");

            WebExtensionTaskpane webExtensionTaskpane1 = new WebExtensionTaskpane() { DockState = "", Visibility = true, Width = 350D, Row = (UInt32Value)0U };

            WebExtensionPartReference webExtensionPartReference1 = new WebExtensionPartReference() { Id = partID };
            webExtensionPartReference1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            webExtensionTaskpane1.Append(webExtensionPartReference1);

            taskpanes1.Append(webExtensionTaskpane1);

            part.Taskpanes = taskpanes1;
        }




        public static void checkinDoc()
        {

            IWebAuthor webAuthorObj = new WebAuthorCaller("00DM0000001YHTB!AQQAQOnhSJ.Kk1EVGTjSgI53Vq6diFbVCvatgBWXU2BUkqyZiLMF9egp14F0yhvUGPbeZgYg4kRCdirfekn35A6b2iPWkfsI",
                "https://apttus.cs7.visual.force.com/services/Soap/u/30.0/00DM0000001YHTB");

            OpenXmlDocCheckInService docCheckIn = new OpenXmlDocCheckInService(webAuthorObj.Session);


            string bSuccess = docCheckIn.CheckIn(new CheckInRequest()
            {
                DocumentPath = @"C:\Apache2.4\htdocs\WAC_QA\test (2).docx",
                MarkAsPrivate = false,
                RemoveWaterMark = true,
                IncludeWaterMark = false,
                SuggestedFileName = "test.docx",
                CreatePdfAttachment = false,
                SubmitForApproval = (int)(2) == 3 ? true : false,
                SaveOption = ApplicationConstants.VersionTypeEnum.Internal,
                AtleastOneClauseChangedOrDeleted = true, // Todo: Need to check for if clause changed or not
                ReconcileFields = false,
                ReconcileClauseApprovals = false,
                ReconcileClauses = false,
                ReconcileTables = false
            });
        }

        public static string WordPackageXML(string sInputDocFile)
        {
            var wu = new WordUtil.WordUtil();
            var wordDoc = wu.ReadWordDoc(sInputDocFile);
            byte[] array = Encoding.UTF8.GetBytes(wordDoc.WordOpenXML);

            //Encode to Base64 twice - Apptus metadata contain base64 metadata, and on UI div tag contain base64 metadata
            string sPropertyValue = Convert.ToBase64String(array);
            array = Encoding.UTF8.GetBytes(sPropertyValue);
            sPropertyValue = Convert.ToBase64String(array);
            wu.Close();
            return sPropertyValue;

        }

        private static void UpdatePlacementOfTag()
        {
            WhtmlToBhtml wtohtml = new WhtmlToBhtml();
            wtohtml.SContent = File.ReadAllText(@"C:\Apache2.4\htdocs\PCTestTable.html");
            wtohtml.ConvertToXHtml();
            wtohtml.DecorateCCs();
            wtohtml.UpdatePlacementOfSpanTag();
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(wtohtml.SContent);

        }

        public static void WordPackageXMLOnly(string sInputDocFile)
        {
            var wu = new WordUtil.WordUtil();
            var wordDoc = wu.ReadWordDoc(sInputDocFile);
            string xml = wordDoc.WordOpenXML;
            byte[] array = Encoding.UTF8.GetBytes(wordDoc.WordOpenXML);

            //Encode to Base64 twice - Apptus metadata contain base64 metadata, and on UI div tag contain base64 metadata
            // string sPropertyValue = Convert.ToBase64String(array);
            //array = Encoding.UTF8.GetBytes(sPropertyValue);
            //sPropertyValue = Convert.ToBase64String(array);
            wu.Close();
            // return sPropertyValue;

        }


        private static void CheckDocPartCategory()
        {
            XmlNamespaceManager m_xmlNamespaceManager = GetNamespaceManager(@"C:\VimalKumar\Dropbox (Apttus)\MAC Code\VerifyOoxml.docx");
            using (WordprocessingDocument package = WordprocessingDocument.Open(@"C:\VimalKumar\Dropbox (Apttus)\MAC Code\VerifyOoxml.docx", false))
            {
                MainDocumentPart mainPart = package.MainDocumentPart;
                //HeaderPart hpart= mainPart.HeaderParts.FirstOrDefault();
                //XmlDocument hXmlPart = GetXmlDocument(hpart);
                //XmlNodeList nodeList = hXmlPart.SelectNodes(".//w:sdt", m_xmlNamespaceManager);

                //AptObject obj = new AptObject(nodeList[0], m_xmlNamespaceManager);


                XmlDocument theXmlPart = GetXmlDocument(mainPart);
                XmlNodeList nodeList = theXmlPart.SelectNodes(".//w:sdt", m_xmlNamespaceManager);

                //  obj = new AptObject(nodeList[5], m_xmlNamespaceManager);

                foreach (var cc in package.ContentControls())
                {
                    SdtProperties policyNumberCC = cc.Descendants<SdtProperties>().FirstOrDefault();
                    //mainPart.Document.Body.Descendants<SdtBlock>().Where(r => r.SdtProperties.GetFirstChild<Tag>().Val == "Clause").SingleOrDefault();

                }
            }
        }

        private static XmlNamespaceManager GetNamespaceManager(string sWordFileName)
        {
            XmlNameTable xnt = new NameTable();
            XmlNamespaceManager xnManager = new XmlNamespaceManager(xnt);

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(sWordFileName, false))
            {
                string sXml = wordDoc.MainDocumentPart.Document.OuterXml;
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(sXml);
                foreach (XmlAttribute att in xmlDoc.DocumentElement.Attributes)
                {
                    string sPrefix = att.Name;
                    int iIndex = sPrefix.IndexOf(':');
                    if (iIndex > -1 && sPrefix.Substring(0, iIndex).Equals("xmlns"))
                    {
                        sPrefix = sPrefix.Substring(iIndex + 1);
                    }
                    string sUri = att.Value;
                    xnManager.AddNamespace(sPrefix, sUri);


                }
            }
            return xnManager;
        }
        private static XmlDocument GetXmlDocument(OpenXmlPart part)
        {
            // Load the main document part into an XDocument
            XmlDocument xmlDoc = new XmlDocument();
            using (Stream str = part.GetStream())
            using (XmlReader xr = XmlReader.Create(str))
            {
                xmlDoc.Load(xr);
            }
            return xmlDoc;
        }

        private static void MarkDocumentAsOffline()
        {
            /*using (WordprocessingDocument package = WordprocessingDocument.Open(@"C:\Apache2.4\htdocs\WAC_QA\offlienWord.docx", true))
             {

               
                DocumentProperties _exisitingProperties= Apttus.XAuthor.Integration.Service.Document.DocumentPropertyHelper.GetProperties(docProp);
                DocProp docProp = new DocProp(package); 
                DocumentProperties docProperties = new DocumentProperties();
                 docProperties.SfBusinessObjectContext = "Apttus__APTS_Agreement__c";
                 docProperties.SfObjectType = "Apttus__APTS_Agreement__c"; // 
                 docProperties.SfObjectId = "a07M0000009TM1YIAW"; // get the object id
                 // empty version info
                 string sDocId = (Guid.NewGuid()).ToString();
                 docProperties.DocumentId = sDocId;
                 Apttus.XAuthor.Integration.Service.Document.DocumentPropertyHelper.SetProperties(docProp, docProperties);

                 Dictionary<string, string> mergeInfoItems = new Dictionary<string, string>();
                 mergeInfoItems.Add("Version", string.Empty);
                 docProp.SetObjectPropertyItem(ApplicationConstants.DocumentPropertyEnum.MergeInfo, mergeInfoItems);

                 docProp.Save();
             }*/

            //Step chekcin

            IWebAuthor webAuthorObj = new WebAuthorCaller("00DM0000001YHTB!AQQAQAZTh.VmQi0W92o3WrNBLKaDjUALbX4IZeZojN4zT0L9Ijh1rPd46DISbtF54W.tL0ie3uyxIitb7RoAFORhkdddpHC5",
                    "https://apttus.cs7.visual.force.com/services/Soap/u/30.0/00DM0000001YHTB");
            //using (WordprocessingDocument package = WordprocessingDocument.Open(@"C:\Apache2.4\htdocs\WAC_QA\new.docx", true))
            //  {


            //Apttus.XAuthor.Integration.Service.Document.DocProp docProp = new Apttus.XAuthor.Integration.Service.Document.DocProp(package);
            Apttus.XAuthor.Integration.Model.DocumentProperties docProperties = new Apttus.XAuthor.Integration.Model.DocumentProperties();
            docProperties.SfBusinessObjectContext = "Apttus__APTS_Agreement__c";
            docProperties.SfObjectType = "Apttus__APTS_Agreement__c"; // 
            docProperties.SfObjectId = "a07M0000009TOYeIAO"; // get the object id

            string sDocId = (Guid.NewGuid()).ToString();
            docProperties.DocumentId = sDocId;

            OpenXmlDocCheckInService docCheckIn = new OpenXmlDocCheckInService(webAuthorObj.Session);


             string attachmentId = docCheckIn.ImportOfflineAgreement(new ImportOfflineAgreementRequest()
             {
                 DocumentPath = @"C:\Apache2.4\htdocs\WAC_QA\offline.docx",
                 MakeDocumentPrivate = false,
                 Properties = docProperties,
                 SuggestedFileName = "OffilienAgreementTest.docx",
                 OfflineMode = ApplicationConstants.ImportOfflineModeEnum.Existing,
                 IsUserPasswordDifferentThanSystemPassword = false
             });
            //  }


            /*string bSuccess = docCheckIn.CheckIn(new CheckInRequest()
            {
                DocumentPath = @"C:\Apache2.4\htdocs\WAC_QA\offlienWord.docx",
                MarkAsPrivate = false,
                RemoveWaterMark = false,
                SuggestedFileName = "offlienWord.docx",
                CreatePdfAttachment = false,
                SubmitForApproval = (int)(2) == 3 ? true : false,
                SaveOption = ApplicationConstants.VersionTypeEnum.Internal,
                AtleastOneClauseChangedOrDeleted = true, // Todo: Need to check for if clause changed or not
                ReconcileFields = true,
                ReconcileClauseApprovals = true,
                ReconcileClauses = true,
                ReconcileTables = true
            });*/

            // get attachment id

            // check out the attached document
        }
    }
}


/*
 <AptDocProperties>
 * <DocumentID>fb14b434-7877-4f94-bcf5-f41c6e093914</DocumentID>
 * <SF_OBJECT_TYPE>Apttus__APTS_Agreement__c</SF_OBJECT_TYPE>
 * <SF_BUSINESS_OBJECT_CONTEXT>Apttus__APTS_Agreement__c</SF_BUSINESS_OBJECT_CONTEXT>
 * <SF_OBJECT_ID>a07M0000009TM1YIAW</SF_OBJECT_ID>
 * <MergeInfo>
 * <Version>
 * </Version></MergeInfo>
 * <SF_OBJECT_VERSION>1.0.0</SF_OBJECT_VERSION>
 * <HasSmartItem>False</HasSmartItem>
 * <ATTACHMENT_ID>00PM0000004LVyUMAW</ATTACHMENT_ID>
 * </AptDocProperties>
 
 */

/*
 <?xml version="1.0" encoding="UTF-8"?>
 * <ActionResponse><Status>Ok</Status>
 * <BrowseAgreementsResult>
 * <SObjectName>Apttus__APTS_Agreement__c</SObjectName>
 * <SObjectId>a07M0000009TM1YIAW</SObjectId>
 * <Name>WAC661</Name>
 * <Action>Select</Action>
 * </BrowseAgreementsResult></ActionResponse>
 */