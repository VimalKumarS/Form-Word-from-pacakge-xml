using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using WordCore = Microsoft.Office.Core;
using Newtonsoft.Json;
using WebAuthorIntegration;
using System.Xml;
using System.Drawing;
using System.Drawing.Imaging;
using ParsingHTML;
using Apttus.XAuthor.Integration.Util;
using Apttus.XAuthor.Integration.Model;
using Apttus.XAuthor.Integration.Service.Document;
using System.IO.Packaging;
namespace UnitTest
{
    public class WordUtilTest
    {
        public void TestGetAllFields()
        {
            WordUtil.WordUtil wordUtil = new WordUtil.WordUtil();
            Word.Document document = wordUtil.ReadWordDoc(@"C:\Apache2.4\htdocs\DynamicItems\a11c4d57-a211-4139-ad49-eafc7b827ebc\Ram_Test_AuthorWeb_Regenerated_FX2_WA_Sample_1_Template_2015-08-26.docx");
            var customXMLParts = document.CustomXMLParts.SelectByNamespace("http://www.apttus.com/externalmetadata");
            XDocument rootElement = null;

            if (customXMLParts != null && customXMLParts.Count == 1)
            {
                rootElement = XDocument.Parse(customXMLParts[1].XML, LoadOptions.None);
            }


            wordUtil.Close();
        }

        public void checkinDoc()
        {
            IWebAuthor webAuthorObj = new WebAuthorCaller("00DM0000001YHTB!AQQAQPRchbYurOMSlGAVoTWr4TVII22IWIrMUxD7K4ZOOJwnTK7HE.3AuzfO91gdGDYsBik3EFudzniNkxGYURrpqrdBzZPJ",
                "https://apttus.cs7.visual.force.com/services/Soap/u/30.0/00DM0000001YHTB");

            WebAuthorIntegration.DocumentCheckInService docCheckIn = new DocumentCheckInService(webAuthorObj.Session);

            bool bSuccess = docCheckIn.CheckIn(new CheckInRequest()
            {
                DocumentPath = @"C:\Apache2.4\htdocs\WAC_DEV\DynamicItems\wac333\[1208]_OutputDocument.docx",
                MarkAsPrivate = false,
                RemoveWaterMark = false,
                SuggestedFileName = "test.docx",
                CreatePdfAttachment = false,
                SubmitForApproval = (int)(2) == 3 ? true : false,
                SaveOption = ApplicationConstants.VersionTypeEnum.Internal,
                AtleastOneClauseChangedOrDeleted = true, // Todo: Need to check for if clause changed or not
                ReconcileFields = true,
                ReconcileClauseApprovals = true,
                ReconcileClauses = true,
                ReconcileTables = true
            });
        }
        public void SaveDocx()
        {
            var wu = new WordUtil.WordUtil();
            wu.ConvertHtmlToDocx(@"C:\Apache2.4\htdocs\DynamicItems\testHMLDiff\[12496]_InpuHtmltDocument.html", 0, false, string.Empty, string.Empty);

        }

        public void IterateContentControl()
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(@"C:\Apache2.4\htdocs\WAC_QA\test.docx", false))
            {
               

                MainDocumentPart mainPart = doc.MainDocumentPart;

                DocumentSettingsPart docSetting =(DocumentSettingsPart)mainPart.GetPartsOfType<DocumentSettingsPart>().FirstOrDefault();
               DocumentProtection protection=(DocumentProtection)docSetting.Settings.Descendants<DocumentProtection>().FirstOrDefault();
                
                foreach (var cc in mainPart.Document.Descendants<SdtElement>())
                {
                    SdtProperties props = cc.Elements<SdtProperties>().FirstOrDefault();
                    SdtId tag = props.Elements<SdtId>().FirstOrDefault();
                    Console.WriteLine(tag.Val.ToString());
                }
            }
        }

        public void GetClausePackage()
        {
            WordUtil.WordUtil wordUtil = new WordUtil.WordUtil();
            Word.Document document =
                wordUtil.ReadWordDoc(@"C:\Apache2.4\htdocs\WAC_QA\Mac\clausewithimag\HD_Image with Text.docx");
            string str = document.WordOpenXML;
            wordUtil.Close();
        }


        public string WordPackageXml()
        {
            Package package = Package.Open(@"C:\Apache2.4\htdocs\WAC_QA\Mac\clausewithimag\HD_Image with Text.docx");
            PackagePartCollection partcoll = package.GetParts();
            StringBuilder strPackageXml = new StringBuilder();
            strPackageXml.Append("<?xml version=\"1.0\" standalone=\"yes\"?>");
            strPackageXml.Append("<?mso-application progid=\"Word.Document\"?>");
            strPackageXml.Append("<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">");


            foreach (PackagePart pckpart in partcoll)
            {
                using (Stream stream = pckpart.GetStream())
                {
                    if (pckpart.ContentType.Contains("image"))
                    {
                        var memoryStream = new MemoryStream();
                        stream.CopyTo(memoryStream);
                        var value = Convert.ToBase64String(memoryStream.ToArray());
                        strPackageXml.Append("<pkg:part pkg:name=\"" + pckpart.Uri.OriginalString + "\" pkg:contentType=\"" + pckpart.ContentType + "\" pkg:compression=\"store\">");
                        strPackageXml.Append("<pkg:binaryData>");
                        strPackageXml.Append(value);
                        strPackageXml.Append("</pkg:binaryData>");
                        strPackageXml.Append("</pkg:part>");
                    }
                    else
                    {
                        XmlDocument xmldoc = new XmlDocument();
                        xmldoc.Load(stream);
                        foreach (XmlNode node in xmldoc)
                        {
                            if (node.NodeType == XmlNodeType.XmlDeclaration)
                            {
                                xmldoc.RemoveChild(node);
                            }
                        }

                        strPackageXml.Append("<pkg:part pkg:name=\"" + pckpart.Uri.OriginalString + "\" pkg:contentType=\"" + pckpart.ContentType + "\">");
                        strPackageXml.Append("<pkg:xmlData>");
                        strPackageXml.Append(xmldoc.InnerXml);
                        strPackageXml.Append("</pkg:xmlData>");
                        strPackageXml.Append("</pkg:part>");
                    }
                }

            }
            strPackageXml.Append("</pkg:package>");

            string packageXML = strPackageXml.ToString();
            byte[] array = Encoding.UTF8.GetBytes(packageXML);

            //Encode to Base64 twice - Apptus metadata contain base64 metadata, and on UI div tag contain base64 metadata
            string sPropertyValue = Convert.ToBase64String(array);
            array = Encoding.UTF8.GetBytes(sPropertyValue);
            sPropertyValue = Convert.ToBase64String(array);
            package.Close();
            return sPropertyValue;
        }
        public void CreateMArkUpClause()
        {
            using (MemoryStream _msStream = new MemoryStream())
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Create("file123.docx", WordprocessingDocumentType.Document))
                {
                    string altChunkId = "myId";
                    MainDocumentPart mainDocPart = doc.AddMainDocumentPart();
                    mainDocPart.Document = new Document();
                    mainDocPart.Document.Body = new Body();
                    MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes("<html><head></head><body><p>hey this is complete </p> <div><p class=\"MsoNormal\" contenteditable=\"true\">                        <b></b><b>To make</b> your document look professionally produced, Word                        provides header, footer, cover page, and text box designs that complement each                        other. For example, you can add a matching cover page, header, and sidebar.                        Click Insert and then choose the elements you want from the different                        galleries.                    </p></body></html>"));
                    AlternativeFormatImportPart formatImportPart = mainDocPart.AddAlternativeFormatImportPart(
                           AlternativeFormatImportPartType.Html, altChunkId);
                    //ms.Seek(0, SeekOrigin.Begin);

                    // Feed HTML data into format import part (chunk).
                    formatImportPart.FeedData(ms);
                    AltChunk altChunk = new AltChunk();
                    altChunk.Id = altChunkId;

                    mainDocPart.Document.Body.Append(altChunk);
                    mainDocPart.Document.Save();
                }
                // _msStream.Position = 0;
                // _msStream.WriteTo(new FileStream("file123.docx", FileMode.Create, FileAccess.Write));
            }

        }

        public void ExtractWaterMark()
        {
            string strWaterMarkText = string.Empty;
            string strWaterMarkStyle = string.Empty;
            string strShapeStyle = string.Empty;
            using (WordprocessingDocument doc = WordprocessingDocument.Open(@"C:\Apache2.4\htdocs\WAC_DEV\DynamicItems\watermark\watermark.docx", false))
            {
                MainDocumentPart mainPart = doc.MainDocumentPart;
                if (mainPart.HeaderParts.Count() > 0)
                {

                    HeaderPart headerPart = mainPart.HeaderParts.FirstOrDefault();
                    Picture picture = headerPart.Header.Descendants<Picture>().FirstOrDefault();
                    if (picture != null)
                    {
                        V.Shape shapes = picture.Descendants<V.Shape>().FirstOrDefault();
                        if (shapes != null)
                        {
                            strShapeStyle = shapes.Style;
                            var textpath = shapes.Descendants<V.TextPath>().FirstOrDefault();

                            if (textpath != null)
                            {
                                strWaterMarkText = textpath.String;
                                strWaterMarkStyle = textpath.Style;
                            }
                        }
                    }


                }
            }

        }

        public void checkIFWaterMarkExist()
        {
            var wu = new WordUtil.WordUtil();
            var document = wu.ReadWordDoc(@"C:\Apache2.4\htdocs\WAC_DEV\DynamicItems\watermark\WaterMarkwithImage.docx");
            DocumentWaterMarker.CheckWaterMarkExist(document, null);

            wu.Close();
        }

        public void readAptOpt()
        {
            var docProp = new Apttus.XAuthor.Integration.Service.Document.DocProp(@"C:\Apache2.4\htdocs\WAC_DEV\DynamicItems\WAC 381\AJ_Sep29_1_Original_XA_Web_Reconcile_4_2015-09-30.docx");
            if (docProp.IsApttusDoc())
            {
                var sMergeInfo = docProp.GetDocumentProperty("MergeInfo");
                var sBusinessObjectContext = docProp.GetDocumentProperty("SF_BUSINESS_OBJECT_CONTEXT");
                if (sMergeInfo != null && sBusinessObjectContext.Length > 0)
                {
                }
            }
            docProp.Close();
        }

        public void extractImageClause()
        {
            // var wu = new WordUtil.WordUtil();
            // var sHtmlDocFilePath = wu.ConvertDocxToHtml(@"C:\Apache2.4\htdocs\WAC_DEV\DynamicItems\wac281\testExtractImage.docx");
            //  wu.Close();

            var wtb = new ParsingHTML.WhtmlToBhtml();
            wtb.SContent =
                File.ReadAllText(@"C:\Apache2.4\htdocs\WAC_DEV\DynamicItems\wac281\testExtractImage.html", Encoding.Default);
            wtb.ConvertToXHtml();
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(wtb.SContent);
            XmlNamespaceManager xmlNameManager = ParsingHTML.HtmlUtils.SetNamespaceManager(xmlDoc);

            XmlNode nBody = xmlDoc.SelectSingleNode(".//default:body", xmlNameManager);
            XmlNodeList imgList = nBody.SelectNodes(".//default:img", xmlNameManager);

            foreach (XmlNode imageNode in imgList)
            {
                var imgsrc = imageNode.Attributes["src"] != null ? imageNode.Attributes["src"].Value : string.Empty;

                if (!string.IsNullOrEmpty(imgsrc))
                {
                    XmlAttribute altAttr = imageNode.Attributes["alt"] == null ? xmlDoc.CreateAttribute("alt") : imageNode.Attributes["alt"];
                    altAttr.Value = imgsrc;
                    imageNode.Attributes.Append(altAttr);

                    using (Image image = Image.FromFile(@"C:\Apache2.4\htdocs\WAC_DEV\DynamicItems\wac281\" + imgsrc))
                    {
                        using (MemoryStream ms = new MemoryStream())
                        {
                            // Convert Image to byte[]
                            image.Save(ms, image.RawFormat);
                            byte[] imageBytes = ms.ToArray();

                            // Convert byte[] to Base64 String
                            string base64String = Convert.ToBase64String(imageBytes);
                            imageNode.Attributes["src"].Value = "data:image/" + ImageCodecInfo.GetImageEncoders().FirstOrDefault(x => x.FormatID == image.RawFormat.Guid).FilenameExtension.ToLower().Remove(0, 2) + ";base64," + base64String;
                        }
                    }

                }

            }

            wtb.SContent = xmlDoc.OuterXml;
        }



        public Image Base64ToImage(string base64String)
        {
            // Convert Base64 String to byte[]
            byte[] imageBytes = Convert.FromBase64String(base64String);
            MemoryStream ms = new MemoryStream(imageBytes, 0,
              imageBytes.Length);

            // Convert byte[] to Image
            ms.Write(imageBytes, 0, imageBytes.Length);
            Image image = Image.FromStream(ms, true);
            return image;
        }

        public void ImageDoctoWord()
        {

            string sFileName = "1234.docx";//"AJ_WAC-339_Sep27_1_CheckIn_FromXAWord";
            string sDirectory = @"C:\Apache2.4\htdocs\WAC_DEV\DynamicItems\docxImage\";
            string sHtmlContent = File.ReadAllText(Path.Combine(sDirectory, sFileName + ".html"));
            BhtmlToWhtml b2w = new BhtmlToWhtml(sHtmlContent, sFileName, sDirectory);
            //sFileName += ".html";
            // XDocument xmldoc = new XDocument();
            //  string FileName = b2w.PrepareForWord(sFileName, out xmldoc);
            // this need to be called in preparefor word
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(b2w.SContent);


            // XmlNode nBody = xmlDoc.SelectSingleNode(".//default:body", namespaceMgr);
            List<XmlNode> imgList = ImageNodeList(xmlDoc.FirstChild.ChildNodes[1]);
            if (!Directory.Exists(sDirectory + sFileName + "_files"))
            {
                Directory.CreateDirectory(sDirectory + sFileName + "_files");

            }
            Dictionary<string, string> dictImgGUIDUrlMapping = new Dictionary<string, string>();
            foreach (XmlNode imageNode in imgList)
            {
                var imgsrc = imageNode.Attributes["src"] != null ? imageNode.Attributes["src"].Value : string.Empty;

                if (!string.IsNullOrEmpty(imgsrc))
                {
                    Image imageStream = Base64ToImage(imgsrc.Split(',')[1]);
                    //string[] imageDetail= imageNode.Attributes["alt"].Value.Split('/');
                    imageStream.Save(sDirectory + sFileName + "_files" + "\\" + imageNode.Attributes["alt"].Value + ".gif");

                    imageNode.Attributes["src"].Value = sFileName + "_files" + "/" + imageNode.Attributes["alt"].Value + ".gif";
                    dictImgGUIDUrlMapping[imageNode.Attributes["alt"].Value] = sFileName + "_files" + "/" + imageNode.Attributes["alt"].Value + ".gif";
                    imageNode.Attributes.Remove(imageNode.Attributes["alt"]);

                }
            }
            string strContent = xmlDoc.OuterXml;
            foreach (KeyValuePair<string, string> guidURLMapping in dictImgGUIDUrlMapping)
            {
                strContent = strContent.Replace("src=\"" + guidURLMapping.Key + "\"", "src=\"" + guidURLMapping.Value + "\"");
            }
            strContent = strContent.Replace("htmlimage_files", sFileName + "_files");
            File.WriteAllText(Path.Combine(sDirectory, sFileName + ".html"), strContent, Encoding.Default);
            var wu = new WordUtil.WordUtil();
            wu.ConvertHtmlToDocx(Path.Combine(sDirectory, sFileName + ".html"), 2, false, new string[] { });
            wu.Close();

        }
        public List<XmlNode> xmlNodeList = new List<XmlNode>();
        private List<XmlNode> ImageNodeList(XmlNode childNode)
        {
            foreach (XmlNode node in childNode.ChildNodes)
            {
                if (node.Name == "img")
                {
                    xmlNodeList.Add(node);
                }
                if (node.HasChildNodes)
                {
                    ImageNodeList(node);
                }
            }
            return xmlNodeList;
        }



        public void loadwordwithsource()
        {
            object oBlank = "";
            object oFalse = false;
            object oTrue = true;

            object encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;
            object openFormat = Word.WdOpenFormat.wdOpenFormatAuto;
            object direction = Word.WdDocumentDirection.wdLeftToRight;
            Word.Application m_wordApp = new Word.Application();
            Word.Documents docs = m_wordApp.Documents;
            // let WORD read the temporary input document
            Word.Document doc = docs.Open(@"C:\Apache2.4\htdocs\WAC_DEV\DynamicItems\docxImage\test\test.htm", OpenAndRepair: true);

            m_wordApp.ActiveWindow.View.ShowFieldCodes = true;



            object oSaveFormat = Word.WdSaveFormat.wdFormatXMLDocument;


            Object missing = Type.Missing;



            object oLineEnd = Word.WdLineEndingType.wdCRLF;

            doc.SaveAs(
                @"C:\Apache2.4\htdocs\WAC_DEV\DynamicItems\docxImage\test\test12.docx",
                ref oSaveFormat,
                ref missing,  // lock comments
                ref missing,  // password
                ref oTrue,  // add to recent files
                ref oBlank, // write password
                ref oFalse,  // suggest read only
                ref oFalse, // embed fonts
                ref oFalse,  // save graphics
                ref oFalse, // save forms
                ref oFalse //mail
                );

            ((Word._Document)doc).Close();

        }

        public void UpdateImageInDocx()
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(@"C:\Apache2.4\htdocs\WAC_DEV\DynamicItems\ca0f5581-9006-4486-9406-afcb38639296\test1.docx", true))
            {
                MainDocumentPart mainDocPart = doc.MainDocumentPart;
                foreach (Drawing d in mainDocPart.Document.Descendants<Drawing>().ToList())
                {
                    foreach (DocumentFormat.OpenXml.Drawing.Blip b in d.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().ToList())
                    {
                        if (string.IsNullOrEmpty(b.Embed))
                        {

                            var docProperties = d.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().ToList();
                            if (docProperties.Count > 0)
                            {
                                ImagePart imagePart = mainDocPart.AddImagePart(ImagePartType.Gif);
                                using (FileStream stream = new FileStream(docProperties[0].Description, FileMode.Open))
                                {
                                    imagePart.FeedData(stream);
                                }

                                b.Embed = mainDocPart.GetIdOfPart(imagePart);
                                //  //if (b.Embed.ToString() == imageRelationshipId)
                                ////  {
                                // // }

                                mainDocPart.Document.Save();
                            }
                        }
                    }
                }
            }

        }

        public void validateXML()
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(@"C:\Apache2.4\htdocs\WAC_DEV\DynamicItems\WAC429\HD39_WebAuthor_7th Oct_Redlines_Root level template_2015-10-07.docx.html");

            XmlNamespaceManager namespaceMgr;
            //Create an XmlNamespaceManager for resolving namespaces.
            namespaceMgr = new XmlNamespaceManager(xmlDoc.NameTable);
            namespaceMgr.AddNamespace("o", "urn:schemas-microsoft-com:office:office");
            namespaceMgr.AddNamespace("w", "urn:schemas-microsoft-com:office:word");
            namespaceMgr.AddNamespace("wx", "http://schemas.microsoft.com/office/word/2003/auxHint");
            namespaceMgr.AddNamespace("aml", "http://schemas.microsoft.com/aml/2001/core");
            namespaceMgr.AddNamespace("m", "http://schemas.microsoft.com/office/2004/12/omml");
            namespaceMgr.AddNamespace("v", "urn:schemas-microsoft-com:vml");
            namespaceMgr.AddNamespace("default", "http://www.w3.org/TR/REC-html40");
            namespaceMgr.AddNamespace("xlmns", "http://www.w3.org/TR/REC-html40");
            xmlDoc.ChildNodes[0].ChildNodes[0].SelectNodes("//o:smarttagtype", namespaceMgr);
        }

        public void ExtractClause()
        {
            string clausetext = File.ReadAllText(@"C:\Apache2.4\htdocs\WAC_DEV\DynamicItems\Comments\c.txt");
            var clauseObj = XmlSerializer<AgreementClauseContext>.Deserialize(clausetext);
        }

        public void TemplateName()
        {
            IWebAuthor webAuthorObj = new WebAuthorCaller("00DM0000001YHTB!AQQAQD_VnoOlOgcfKgNchlT95lAQDVyGrKGBjp66XBihTKJMGazXTJi.B9dKuUAFV7DoCn_EQSOMgR3I1yOEfNtDAjLHJKSI", "https://apttus.cs7.visual.force.com/services/Soap/u/30.0/00DM0000001YHTB");
            var attachmentService = new AttachmentService(webAuthorObj.Session);
            var aptAttachment = attachmentService.GetAttachment("00PM0000003z1nWMAQ");
        }

        public void GetTemplateName()
        {
            IWebAuthor webAuthorObj = new WebAuthorCaller("00DM0000001YHTB!AQQAQAGL2pBkLFt_JSBvIENi3UrZIZNpDpsFXfLT506Vs9HaHh3ofim0TfNk55G.IThJpQfhwGh1SoHHByLkw.nJ1wEfst7V", "https://apttus.cs7.visual.force.com/services/Soap/u/30.0/00DM0000001YHTB");
            var templateObject = new Apttus.XAuthor.Integration.Access.TemplateManager(webAuthorObj.Session).GetTemplateByReferenceID("c72635f5-9381-4f43-a0b8-29972b0ac171");
            var templateName = templateObject != null ? templateObject.Name : string.Empty;

        }

        public void GetAgreementName()
        {
            IWebAuthor webAuthorObj = new WebAuthorCaller("00DM0000001YHTB!AQQAQGbQrmpJaBDlRVK0h5VgtUuLVKaa5tuZ9.PrOkRd5K17A4eVlgdAum2u06Y7t_3kNBeNeyDFCOFwZxoBfZJhI3tEImFd", "https://apttus.cs7.visual.force.com/services/Soap/u/30.0/00DM0000001YHTB");

            var templateObject = new Apttus.XAuthor.Integration.Access.AgreementManager(webAuthorObj.Session).GetById("a07M0000008VsMTIA0");
            var templateName = templateObject != null ? templateObject.Name : string.Empty;
        }
    }
}
