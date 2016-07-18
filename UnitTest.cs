using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ParsingHTML;
using System.IO;
using WordUtil;
using Newtonsoft.Json;
using System.Xml;
using System.Xml.Linq;
using Apttus.XAuthor.Integration.Service.Document;
using System.Web;
using System.Globalization;
using System.Net.Http;
using System.Net;

namespace UnitTest
{
    class UnitTest
    {
        static void Main(string[] args)
        {
            //GetUserInfo();
            //MakeSFConnection();
            //MakeSFConnectionspring16();
            //SoapHttpClient();
            //SoapHttpClientspring();
            //getfieldsdata();
            /*XmlDocument doc1 = new XmlDocument();
            string encode = HttpUtility.HtmlEncode("value&vaue");
            doc1.LoadXml("<value>A &amp; A </value>");
            */
            //string sPath = Environment.CurrentDirectory + @"\..\..\..\TestDocs\";
            //string sFilename = Path.Combine(sPath, "Sample - with comments.docx.zip");
            //string sAssetFolder = Path.Combine(sPath, "assets");
            //WhtmlToBhtml wtb = new WhtmlToBhtml(sFilename, sAssetFolder);
            //wtb.PrepareForBrowser("OutputFile.html");
            //UInt32 uiVal = (UInt32)Convert.ToInt64(-1852172378 );
            // DateTime date1 = new DateTime(DateTime.Parse("2015-10-06").Ticks, DateTimeKind.Local);

            /*string s1=Convert.ToDateTime("9/25/2015").ToString("MM/dd/yyyy");
            DateTime dt;
            if(DateTime.TryParse("9/25/2015", out dt))
            {
                Console.Write("");
            }
            */

            /*string strr = @"<meta http-equiv=Content-Type content='text/html; charset=windows-1252'>
                            <meta name=ProgId content=Word.Document>
                            <meta name=Generator content='Microsoft Word 15'>
                            <meta name=Originator content=Microsoft Word 15'>";


            string sFileName = "WAC333__2015-10-19_Redlines_2015-11-20.docx";//"AJ_WAC-339_Sep27_1_CheckIn_FromXAWord";
            string sDirectory = @"C:\Apache2.4\htdocs\WAC_DEV\DynamicItems\wac333\";
            string sHtmlContent = File.ReadAllText(Path.Combine(sDirectory, sFileName + ".html"));*/
            // ParsingHTML();
            //  string sDocxFile = SaveTest(sFileName, sHtmlContent, sDirectory);
            // convertdocx();
            //string sProp = DocPropTest(@"c:\temp\[5068]_OutputDocument.docx");
            //GetClauses();
            //string s = GetAgrementClause();
            //GetClauseContent();
            // string s1 = GetClauseRefernce();
            //base64test();
            //getfieldsdata();
            // UpdatePlacementOfTag();

            GetDocumentProperties("H4sIAAAAAAAEAI2Ra0vDMBSGf1HbdFtbCyHQbRGj9qKJQz+FrDmOgb2Qpv/fdlCJDtF8yznvc3hOgrPe7ru6Ml0Pxp5hIHi6jg20lu1Jgo5hfNSpp6M69jZpFHlKQeSt0lDrWIe1Xic4cADMb2W5vac7IcVbRck03Y6DlFkluMxOBmAOSlnj4EdyJrcvnBWU86WxKwtBX8UfU36j5omC5tVjJqg80GfOyoLc+AhdqKuOoz4tolCSo8tJn1SmWfbgGs+b3qmBN8pYZqEhwoyAg28lLKDpP5SFf7+jAzgyi1/oIx+5El/iOZgTsPa9I/gAZjh3Ldn46ym+StbJhCxFHDjJ4OrjPwEKe/VFCwIAAA==");
                //"H4sIAAAAAAAAA41S70/CMBT8i/YLtuGSpkmBKlU3FloX/dSM9UlIHFu6Tv99NwykQgz22717d7l7KSKtWTZVrpsWtNlDh9EA+xoOhi3xzN8G8VYljoqq2AmTKHLKEiJnkgRKxSqo1HSGPEuA+L1czx/pQkjxllM8uJu+k5Lkgkuy0wDjopQV8i42R+X8hbOMcn4iFutM0Fdxw+Uv1egoaJo/E0FlQTecrTN85/r+UXXFWNGHIqU/S/3jS57SYs5IaCcem67KjtelNsxAjYXuAXm/RkhA3X6UBv59R0tghTnlC9yJ69shzsFT0Dtgh/cGowJ0t28OOHRDd+qGQRwnyDsNb7Ce5UOEIItVSrOfrt4F3sDnHr4edNO3TA30Jb76VN9xb+WaZwIAAA==" );


           

            /// getLockStatus();
            //  getfields();

            //WordUtilTest wutest = new WordUtilTest();
           // wutest.WordPackageXml();
            /*string str = HttpUtility.HtmlEncode("AT&T");
            XmlDocument xmldoc = new XmlDocument();
            xmldoc.LoadXml("<Request><ActionName>GetClauseContent</ActionName><ClauseAttachmentID>a09M0000007BYi9IAG</ClauseAttachmentID><ServerURL>https://apttus.cs7.visual.force.com/services/Soap/u/30.0/00DM0000001YHTB</ServerURL><SessionID>00DM0000001YHTB!AQQAQOdoOmHEbWHX3_FVTgCILt1exHE0zn3gxWdzouK2gztELxaa1ghXnczHWExELrIkXRcjmhCq6HcS1DZi7dX79YvqDEoq</SessionID><AttachmentID>00PM0000003z1nWMAQ</AttachmentID><MarkText>On the Insert                            tab, the galleries include items that are designed to coordinate with the                            overall look of your document. You can use these galleries to insert tables,                            headers, footers,  lists, cover pages, and other document building blocks.                                 When you                                create pictures, charts                            , or diagrams                        [VK4]1                                                             ,                                they also coordinate with your current document look.</MarkText></Request>");

            xmldoc.Load(@"C:\Apache2.4\htdocs\WAC_DEV\DynamicItems\WAC339\WAC355.html");
            string sxml = @"<AptDocProperties>
                          <DocumentID>8d5d938e-40fe-4cb1-b960-9a0fd0e40318</DocumentID> 
                          <SF_OBJECT_TYPE>Apttus__APTS_Agreement__c</SF_OBJECT_TYPE> 
                          <SF_BUSINESS_OBJECT_CONTEXT>Apttus__APTS_Agreement__c</SF_BUSINESS_OBJECT_CONTEXT> 
                          <SF_TEMPLATE_VERSION>8.00</SF_TEMPLATE_VERSION> 
                          <SF_OBJECT_ID>a01i000000aVI13AAG</SF_OBJECT_ID> 
                          <TemplateID>8d5d938e-40fe-4cb1-b960-9a0fd0e40318</TemplateID> 
                          <HasSmartItem>True</HasSmartItem> 
                          <SF_OBJECT_VERSION>6.0.0</SF_OBJECT_VERSION> 
                         <MergeInfo>
                          <Version>4.2.0.26784</Version> 
                          </MergeInfo>
                          </AptDocProperties>";
            XmlDocument xmlPropDoc = new XmlDocument();
            xmlPropDoc.LoadXml(sxml);



            XDocument doc = XDocument.Parse(sxml);
            Dictionary<string, string> dataDictionary = new Dictionary<string, string>();

            foreach (XElement element in doc.Descendants().Where(p => p.HasElements == false))
            {
                int keyInt = 0;
                string keyName = element.Name.LocalName;

                while (dataDictionary.ContainsKey(keyName))
                {
                    keyName = element.Name.LocalName + "_" + keyInt++;
                }

                dataDictionary.Add(keyName, element.Value);
            }*/
           
        }

      

        private static void MakeSFConnection()
        {
            Apttus.Common.SForce.SforceService forceObj = new Apttus.Common.SForce.SforceService();
            forceObj.Url = "https://test.salesforce.com/services/Soap/u/26.0";
            Apttus.Common.SForce.LoginResult res = forceObj.login("anair@apttus.com.ee.test.wauthor", "Apttus123$rH59kl2XPOZzsobTdbhUrhR74");
            //forceObj.SessionHeaderValue = new Apttus.Common.SForce.SessionHeader();

            forceObj.SessionHeaderValue = new Apttus.Common.SForce.SessionHeader();
            forceObj.SessionHeaderValue.sessionId = res.sessionId;
            forceObj.Url = res.serverUrl;
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Ssl3 | System.Net.SecurityProtocolType.Tls | System.Net.SecurityProtocolType.Tls11 |
                 System.Net.SecurityProtocolType.Tls12;
            forceObj.getUserInfo();

        }

        private static void MakeSFConnectionspring16()
        {
            try
            {

                //var httpClient = new HttpClient (new NativeMessageHandler ());

                Apttus.Common.SForce.SforceService forceObj = new Apttus.Common.SForce.SforceService();
                forceObj.Url = "https://login.salesforce.com/services/Soap/u/30.0/";
                Apttus.Common.SForce.LoginResult res = forceObj.login("obelorusets@summer15.apttus.com", "apTTus16Kna2sYCqMWmclkqTRLL9RRFfq");
                //forceObj.SessionHeaderValue = new Apttus.Common.SForce.SessionHeader();

                forceObj.SessionHeaderValue = new Apttus.Common.SForce.SessionHeader();
                forceObj.SessionHeaderValue.sessionId = res.sessionId;
                forceObj.Url = res.serverUrl;
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Ssl3 | System.Net.SecurityProtocolType.Tls | System.Net.SecurityProtocolType.Tls11 |
                System.Net.SecurityProtocolType.Tls12;
                forceObj.getUserInfo();
            }
            catch (Exception exp)
            {
                Console.WriteLine(exp.StackTrace);
            }

        }

        private string LoginActionSoap(string username, string pass)
        {
            return String.Format(@"<?xml version=""1.0"" encoding=""utf-8""?>
                                <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" 
                            xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
                        <soap:Body>
                        <login xmlns=""urn:partner.soap.sforce.com"">
                            <username>{0}</username>
                            <password>{1}</password>
                    </login>
                    </soap:Body></soap:Envelope>", username, pass);
        }

        private static string GetUserInfoSoap(string sessionid)
        {
            return String.Format(@"<?xml version=""1.0"" encoding=""utf-8""?>
                                <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
                                <soap:Header><SessionHeader xmlns=""urn:partner.soap.sforce.com"">
                                <sessionId>{0}</sessionId></SessionHeader>
                                </soap:Header><soap:Body><getUserInfo xmlns=""urn:partner.soap.sforce.com"" /></soap:Body></soap:Envelope>", sessionid);
        }


        private static async void SoapHttpClient()
        {
            Apttus.Common.SForce.SforceService forceObj = new Apttus.Common.SForce.SforceService();
            forceObj.Url = "https://test.salesforce.com/services/Soap/u/26.0/";
            Apttus.Common.SForce.LoginResult res = forceObj.login("anair@apttus.com.ee.test.wauthor", "Apttus123$rH59kl2XPOZzsobTdbhUrhR74");
            string orgID = "00DM0000001YHTB";
            var content = new StringContent(GetUserInfoSoap(res.sessionId), Encoding.UTF8, "text/xml");
            using (var httpClient = new HttpClient())
            {
                var request = new HttpRequestMessage();

                request.RequestUri = new Uri(res.serverUrl);
                request.Method = HttpMethod.Post;
                request.Content = content;
                request.Headers.Add("SOAPAction", "getUserInfo");
                var responseMessage = await httpClient.SendAsync(request);
                var response = await responseMessage.Content.ReadAsStringAsync();

                if (responseMessage.IsSuccessStatusCode)
                {



                }
            }

        }

        private static async void SoapHttpClientspring()
        {
            Apttus.Common.SForce.SforceService forceObj = new Apttus.Common.SForce.SforceService();
            forceObj.Url = "https://login.salesforce.com/services/Soap/u/30.0/";
            Apttus.Common.SForce.LoginResult res = forceObj.login("obelorusets@summer15.apttus.com", "apTTus16Kna2sYCqMWmclkqTRLL9RRFfq");
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls; // comparable to modern browsers

            var content = new StringContent(GetUserInfoSoap(res.sessionId), Encoding.UTF8, "text/xml");
            using (var httpClient = new HttpClient())
            {
                var request = new HttpRequestMessage();

                request.RequestUri = new Uri(res.serverUrl);
                request.Method = HttpMethod.Post;
                request.Content = content;
                request.Headers.Add("SOAPAction", "getUserInfo");
                var responseMessage = await httpClient.SendAsync(request);
                var response = await responseMessage.Content.ReadAsStringAsync();

                if (responseMessage.IsSuccessStatusCode)
                {



                }
            }

        }

        private static void GetUserInfo()
        {
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls11 | System.Net.SecurityProtocolType.Tls12 | System.Net.SecurityProtocolType.Tls | System.Net.SecurityProtocolType.Ssl3;
            String sessionID = "00DB00000001Sln!AQgAQFzlIRbtpxhoqIeeyDAjX6XH2NkZmP2SEEfWE2qRL.uJnJDPJ5TCQ2Q8z.eOjNz39ij.T27hhAB02N4bknxXjPCJeyNj";
            // "00DB00000001Sln!AQgAQFzlIRbtpxhoqIeeyDAjX6XH2NkZmP2SEEfWE2qRL.uJnJDPJ5TCQ2Q8z.eOjNz39ij.T27hhAB02N4bknxXjPCJeyNj"
            //"00DM0000001YHTB!AQQAQHpeDhhmEC_9LHj7ZQkD8uCXgQZA_JxJU3q65jymZB_Y874arpEKpIUJqIoLboJ84Z.VMua2u0zt0hCP1yP6eQt0Yog.";
            String sessionURl = "https://gs0.salesforce.com/services/Soap/u/30.0/00DB00000001Sln";
            Apttus.Common.SForce.SforceService forceObj = new Apttus.Common.SForce.SforceService();

            //POCTLS12.SforceRef.SforceService forceObj = new POCTLS12.SforceRef.SforceService ();
            forceObj.Url = sessionURl;
            forceObj.SessionHeaderValue = new Apttus.Common.SForce.SessionHeader();
            forceObj.SessionHeaderValue.sessionId = sessionID;

            var userInfo = forceObj.getUserInfo();
        }

        private static void UpdatePlacementOfTag()
        {
            WhtmlToBhtml wtohtml = new WhtmlToBhtml();
            wtohtml.SContent = File.ReadAllText(@"C:\Apache2.4\htdocs\WAC_QA\Mac\HD_23 Dec Full Access_Redlines_XA_Web_Reconcile_4_2016-01-05.html");
            wtohtml.ConvertToXHtml();
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(wtohtml.SContent);
        }


        private static void base64test()
        {
            string encodedString = "H4sIAAAAAAAEAI2RbWuDMBSFf1E1VucUQiBtsy1jVVky2T4Fq9cizBdi/P/TgiNrGVs+BHLvOYfn3mA6mENfZrofQJsGRoLn59RCZ/iBRNVdFfsRbAJUz1d58janOESbuEB1hSBAvhdh1zJg8aDS3TPbSyU/MkbmdDONStFMCkXPGmARKlVi90q5OHdvgidMiLWxTxPJ3uUfKb+5lkTJjtkLlUzl7FXwNCGRg9DFddOx0OdBCuQ16HKKnHs+pY828TKphHb4LAz8e02WAT8Vo2gLbbiBlkg9AXZ/lCyYlS90kINsiG/wI+gz8K7uCc5Bj03fkcDZzvJteB8F2F2L2LWU7s3HfwHEIMdSCwIAAA==";
            byte[] data = Convert.FromBase64String(encodedString);
            string decodedString = Encoding.UTF8.GetString(data);
        }
        static void ParsingHTML()
        {
            // Set up the log path
            string sHomeDirectory = @"C:\VimalKumar\project\CGITEST"; //System.Environment.CurrentDirectory;
            TraceLog m_log = TraceLog.getInstance(sHomeDirectory + @"\");
            m_log.traceLevel = "Debug";
            //m_log.debugEnabled = true;
            // m_log.debug("debug data");

            //Got the word document created
            // WordUtil.WordUtil wu = new WordUtil.WordUtil();
            // string sHtmlDocFilePath = wu.ConvertDocxToHtml(@"C:\Apache2.4\htdocs\WAC_DEV\DynamicItems\WAC337\__2015-10-13.docx");
            //  wu.Close();
            /*
             string sFileName, string sCgiPath, string sObjectId, 
                                        string sLockStatus, string sUserInfo,
                                       string sAttachmentID, string sServerUrl, string sSessionID, bool bHasWaterMark
             */
            WhtmlToBhtml wtb = new WhtmlToBhtml(@"C:\Apache2.4\htdocs\WAC_DEV\DynamicItems\wac333\WAC333__2015-10-19.docx.html", string.Empty);
            wtb.PrepareForBrowser("1234.docx", "cgipath", "objectid", "lockstatus", "userinfo", "AttachmentID", "serverurl", "sessionid", true, "", "");
        }
        static string SaveTest(string sFileName, string sHtmlContent, string sDirectory)
        {
            string sWordReadyHtmlFile = ConvertToWordHtml(sFileName, sHtmlContent, sDirectory);
            WordUtil.WordUtil wu = new WordUtil.WordUtil();

            return wu.ConvertHtmlToDocx(sWordReadyHtmlFile, 2, false, new string[] { });
        }
        static void convertdocx()
        {
            WordUtil.WordUtil wu = new WordUtil.WordUtil();

            wu.ConvertHtmlToDocx(@"C:\Apache2.4\htdocs\WAC_DEV\DynamicItems\wac333\WAC333__2015-10-19_Redlines_2015-11-20.docx.html", 2, false, new string[] { });
        }

        static string ConvertToWordHtml(string sFileName, string sModifiedHtmlContent, string sDirectory)
        {
            BhtmlToWhtml b2w = new BhtmlToWhtml(sModifiedHtmlContent, sFileName, sDirectory);
            sFileName += ".html";
            XDocument xmldoc = new XDocument();
            sFileName = b2w.PrepareForWord(sFileName, out xmldoc);

            return sFileName;
        }

        //static string DocPropTest(string sFileName)
        //{
        //    DocProp docProp = new DocProp(sFileName);

        //    bool bApttusDoc = docProp.IsApttusDoc();
        //    string sMergeInfo = docProp.GetDocumentProperty("MergeInfo/Version");

        //    return sMergeInfo;
        //}
        static string GetClauses()
        {
            String SessionID = "00D37000000KIWD!AQwAQCIQ_ax5wFBCRv5Nd.HZczKuMSCQJIeVofG2tYBDmdIyI4OzWzXAQmOKRoXj7.4cvpXvs_VAhmqLbpNsoFug96Z6ApPr";
            String URL = "https://na31.salesforce.com/services/Soap/u/30.0/00D37000000KIWD";
            String ObjectID = "a07370000029mpYAAQ";

            WebAuthorIntegration.IWebAuthor webobj = new WebAuthorIntegration.WebAuthorCaller(SessionID, URL);
            WebAuthorIntegration.IPlaybook playobj = webobj.PlayBook;
            playobj.AgreementID = ObjectID;
            WebAuthorIntegration.PlaybookFilterCriteria playbookFilter = playobj.PlaybookFilter;

            string s = playobj.GetClauses(playbookFilter);
            return "";

        }

        //require admin setting  this give the clauses exception 
        static string GetAgrementClause()
        {
            String SessionID = "ARgAQNeO8LDZNDhsdp0BLHKiqRT4sVWcRqDwfMlmCHOXCkRoY4kuzOYDgnfXdvHxAB3rAPEi5bBly1SPkETLsJuZbZzusj2R";
            String URL = "https://apttus.na15.visual.force.com/services/Soap/u/30.0/00Di000000L4aS";
            String AgreementIDObjectID = "a01i000000D63KgAAJ";//"a01i000000D63KgAAJ"; // this is agreement id

            WebAuthorIntegration.IWebAuthor webobj = new WebAuthorIntegration.WebAuthorCaller(SessionID, URL);

            WebAuthorIntegration.IPlaybook playobj = webobj.PlayBook;
            playobj.ObjectID = AgreementIDObjectID;
            var str = (playobj.GetAgreementClauses());
            return "";
        }

        static string GetClauseRefernce()
        {
            String SessionID = "00Di0000000L4aS!ARgAQJUC4eez55tJxBJH4ecb024mtFuHg4dpSn_mXsxXwedjMJc.vCzONuck8X0Cc9KXOisFJp0lKUUBMSp6vtGeGyZFdUUL";
            String URL = "https://apttus.na15.visual.force.com/services/Soap/u/30.0/00Di0000000L4aS";
            String AgreementIDObjectID = "a01i000000D63KgAAJ"; //a03i000000O1l9nAAB

            WebAuthorIntegration.IWebAuthor webobj = new WebAuthorIntegration.WebAuthorCaller(SessionID, URL);
            WebAuthorIntegration.IPlaybook playobj = webobj.PlayBook;
            playobj.AgreementID = AgreementIDObjectID;
            string str = (playobj.GetClausesRefrences());
            return "";
        }

        static void GetClauseContent()
        {
            String SessionID = "00Di0000000L4aS!ARgAQJTBgWfge7_MsZBrJAgL1ptYSHbB1reFCWLciuhnAldN3KtT9lp0sKoiUA_gkDs5ApwzK05wsxE9hcZFkTcpd9Zhqe3L";
            String URL = "https://apttus.na15.visual.force.com/services/Soap/u/30.0/00Di0000000L4aS";
            String ObjectID = "a03i000000O1l9nAAB";

            WebAuthorIntegration.IWebAuthor webobj = new WebAuthorIntegration.WebAuthorCaller(SessionID, URL);
            WebAuthorIntegration.IPlaybook playobj = webobj.PlayBook;
            playobj.ObjectID = ObjectID;

            byte[] byteArray = playobj.GetClauseContent();
            //string filePath = string.Empty;
            //using (MemoryStream mem = new MemoryStream())
            //{
            //    mem.Write(byteArray, 0, (int)byteArray.Length);
            //    string _clauseDocx = ObjectID + "_Clause.docx";
            //    // check for file already exist delete it create new
            //    using (FileStream fileStream = new FileStream(@"C:\VimalKumar\webAuthor\SourceCode\apttus_webauthor\dev\WebAuthor\main\UnitTest\bin\Debug\" + _clauseDocx, System.IO.FileMode.CreateNew))
            //    {
            //        mem.WriteTo(fileStream);
            //        filePath = fileStream.Name;
            //    }
            //}

            string sHomeDirectory = @"C:\VimalKumar\project\CGITEST"; //System.Environment.CurrentDirectory;
            TraceLog m_log = TraceLog.getInstance(sHomeDirectory + @"\");
            m_log.traceLevel = "Debug";
            //m_log.debugEnabled = true;
            // m_log.debug("debug data");

            //Got the word document created
            WordUtil.WordUtil wu = new WordUtil.WordUtil();
            string sHtmlDocFilePath = wu.ConvertDocxToHtml(@"C:\temp\DynamicItems\003cd68f-f133-45a4-a6e6-07fa005afa6a\" + ObjectID + "_Clause.docx"); //C:\VimalKumar\webAuthor\SourceCode\apttus_webauthor\dev\WebAuthor\main\UnitTest\bin\Debug
            wu.Close();

            //  WhtmlToBhtml wtb = new WhtmlToBhtml(@"C:\VimalKumar\webAuthor\SourceCode\apttus_webauthor\dev\WebAuthor\main\UnitTest\bin\Debug\" + ObjectID + "_Clause.html", string.Empty);
            //   wtb.PrepareForBrowser(@"C:\VimalKumar\webAuthor\SourceCode\apttus_webauthor\dev\WebAuthor\main\UnitTest\bin\Debug\content.html", @"C:\VimalKumar\webAuthor\SourceCode\apttus_webauthor\dev\WebAuthor\main\UnitTest\bin\Debug");

            //    WhtmlToBhtml wtb = new WhtmlToBhtml();
            //    wtb.SContent = File.ReadAllText(@"C:\VimalKumar\webAuthor\SourceCode\apttus_webauthor\dev\WebAuthor\main\UnitTest\bin\Debug\" + ObjectID + "_Clause.html", Encoding.Default);
            //  // wtb.PrepareForBrowser(@"C:\VimalKumar\webAuthor\SourceCode\apttus_webauthor\dev\WebAuthor\main\UnitTest\bin\Debug\content.html", @"C:\VimalKumar\webAuthor\SourceCode\apttus_webauthor\dev\WebAuthor\main\UnitTest\bin\Debug");
            //  string str = wtb.ReadClauseContent();    
        }



        private static void getfields()
        {
            String SessionID = "00Di0000000L4aS!ARgAQGu2G6YlTPXZmjojhqBOYhkrzTYzixEoQiSuY76HHvG.eRoYoWPODGbiz.J_AMhDRYHAmCoclUvIRArXSGT7R9oAvQHQ";
            String URL = "https://apttus.na15.visual.force.com/services/Soap/u/30.0/00Di0000000L4aS";
            String AgreementIDObjectID = "a01i000000D63KgAAJ"; //a03i000000O1l9nAAB

            WebAuthorIntegration.IWebAuthor webobj = new WebAuthorIntegration.WebAuthorCaller(SessionID, URL);
            WebAuthorIntegration.IPlaybook playobj = webobj.PlayBook;
            playobj.AgreementID = AgreementIDObjectID;

            string str = playobj.GetPlaybookField();
        }
        private static void getfieldsdata()
        {
            String SessionID = "00DM0000001YHTB!AQQAQI9eMp_1EPGOZtsa5IZEfIaz.oLWbnodItvC1jcNbbmJvuvNcBO_74hS6b6w1DEDppsQEu3ffax3r2at39.WWlYiVSyj";
            String URL = "https://apttus.cs7.visual.force.com/services/Soap/u/30.0/00DM0000001YHTB";
            String AgreementIDObjectID = "00PM0000004J1GaMAK"; //a03i000000O1l9nAAB

            WebAuthorIntegration.IWebAuthor webobj = new WebAuthorIntegration.WebAuthorCaller(SessionID, URL);
            WebAuthorIntegration.IDocument playobj = webobj.Document;
            //playobj.AgreementID = AgreementIDObjectID;
            String str1 = @"[{
                            'RecId':'a07M0000009TDPUIA4',
                             'FieldName' : 'Name',
                              'ReferenceObjectName': '',
                              'ObjectSource':'Apttus__APTS_Agreement__c/apts_agreement_name'
                            }]";
            var str = playobj.GetFieldData("Apttus__APTS_Agreement__c", str1);

            //Apttus__APTS_Agreement__c
            /*
             ,
                             {
                            'RecId':'a01i000000D63KgAAJ',
                             'FieldName' : 'Apttus__Contract_End_Date__c',
                              'ReferenceObjectName': '',
                              'ObjectSource':'Apttus__APTS_Agreement__c'
                            }
             */

        }

        private static void getLockStatus()
        {
            String SessionID = "00Di0000000L4aS!ARgAQOnVDl9fIRLzshf_Nv8D0Jciwtg7QA4ZWGSepvG6yu8l6aQ51IzJdatri1ws3ZF_OmSm67mpqLXJbPnrxjWs_EYvnctI";
            String URL = "https://apttus.na15.visual.force.com/services/Soap/u/30.0/00Di0000000L4aS";
            String AgreementIDObjectID = "a01i000000D63KgAAJ"; //a03i000000O1l9nAAB

            WebAuthorIntegration.IWebAuthor webobj = new WebAuthorIntegration.WebAuthorCaller(SessionID, URL);

            string userinfo = webobj.GetUserInfo();

            WebAuthorIntegration.IDocument playobj = webobj.Document;
            playobj.docLockService.GetLockStatus("a01i000000aVI13AAG", "8d5d938e-40fe-4cb1-b960-9a0fd0e40318");
        }

        private static Dictionary<string, string> GetDocumentProperties(string encodedDocProperties)
        {
            var _docAptusProperties = new DocProp();
            var ba = Convert.FromBase64String(encodedDocProperties);
            var strApttusDocPropXml = _docAptusProperties.Unzip(ba);

            var doc = XDocument.Parse(strApttusDocPropXml);
            var dataDictionary = new Dictionary<string, string>();

            foreach (var element in doc.Descendants().Where(p => p.HasElements == false))
            {
                var keyInt = 0;
                var keyName = element.Name.LocalName;

                while (dataDictionary.ContainsKey(keyName))
                {
                    keyName = element.Name.LocalName + "_" + keyInt++;
                }

                dataDictionary.Add(keyName, element.Value);
            }

            return dataDictionary;
        }
    }
}


//Microsoft.Office.Interop.Word.Application objWord = new Microsoft.Office.Interop.Word.Application();
//           objWord.Documents.Open(FileName: sInputDocFile);
//           objWord.Visible = false;
//           Microsoft.Office.Interop.Word.Document oDoc = objWord.ActiveDocument;
//           oDoc.SaveAs(FileName: oHtmlFileName, FileFormat: 10);
//           oDoc.Close(SaveChanges: false);
//           objWord.Application.Quit(SaveChanges: false);

//public StringBuilder Convert()
//{
//    Application objWord = new Application();

//    if (File.Exists(FileToSave))
//    {
//        File.Delete(FileToSave);
//    }
//    try
//    {
//        objWord.Documents.Open(FileName: FullFilePath);
//        objWord.Visible = false;
//        if (objWord.Documents.Count > 0)
//        {
//            Microsoft.Office.Interop.Word.Document oDoc = objWord.ActiveDocument;
//            oDoc.SaveAs(FileName: FileToSave, FileFormat: 10);
//            oDoc.Close(SaveChanges: false);
//        }
//    }
//    finally
//    {
//        objWord.Application.Quit(SaveChanges: false);
//    }
//    return base.ReadConvertedFile();
//}