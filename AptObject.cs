using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;


using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using Util;

namespace UnitTest
{
    public class AptObject
    {
        private XmlNode m_nSdtRun = null;
        private XmlDocument m_xmlDoc = null;
        private XmlNamespaceManager m_xmlNamespaceManager = null;
        private string m_sType = "";
        private string m_sSubType = "";
        private string m_sID = "";
        private TraceLog m_log = TraceLog.getInstance();

        public static bool IsAptObject(XmlNode nSdtRun, XmlNamespaceManager xmlNamespaceManager)
        {
            XmlNode nDocParCategory = nSdtRun.SelectSingleNode(".//w:docPartCategory", xmlNamespaceManager);
            if (nDocParCategory != null)
            {
                XmlAttribute attVal = nDocParCategory.Attributes["w:val"];
                if (attVal != null)
                {
                    string sCodedMetadata = attVal.Value;
                    string sMetadata = Utilities.DecodeString(sCodedMetadata);
                    if (!string.IsNullOrEmpty(sMetadata))
                    {
                        // this must be an Apttus content control
                        return true;
                    }
                }
            }
            return false;
        }

        public string Type
        {
            get
            {
                return m_sType;
            }
        }

        public string SubType
        {
            get
            {
                return m_sSubType;
            }
        }

        public AptObject(XmlNode nSdtRun, XmlNamespaceManager xmlNamespaceManager)
        {
            m_nSdtRun = nSdtRun;
            m_xmlNamespaceManager = xmlNamespaceManager;
            //xml->w:document->w:body->w:sdt->w:sdtPr->w:alias->w:val
            XmlNode nAlias = nSdtRun.SelectSingleNode(".//w:sdtPr/w:alias", m_xmlNamespaceManager);
            if (nAlias != null)
            {
                XmlAttribute attVal = nAlias.Attributes["w:val"];
                if (attVal != null)
                {
                    m_sType = attVal.Value;
                }
            }

            XmlNode nID = nSdtRun.SelectSingleNode(".//w:sdtPr/w:id", m_xmlNamespaceManager);
            if (nID != null)
            {
                XmlAttribute attID = nID.Attributes["w:val"];
                if (attID != null)
                {
                    m_sID = attID.Value;
                }
            }

            XmlNode nDocParCategory = m_nSdtRun.SelectSingleNode(".//w:docPartCategory", xmlNamespaceManager);
            m_xmlDoc = new XmlDocument();

            if (nDocParCategory != null)
            {
                XmlAttribute attVal = nDocParCategory.Attributes["w:val"];
                if (attVal != null)
                {
                    string sCodedMetadata = attVal.Value;
                    string sMetadata = Utilities.DecodeString(sCodedMetadata);
                    if (string.IsNullOrEmpty(sMetadata))
                    {
                        // not an Apttus content control
                        return;
                    }

                    System.Diagnostics.Debug.WriteLine(sMetadata);
                    m_xmlDoc = new XmlDocument();
                    m_xmlDoc.LoadXml(sMetadata);

                    XmlNode nType = m_xmlDoc.SelectSingleNode(".//Type");
                    if (nType != null)
                    {
                        m_sType = nType.InnerText;
                    }

                    XmlNode nSubType = m_xmlDoc.SelectSingleNode(".//SubType");
                    if (nSubType != null)
                    {
                        m_sSubType = nSubType.InnerText;
                    }
                }
            }
            else
            {
                m_xmlDoc.AppendChild(m_xmlDoc.CreateElement("Metadata"));
            }

            //XmlNode nAlias = m_nSdtRun.SelectSingleNode(".//w:sdtPr/w:alias", m_xmlNamespaceManager);
            //if (nAlias != null)
            //{
            //    m_sType = nAlias.Attributes["w:val"].Value;
            //}
        }

        private XmlAttribute CreateDocPartCateforyValueAttrib(XmlNode nSdtProp)
        {
            string sURI = m_xmlNamespaceManager.LookupNamespace("w");

            XmlNode nDocPartList = nSdtProp.SelectSingleNode(".//w:docPartList", m_xmlNamespaceManager);
            if (nDocPartList == null)
            {
                nDocPartList = nSdtProp.OwnerDocument.CreateElement("w", "docPartList", sURI);
                nSdtProp.AppendChild(nDocPartList);
            }
            XmlNode nDocPartCat = nDocPartList.SelectSingleNode(".//w:docPartCategory", m_xmlNamespaceManager);
            if (nDocPartCat == null)
            {
                nDocPartCat = nSdtProp.OwnerDocument.CreateElement("w", "docPartCategory", sURI);
                nDocPartList.AppendChild(nDocPartCat);
            }

            XmlAttribute attVal = nDocPartCat.Attributes["w:val"];
            if (attVal == null)
            {
                attVal = nSdtProp.OwnerDocument.CreateAttribute("w:val", sURI);
                nDocPartCat.Attributes.Append(attVal);
            }
            return attVal;
        }

        public void Save()
        {
            XmlAttribute attVal = null;
            string sXmlContent = m_xmlDoc.OuterXml;
            XmlNode nDocParCategory = m_nSdtRun.SelectSingleNode(".//w:docPartCategory", m_xmlNamespaceManager);
            if (nDocParCategory == null)
            {
                XmlNode nProp = m_nSdtRun["w:sdtPr"];
                if (nProp != null)
                {
                    attVal = CreateDocPartCateforyValueAttrib(nProp);
                }
            }
            else
            {
                attVal = nDocParCategory.Attributes["w:val"];

            }

            string sCodedMetadata = Utilities.CodeString(sXmlContent);
            attVal.Value = sCodedMetadata;


        }

        public void RemoveObjectProperty(string sPropName)
        {
            if (m_xmlDoc.DocumentElement == null)
            {
                Exception ex = new Exception("MergeService Internal Error : Invalid AptObject");
                throw ex;
            }

            XmlNode nProperty = m_xmlDoc.SelectSingleNode(".//" + sPropName);
            if (nProperty != null)
            {
                nProperty.ParentNode.RemoveChild(nProperty);
            }
        }

        public string GetObjectProperty(string sPropName)
        {
            if (m_xmlDoc.DocumentElement == null)
            {
                Exception ex = new Exception("MergeService Internal Error : Invalid AptObject");
                throw ex;
            }

            XmlNode nProperty = m_xmlDoc.SelectSingleNode(".//" + sPropName);
            if (nProperty != null)
            {
                return nProperty.InnerText;
            }
            return "";
        }

        public void SetObjectProperty(string sPropName, string sPropValue)
        {
            if (m_xmlDoc.DocumentElement == null)
            {
                Exception ex = new Exception("MergeService Internal Error : Invalid AptObject");
                throw ex;
            }

            XmlNode nProperty = m_xmlDoc.SelectSingleNode(".//" + sPropName);
            if (nProperty == null)
            {
                nProperty = m_xmlDoc.CreateElement(sPropName);
                m_xmlDoc.DocumentElement.AppendChild(nProperty);
            }
            nProperty.InnerText = sPropValue;
        }

        /// <summary>
        /// Retrieves all children of the selecteed property in a dictionary.
        /// </summary>
        /// <param name="sPropName">Name of the property</param>
        /// <returns>A dictionary with name/value pairs for each child item</returns>
        public Dictionary<string, string> GetObjectPropertyItem(string sPropName)
        {
            Dictionary<string, string> items = new Dictionary<string, string>();
            XmlNode nProperty = m_xmlDoc.SelectSingleNode(".//" + sPropName);
            if (nProperty != null && nProperty.HasChildNodes)
            {
                foreach (XmlNode nItem in nProperty.ChildNodes)
                {
                    items.Add(nItem.Name, nItem.InnerText);
                }
            }
            return items;
        }

        /// <summary>
        /// Adds children to the selected property based on the given name/value pairs in the given dictionay
        /// </summary>
        /// <param name="sPropName">Name of the property</param>
        /// <param name="items">A dictonary with children's name/value pairs.</param>
        public void SetObjectPropertyItem(string sPropName, Dictionary<string, string> items)
        {
            XmlNode nProperty = m_xmlDoc.SelectSingleNode(".//" + sPropName);
            if (nProperty == null)
            {
                nProperty = m_xmlDoc.CreateElement(sPropName);
                m_xmlDoc.DocumentElement.AppendChild(nProperty);
            }
            else if (nProperty.HasChildNodes)
            {
                nProperty.RemoveAll();
            }
            foreach (string sKey in items.Keys)
            {
                XmlNode nItem = m_xmlDoc.CreateElement(sKey);
                nProperty.AppendChild(nItem);
                nItem.InnerText = items[sKey];
            }
        }

        /// <summary>
        /// Get all proerty name/value pairs in a dictionry. If a property has child nodes, 
        /// thier values will be returned in a comma separated string instead.
        /// </summary>
        /// <returns>A dictionary with properties name/value pairs</returns>
        public Dictionary<string, string> GetAllProperties()
        {
            Dictionary<string, string> props = new Dictionary<string, string>();
            if (m_xmlDoc.DocumentElement == null)
            {
                m_log.error("The node with the type: {0} and id: {1} has no metadata!", m_sType, m_sID);
                return props;
            }

            foreach (XmlNode nProp in m_xmlDoc.DocumentElement.ChildNodes)
            {
                if (nProp.HasChildNodes)
                {
                    string sValue = "";
                    foreach (XmlNode nChild in nProp.ChildNodes)
                    {
                        sValue += nChild.InnerText + ", ";
                    }
                    sValue = sValue.Substring(0, sValue.Length - 2);
                    props.Add(nProp.Name, sValue);
                }
                else
                {
                    props.Add(nProp.Name, nProp.InnerText);
                }
            }
            return props;
        }

        public XmlNode GetAllPropertiesAsXml()
        {
            return m_xmlDoc.DocumentElement;
        }

    }
}

