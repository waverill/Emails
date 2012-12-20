using System;
using System.IO;
using System.Text;
using System.Xml;
using ICSharpCode.SharpZipLib.Zip;

namespace Emails
{
        class DocxTextReader
        {               
                private string file = "";
                private string location = "";
                
                // constructor, with the fileName you want to extract the text from
                public DocxTextReader(string theFile)   {               file = theFile;   }
 
                // Here the do it all method, call it after the constructor
                // it will try to find and parse document.xml from the zipped file
                // and return the docx's text in a string
                public string getDocumentText()
                {
                        if (string.IsNullOrEmpty(file))
                        {
                                throw new Exception("No Input file");
                        }
                
                        location = getDocumentXmlFile_FromZipFile();

                        if (string.IsNullOrEmpty(location))
                        {
                                throw new Exception("Invalid Docx");
                        }

                        return ReadDocumentText();
                }

                // we go to the xml file location
                // load it
                // and return the extracted text
                private string ReadDocumentText()
                {
                        StringBuilder result = new StringBuilder();

                        string bodyXPath = "/w:document/w:body";

                        ZipFile zipped = new ZipFile(file);
                        foreach (ZipEntry entry in zipped)
                        {
                                if (string.Compare(entry.Name, location, true) == 0)
                                {
                                        XmlDocument xmlDoc = new XmlDocument();
                                        xmlDoc.PreserveWhitespace = true;
                                        xmlDoc.Load(zipped.GetInputStream(entry));
                                        
                                        XmlNamespaceManager xnm = new XmlNamespaceManager(xmlDoc.NameTable);
                                        xnm.AddNamespace("w", @"http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                                        XmlNode node = xmlDoc.DocumentElement.SelectSingleNode(bodyXPath, xnm);

                                        if (node == null) { return ""; }
                                        result.Append(ReadNode(node));
                                        break;
                                }
                        }
                        zipped.Close();

                        return result.ToString();
                }

                // Xml node reader helper :D
                private string ReadNode(XmlNode node)
                {
                        // not a good node ?
                        if (node == null || node.NodeType != XmlNodeType.Element) { return ""; }

                        StringBuilder result = new StringBuilder();
                        foreach (XmlNode child in node.ChildNodes)
                        {
                                // not an element node ?
                                if (child.NodeType != XmlNodeType.Element) { continue; }

                                // lets get the text, or replace the tags for the actua text's characters
                                switch (child.LocalName)
                                {
                                        case "tab": result.Append("\t"); break;
                                        case "p": result.Append(ReadNode(child)); result.Append("\r\n\r\n"); break;
                                        case "cr":
                                        case "br": result.Append("\r\n"); break;

                                        case "t": // its Text !
                                                result.Append(child.InnerText.TrimEnd());
                                                string space = ((XmlElement)child).GetAttribute("xml:space");
                                                if (!string.IsNullOrEmpty(space) && space == "preserve") { result.Append(' '); }
                                        break;

                                        default:  result.Append(ReadNode(child));   break;
                                }
                        }

                        return result.ToString();
                }

                // lets open the zip file and look up for the
                // document.xml file
                // and save its zip location into the location variable
                private string getDocumentXmlFile_FromZipFile()
                {
                        // ICsharpCode helps here to open the zipped file
                        ZipFile zip = new ZipFile(file);

                        // lets take a look to the file entries inside the zip file
                        // up to we get
                        foreach (ZipEntry entry in zip)
                        {

                                if (string.Compare(entry.Name, "[Content_Types].xml", true) == 0)
                                {
                                        Stream contentTypes = zip.GetInputStream(entry);

                                        XmlDocument xmlDoc = new XmlDocument();
                                        xmlDoc.PreserveWhitespace = true;
                                        xmlDoc.Load(contentTypes);

                                        contentTypes.Close();

                                        // we need a XmlNamespaceManager for resolving namespaces
                                        XmlNamespaceManager xnm = new XmlNamespaceManager(xmlDoc.NameTable);
                                        xnm.AddNamespace("t", @"http://schemas.openxmlformats.org/package/2006/content-types");

                                        // lets find the location of document.xml
                                        XmlNode node = xmlDoc.DocumentElement.SelectSingleNode("/t:Types/t:Override[@ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"]", xnm);

                                        if (node != null)
                                        {
                                                string location = ((XmlElement)node).GetAttribute("PartName");
                                                return location.TrimStart(new char[] { '/' });
                                        }
                                        break;
                                }
                        }

                        // close the zip
                        zip.Close();

                        return null;
                }

        }
                
}
