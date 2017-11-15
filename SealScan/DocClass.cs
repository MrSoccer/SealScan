using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SealScan
{
    class DocClass
    {
        public DocClass()
        {
            //  Read the temp file
            //  Generate the form
            //  Print the form
            //  Delete the temp file
            //  Log the print job

        }

        public void UpdateDateStamp(string filepath)
        {
            // WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true);

            //const string wordmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            //const string relationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordprocessingDocument.MainDocumentPart.GetStream()))
                    docText = sr.ReadToEnd();

                DateTime dateValue = DateTime.Now;
                string timeStamp = dateValue.ToString("MM/dd/yy");

                //string templateFile = Properties.Settings.Default.FormTemplate;
                //string tempFile = "SealReport_" + timeStamp + ".docx";

                string dateMarker = string.Format(@"DATE:");
                string replacementText = string.Format("DATE: " + timeStamp);
                docText = new Regex(dateMarker).Replace(docText, replacementText);

                using (StreamWriter sw = new StreamWriter(wordprocessingDocument.MainDocumentPart.GetStream(FileMode.Create)))
                    sw.Write(docText);
            }
        }

        /*public void UpdateWordDocument1(string filepath)
        {
            // WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true);

            //const string wordmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            //const string relationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordprocessingDocument.MainDocumentPart.GetStream()))
                    docText = sr.ReadToEnd();

                for (int scanNumber = 1; scanNumber <= 51; scanNumber++)
                {
                    string formatedNumber = string.Format("{0,2:D2}", scanNumber);
                    string testText = string.Format(@"SEAL" + formatedNumber);
                    string replacementText = string.Format("SEAL #" + formatedNumber + " - " + "today");
                    docText = new Regex(testText).Replace(docText, replacementText);
                }

                using (StreamWriter sw = new StreamWriter(wordprocessingDocument.MainDocumentPart.GetStream(FileMode.Create)))
                    sw.Write(docText);
            }


        }

        */


        public void UpdateWordDocument(string filepath, List<string> barcodeList)
        {
            WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true);


            var tables = wordprocessingDocument.MainDocumentPart.Document.Descendants<Table>().ToList();

            foreach (Table t in tables)
            {
                int cellCounter = 0;

                Table table = wordprocessingDocument.MainDocumentPart.Document.Body.Elements<Table>().First();
                var rows = table.Elements<TableRow>();
                foreach (TableRow row in rows)
                {
                    var cells = row.Elements<TableCell>();

                    foreach (TableCell cell in cells)
                    {
                        if (cellCounter < barcodeList.Count)
                        {
                            cell.RemoveAllChildren();
                            int index = cellCounter + 1;
                            string sealText = string.Format("Seal #" + index + " - " + barcodeList[cellCounter]);
                            cell.Append(new Paragraph(new Run(new Text(sealText))));
                            //cell.Append(new Run(new Text(string.Format("SEAL # " + barcodeList[cellCounter]))));
                        }
                        cellCounter++;
                    }
                }
            }
            // Close the handle explicitly.
            wordprocessingDocument.Close();

        }

        internal bool CopyMasterFile(string masterFile, string tempFile)
        {
            bool RC = false;

            if (File.Exists(masterFile) && !File.Exists(tempFile))
            {
                if (!Directory.Exists(Path.GetDirectoryName(tempFile)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(tempFile));
                }
                File.Copy(masterFile, tempFile, true);
                RC = true;
            }
            else
            {
                MessageBox.Show("Template File is missing");
            }
            return RC;
        }

        internal void UpdateOrderNumber(string filepath, string orderNumberText)
        {
            // WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true);

            //const string wordmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            //const string relationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordprocessingDocument.MainDocumentPart.GetStream()))
                    docText = sr.ReadToEnd();

                DateTime dateValue = DateTime.Now;
                string timeStamp = dateValue.ToString("MM/dd/yy");

                string orderNumberMarker = string.Format(@"ORDER NUMBER:");
                string replacementText = string.Format("ORDER NUMBER: " + orderNumberText);
                docText = new Regex(orderNumberMarker).Replace(docText, replacementText);

                using (StreamWriter sw = new StreamWriter(wordprocessingDocument.MainDocumentPart.GetStream(FileMode.Create)))
                    sw.Write(docText);
            }
        }

        internal void UpdateVesselNumber(string filepath, string vesselNumberText)
        {
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordprocessingDocument.MainDocumentPart.GetStream()))
                    docText = sr.ReadToEnd();

                DateTime dateValue = DateTime.Now;
                string timeStamp = dateValue.ToString("MM/dd/yy");

                string orderNumberMarker = string.Format(@"VESSEL NUMBER:");
                string replacementText = string.Format("VESSEL NUMBER: " + vesselNumberText);
                docText = new Regex(orderNumberMarker).Replace(docText, replacementText);

                using (StreamWriter sw = new StreamWriter(wordprocessingDocument.MainDocumentPart.GetStream(FileMode.Create)))
                    sw.Write(docText);
            }
        }

        /*
        public string ReadWordDocument()
        {
            StringBuilder sb = new StringBuilder();
            OpenXmlElement element = Package.MainDocumentPart.Document.Body;
            if (element == null)
            {
                return string.Empty;
            }


            sb.Append(GetPlainText(element));
            return sb.ToString();
        }


        public string GetPlainText(OpenXmlElement element)
        {
            StringBuilder PlainTextInWord = new StringBuilder();
            foreach (OpenXmlElement section in element.Elements())
            {
                switch (section.LocalName)
                {
                    // Text 
                    case "t":
                        PlainTextInWord.Append(section.InnerText);
                        break;


                    case "cr":                          // Carriage return 
                    case "br":                          // Page break 
                        PlainTextInWord.Append(Environment.NewLine);
                        break;


                    // Tab 
                    case "tab":
                        PlainTextInWord.Append("\t");
                        break;


                    // Paragraph 
                    case "p":
                        PlainTextInWord.Append(GetPlainText(section));
                        PlainTextInWord.AppendLine(Environment.NewLine);
                        break;


                    default:
                        PlainTextInWord.Append(GetPlainText(section));
                        break;
                }
            }


            return PlainTextInWord.ToString();
        }
        */
    }
}
