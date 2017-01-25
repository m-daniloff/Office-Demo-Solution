using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Document_Generation_Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            string destinationFile = Path.Combine(Environment.CurrentDirectory, "Sample Document.docx");
            string sourceFile = Path.Combine(Environment.CurrentDirectory, "Sample Template.dotx");

            GenerateDocumentFromTemplate2(sourceFile, destinationFile);

            string sSourceXML = Path.Combine(Environment.CurrentDirectory, "Data.xml");
            UpdateDocument(destinationFile, sSourceXML);
        }

        static void GenerateDocumentFromTemplate2(string inputPath, string outputPath)
        {
            MemoryStream documentStream;
            using (Stream tplStream = File.OpenRead(inputPath))
            {
                documentStream = new MemoryStream((int)tplStream.Length);
                CopyStream(tplStream, documentStream);
                documentStream.Position = 0L;
            }

            using (WordprocessingDocument template = WordprocessingDocument.Open(documentStream, true))
            {
                template.ChangeDocumentType(DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
                MainDocumentPart mainPart = template.MainDocumentPart;
                mainPart.DocumentSettingsPart.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate",
                   new Uri(inputPath, UriKind.Absolute));

                mainPart.Document.Save();
            }


            File.WriteAllBytes(outputPath, documentStream.ToArray());
        }

        static void CopyStream(Stream source, Stream target)
        {
            if (source != null)
            {
                MemoryStream mstream = source as MemoryStream;
                if (mstream != null) mstream.WriteTo(target);
                else
                {
                    byte[] buffer = new byte[2048];
                    int length = buffer.Length, size;
                    while ((size = source.Read(buffer, 0, length)) != 0)
                        target.Write(buffer, 0, size);
                }
            }
        }

        static void UpdateDocument(string outputPath, string sSourceXML)
        {
            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(outputPath, true))
            {
                MainDocumentPart mainDocumentPart = wdDoc.MainDocumentPart;
                CustomXmlPart cusPart = mainDocumentPart.CustomXmlParts.First();
                cusPart.FeedData(new FileStream(sSourceXML, FileMode.Open, FileAccess.Read));
                mainDocumentPart.Document.Save();
            }
        }
    }
}
