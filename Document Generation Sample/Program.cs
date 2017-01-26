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
            DirectoryInfo di = new DirectoryInfo(Environment.CurrentDirectory);
            string destinationFile = Path.Combine(di.Parent.Parent.FullName, "Sample Document.docx");
            string sourceFile = Path.Combine(di.Parent.Parent.FullName, "Sample Template.dotx");

            GenerateDocumentFromTemplate(sourceFile, destinationFile);

            string sSourceXML = Path.Combine(di.Parent.Parent.FullName, "Data.xml");
            UpdateDocument(destinationFile, sSourceXML);
        }

        static void GenerateDocumentFromTemplate(string inputPath, string outputPath)
        {
            MemoryStream documentStream;
            using (Stream stream = File.OpenRead(inputPath))
            {
                documentStream = new MemoryStream((int)stream.Length);
                CopyStream(stream, documentStream);
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
