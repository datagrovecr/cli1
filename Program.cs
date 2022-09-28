using Markdig;
using System;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;

internal class Program
{
    static void Main(string[] args)
    {
        String md = File.ReadAllText("test.md");
        var html = Markdown.ToHtml(md);

        const string filename = "test.docx";

        if (File.Exists(filename)) File.Delete(filename);

        using (MemoryStream generatedDocument = new MemoryStream())
        {
            // Uncomment and comment the second using() to open an existing template document
            // instead of creating it from scratch.


            using (var buffer = new FileStream("template.docx", FileMode.Open, FileAccess.Read))
            {
                buffer.CopyTo(generatedDocument);
            }

            generatedDocument.Position = 0L;
            using (WordprocessingDocument package = WordprocessingDocument.Open(generatedDocument, true))
            //using (WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = package.MainDocumentPart;
                if (mainPart == null)
                {
                    mainPart = package.AddMainDocumentPart();
                    new Document(new Body()).Save(mainPart);
                }

                HtmlConverter converter = new HtmlConverter(mainPart);
                Body body = mainPart.Document.Body;

                converter.ParseHtml(html);
                mainPart.Document.Save();

                AssertThatOpenXmlDocumentIsValid(package);
            }

            File.WriteAllBytes(filename, generatedDocument.ToArray());
            md2docx(filename, filename + ".md");
        }


    }
    static void md2docx(String filename, String outfile)
    {
        WordprocessingDocument wordDoc = WordprocessingDocument.Open(filename, false);
        DocumentFormat.OpenXml.Wordprocessing.Body body
        = wordDoc.MainDocumentPart.Document.Body;
        var totaltext = body.InnerText;
        String text = totaltext;

        File.WriteAllText(outfile, text);
    }

    static void AssertThatOpenXmlDocumentIsValid(WordprocessingDocument wpDoc)
    {

        var validator = new OpenXmlValidator(FileFormatVersions.Office2010);
        var errors = validator.Validate(wpDoc);

        if (!errors.GetEnumerator().MoveNext())
            return;

        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("The document doesn't look 100% compatible with Office 2010.\n");

        Console.ForegroundColor = ConsoleColor.Gray;
        foreach (ValidationErrorInfo error in errors)
        {
            Console.Write("{0}\n\t{1}", error.Path.XPath, error.Description);
            Console.WriteLine();
        }

        Console.ReadLine();
    }
}