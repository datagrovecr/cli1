using Markdig;
using System;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using System.Text;

internal class Program
{
    static void Main(string[] args)
    {
        String md = File.ReadAllText("/Users/fabianvalverde/Documents/GitHub/cli1/test.md");
        String html = Markdown.ToHtml(md);

        const string filename = "/Users/fabianvalverde/Documents/GitHub/cli1/test.docx";

        if (File.Exists(filename)) File.Delete(filename);

        using (MemoryStream generatedDocument = new MemoryStream())
        {
            // Uncomment and comment the second using() to open an existing template document
            // instead of creating it from scratch.


            //using (var buffer = new FileStream("/Users/fabianvalverde/Documents/GitHub/cli1/template.docx", FileMode.Open, FileAccess.Read))
            //{
              //  buffer.CopyTo(generatedDocument);
            //}

            //using (WordprocessingDocument wordDocument = WordprocessingDocument.Open("/Users/fabianvalverde/Documents/GitHub/cli1/template.docx", true))
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.MainDocumentPart;
                if (mainPart == null)
                {
                    mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document();
                }

                HtmlConverter converter = new HtmlConverter(mainPart);

                converter.ParseHtml(html);
                mainPart.Document.Save();

                AssertThatOpenXmlDocumentIsValid(wordDocument);
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

        File.WriteAllText(outfile + ".all.txt", text);

        StringBuilder textBuilder = new StringBuilder();
        var parts = wordDoc.MainDocumentPart.Document.Descendants().FirstOrDefault();
        StyleDefinitionsPart styleDefinitionsPart = wordDoc.MainDocumentPart.StyleDefinitionsPart;
        if (parts != null)
        {
            foreach (var node in parts.ChildElements)
            {
                if (node is Paragraph)
                {
                    ProcessParagraph((Paragraph)node, textBuilder);
                    textBuilder.AppendLine("");
                }

                if (node is Table)
                {
                    ProcessTable((Table)node, textBuilder);
                }
            }
        }
        File.WriteAllText(outfile, textBuilder.ToString());
    }

    private static void ProcessTable(Table node, StringBuilder textBuilder)
    {
        foreach (var row in node.Descendants<TableRow>())
        {
            textBuilder.Append("| ");
            foreach (var cell in row.Descendants<TableCell>())
            {
                foreach (var para in cell.Descendants<Paragraph>())
                {
                    ProcessParagraph(para, textBuilder);
                }
                textBuilder.Append(" | ");
            }
            textBuilder.AppendLine("");
        }
    }

    private static void ProcessParagraph(Paragraph node, StringBuilder textBuilder)
    {

        foreach (var run in node.Descendants<Run>())
        {
            String prefix = "";
            if (run.RunProperties != null)
            {
                if (run.RunProperties.Bold != null)
                {
                    prefix += "*";
                }
                if (run.RunProperties.Italic != null)
                {
                    prefix += "_";
                }
            }
            textBuilder.Append(prefix + run.InnerText + prefix + " ");
            prefix = "";
            //text.GetAttributes();

        }
        textBuilder.Append("\n\n");
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







/*
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
 */