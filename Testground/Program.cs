// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;

SearchAndReplace("file.docx", "{replacableText}", "REPLACED");

static void SearchAndReplace(string path, string regex, string value)
{
    var docAsArray = File.ReadAllBytes(path);

    using var stream = new MemoryStream();
    stream.Write(docAsArray, 0, docAsArray.Length);
    stream.Position = 0;

    using (var doc = WordprocessingDocument.Open(stream, true))
    {
        string originalText;
        // Read the document XML
        using (var sr = new StreamReader(doc.MainDocumentPart.GetStream()))
        {
            originalText = sr.ReadToEnd();
        }

        // Replace text
        var regexText = new Regex(regex);
        var newText = regexText.Replace(originalText, value);

        // Write the modified XML back
        using (var sw = new StreamWriter(doc.MainDocumentPart.GetStream(FileMode.Create)))
        {
            sw.Write(newText);
        }

        doc.Save();
    }

    File.WriteAllBytes("fixedFile.docx", stream.ToArray());
}

return;
