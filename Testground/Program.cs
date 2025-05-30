// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;

SearchAndReplace("file.docx", new Dictionary<string, string>
{
    { "{replacableText}", "I'm just here to say" },
    { "{differentReplacableText}", "world" },
});

static void SearchAndReplace(string path, Dictionary<string, string> replacedValues)
{
    var docAsArray = File.ReadAllBytes(path);

    using var stream = new MemoryStream();
    stream.Write(docAsArray, 0, docAsArray.Length);

    using (var doc = WordprocessingDocument.Open(stream, true))
    {
        string textContent;
        // Read the document XML
        using (var sr = new StreamReader(doc.MainDocumentPart.GetStream()))
        {
            textContent = sr.ReadToEnd();
        }

        // Replace text
        foreach (var item in replacedValues)
        {
            var regex = new Regex(item.Key);
            // Replace the text using regex
            textContent = regex.Replace(textContent, item.Value);
        }

        // Write the modified XML back
        using (var sw = new StreamWriter(doc.MainDocumentPart.GetStream(FileMode.Create)))
        {
            sw.Write(textContent);
        }

        doc.Save();
    }

    File.WriteAllBytes("fixedFile.docx", stream.ToArray());
}

return;
