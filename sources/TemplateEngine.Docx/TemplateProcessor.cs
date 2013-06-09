using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace TemplateEngine.Docx
{
    public class TemplateProcessor
    {
        public readonly XDocument Document;
        private readonly WordprocessingDocument wordDocument;

        public TemplateProcessor(string fileName)
        {
            wordDocument = WordprocessingDocument.Open(fileName, true);

            var xdoc = wordDocument.MainDocumentPart.Annotation<XDocument>();
            if (xdoc == null)
            {
                using (Stream str = wordDocument.MainDocumentPart.GetStream())
                using (StreamReader streamReader = new StreamReader(str))
                using (XmlReader xr = XmlReader.Create(streamReader))
                    xdoc = XDocument.Load(xr);

                wordDocument.MainDocumentPart.AddAnnotation(xdoc);
            }

            Document = xdoc;
        }

        public void SaveChanges()
        {
            if (Document == null) return;

            // Serialize the XDocument object back to the package.
            using (XmlWriter xw = XmlWriter.Create(wordDocument.MainDocumentPart.GetStream (FileMode.Create, FileAccess.Write)))
            {
                Document.Save(xw);
            }
        }

        public TemplateProcessor(XDocument templateSource)
        {
            this.Document = templateSource;
        }

        public TemplateProcessor FillContent(Content content)
        {
            List<string> errors = new List<string>();

            // Filling a fields
            if (content.Fields != null)
            {
                foreach (var field in content.Fields)
                {
                    var fieldContentControl = Document.Root
                        .Element(W.body)
                        .Descendants(W.sdt)
                        .Where(sdt => field.Name == sdt.Element(W.sdtPr).Element(W.tag).Attribute(W.val).Value)
                        .FirstOrDefault();

                    // If there isn't a field with that name, add an error to the error string,
                    // and continue with next field.
                    if (fieldContentControl == null)
                    {
                        errors.Add(String.Format("Field Content Control '{0}' not found.",
                            field.Name));
                        continue;
                    }

                    // Set content control value to the new value
                    fieldContentControl
                        .Element(W.sdtContent)
                        .Descendants(W.t)
                        .FirstOrDefault()
                        .Value = field.Value;
                }
            }
            
            // Filling a tables
            if (content.Tables != null)
            {
                foreach (var table in content.Tables)
                {
                    // Find the content control with Table Name
                    var listName = table.Name;
                    var tableContentControl = Document.Root
                        .Element(W.body)
                        .Elements(W.sdt)
                        .Where(sdt => listName == sdt.Element(W.sdtPr).Element(W.tag).Attribute(W.val).Value)
                        .FirstOrDefault();

                    // If there isn't a table with that name, add an error to the error string,
                    // and continue with next table.
                    if (tableContentControl == null)
                    {
                        errors.Add(String.Format("Table Content Control '{0}' not found.",
                            listName));
                        continue;
                    }

                    // If the table doesn't contain content controls in cells, then error and continue with next table.
                    var cellContentControl = tableContentControl
                        .Descendants(W.sdt)
                        .FirstOrDefault();
                    if (cellContentControl == null)
                    {
                        errors.Add(String.Format(
                            "Table Content Control '{0}' doesn't contain content controls in cells.",
                            listName));
                        continue;
                    }

                    // Determine the element for the row that contains the content controls.
                    // This is the prototype for the rows that the code will generate from data.
                    var prototypeRow = tableContentControl
                        .Descendants(W.sdt)
                        .Ancestors(W.tr)
                        .FirstOrDefault();

                    // Create a list of new rows to be inserted into the document.  Because this
                    // is a document centric transform, this is written in a non-functional
                    // style, using tree modification.
                    var newRows = new List<XElement>();
                    foreach (var row in table.Rows)
                    {
                        // Clone the prototypeRow into newRow.
                        var newRow = new XElement(prototypeRow);

                        // Create new rows that will contain the data that was passed in to this
                        // method in the XML tree.
                        foreach (var sdt in newRow.Descendants(W.sdt).ToList())
                        {
                            // Get fieldName from the content control tag.
                            string fieldName = sdt
                                .Element(W.sdtPr)
                                .Element(W.tag)
                                .Attribute(W.val)
                                .Value;

                            // Get the new value out of contentControlValues.
                            var newValueElement = row
                                .Fields
                                .Where(f => f.Name == fieldName)
                                .FirstOrDefault();

                            // Generate error message if the new value doesn't exist.
                            if (newValueElement == null)
                            {
                                errors.Add(String.Format(
                                    "Table '{0}', Field '{1}' value isn't specified.",
                                    listName, fieldName));
                                continue;
                            }

                            // Set content control value th the new value
                            sdt.Element(W.sdtContent)
                                .Descendants(W.t)
                                .FirstOrDefault()
                                .Value = newValueElement.Value;
                        }

                        // Add the newRow to the list of rows that will be placed in the newly
                        // generated table.
                        newRows.Add(newRow);
                    }

                    // Remove the prototype row and add all of the newly constructed rows.
                    prototypeRow.AddAfterSelf(newRows);
                    prototypeRow.Remove();

                }
            }

            // Add any errors as red text on yellow at the beginning of the document.
            if (errors.Any())
                Document.Root
                    .Element(W.body)
                    .AddFirst(errors.Select(s =>
                        new XElement(W.p,
                            new XElement(W.r,
                                new XElement(W.rPr,
                                    new XElement(W.color,
                                        new XAttribute(W.val, "red")),
                                    new XElement(W.sz,
                                        new XAttribute(W.val, "28")),
                                    new XElement(W.szCs,
                                        new XAttribute(W.val, "28")),
                                    new XElement(W.highlight,
                                        new XAttribute(W.val, "yellow"))),
                                new XElement(W.t, s)))));

            return this;
        }
    }
}
