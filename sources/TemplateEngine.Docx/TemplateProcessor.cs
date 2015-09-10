using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;


namespace TemplateEngine.Docx
{
    public class TemplateProcessor : IDisposable
    {
        public readonly XDocument Document;
        private readonly WordprocessingDocument _wordDocument;
	    private bool _isNeedToRemoveContentControls = false;

        public TemplateProcessor(string fileName)
        {
            _wordDocument = WordprocessingDocument.Open(fileName, true);
			
            var xdoc = _wordDocument.MainDocumentPart.Annotation<XDocument>();
            if (xdoc == null)
            {
                using (var str = _wordDocument.MainDocumentPart.GetStream())
                using (var streamReader = new StreamReader(str))
                using (var xr = XmlReader.Create(streamReader))
                    xdoc = XDocument.Load(xr);

                _wordDocument.MainDocumentPart.AddAnnotation(xdoc);
            }
            
            Document = xdoc;
        }

	    public TemplateProcessor SetRemoveContentControls(bool isNeedToRemove)
	    {
		    _isNeedToRemoveContentControls = isNeedToRemove;
		    return this;
	    }

        public void SaveChanges()
        {
            if (Document == null) return;

            // Serialize the XDocument object back to the package.
            using (var xw = XmlWriter.Create(_wordDocument.MainDocumentPart.GetStream (FileMode.Create, FileAccess.Write)))
            {
                Document.Save(xw);
            }

            _wordDocument.Close();
        }

        public TemplateProcessor(XDocument templateSource)
        {
            Document = templateSource;
        }

        public TemplateProcessor FillContent(Content content)
        {

            var fillFieldsErrors = FillFields(content);          
            var fillTablesErrors = FillTables(content);

	        var errors = fillFieldsErrors.Concat(fillTablesErrors).ToList();

            AddErrors(errors);

            return this;
        }

	    // Filling a tables
		private IEnumerable<string> FillTables(Content content)
		{
			var errors = new List<string>();

			if (content.Tables == null) return errors;
			foreach (var table in content.Tables)
			{
				errors.AddRange(new TableProcessor(Document).SetRemoveContentControls(_isNeedToRemoveContentControls).FillTableContent(table));
			}

			return errors;
		}

		// Filling a fields
		private IEnumerable<string> FillFields(Content content)
	    {
		    var errors = new List<string>();
		    

		    if (content.Fields != null)
		    {
			    foreach (var field in content.Fields)
			    {
				    var fieldsContentControl = Document.Root
					    .Element(W.body)
					    .Descendants(W.sdt)
					    .Where(sdt => field.Name == sdt.Element(W.sdtPr).Element(W.tag).Attribute(W.val).Value)
					    .ToList();

				    // If there isn't a field with that name, add an error to the error string,
				    // and continue with next field.
				    if (!fieldsContentControl.Any())
				    {
					    errors.Add(String.Format("Field Content Control '{0}' not found.",
						    field.Name));
					    continue;
				    }

				    // Set content control value to the new value
				    foreach (var fieldContentControl in fieldsContentControl)
				    {
					    fieldContentControl.ReplaceContentControlWithNewValue(field.Value, _isNeedToRemoveContentControls);
				    }
			    }
		    }
		    return errors;
	    }

	    // Add any errors as red text on yellow at the beginning of the document.
	    private void AddErrors(List<string> errors)
	    {
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
	    }

	    public void Dispose()
        {
            if (_wordDocument == null) return;

            _wordDocument.Dispose();
        }

	 
    }
}
