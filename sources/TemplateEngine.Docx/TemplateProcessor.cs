using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using TemplateEngine.Docx.Processors;


namespace TemplateEngine.Docx
{
    public class TemplateProcessor : IDisposable
    {
	    private readonly WordDocumentContainer _wordDocument;
	    private bool _isNeedToRemoveContentControls;
	    private bool _isNeedToNoticeAboutErrors;

	    public XDocument Document { get { return _wordDocument.MainDocumentPart; } }

	    public XDocument NumberingPart { get { return _wordDocument.NumberingPart; } }

	    public XDocument StylesPart { get { return _wordDocument.StylesPart; } }

	    private TemplateProcessor(WordprocessingDocument wordDocument)
        {
            _wordDocument = new WordDocumentContainer(wordDocument);
			_isNeedToNoticeAboutErrors = true;
        }

        public TemplateProcessor(string fileName) : this(WordprocessingDocument.Open(fileName, true))
        {
        }

        public TemplateProcessor(Stream stream) : this(WordprocessingDocument.Open(stream, true))
        {
        }

        public TemplateProcessor(XDocument templateSource, XDocument stylesPart = null, XDocument numberingPart = null)
		{
			_isNeedToNoticeAboutErrors = true;
			_wordDocument = new WordDocumentContainer(templateSource, stylesPart, numberingPart);
		}

	    public TemplateProcessor SetRemoveContentControls(bool isNeedToRemove)
	    {
		    _isNeedToRemoveContentControls = isNeedToRemove;
		    return this;
	    }

	    public TemplateProcessor SetNoticeAboutErrors(bool isNeedToNotice)
	    {
			_isNeedToNoticeAboutErrors = isNeedToNotice;
		    return this;
	    }

		public TemplateProcessor FillContent(Content content)
        {
			var processResult =
		        new ContentProcessor(
					new ProcessContext(_wordDocument))
					.SetRemoveContentControls(_isNeedToRemoveContentControls)
			        .FillContent(Document.Root.Element(W.body), content);

			if (_isNeedToNoticeAboutErrors)
				AddErrors(processResult.Errors);

            return this;
        }
		
		public void SaveChanges()
		{
			_wordDocument.SaveChanges();
		}

		/// <summary>
		/// Adds a list of errors as red text on yellow at the beginning of the document.
		/// </summary>
		/// <param name="errors">List of errors.</param>
	    private void AddErrors(IList<string> errors)
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
			if (_wordDocument != null)
				_wordDocument.Dispose();
        }
    }
}
