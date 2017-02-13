using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using TemplateEngine.Docx.Errors;
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

	    public IEnumerable<ImagePart> ImagesPart { get { return _wordDocument.ImagesPart; } }

		public Dictionary<string, XDocument> HeaderParts {
			get
			{
				return _wordDocument.HeaderParts
					.Select(x => new {x.Identifier, x.MainDocumentPart})
					.ToDictionary(x => x.Identifier, y => y.MainDocumentPart);
			}
		}

        public Dictionary<string, IEnumerable<ImagePart>> HeaderImagesParts
        {
            get
            {
                return _wordDocument.HeaderParts
                    .Select(x => new { x.Identifier, x.ImagesPart })
                    .ToDictionary(x => x.Identifier, y => y.ImagesPart);
            }
        }

		public Dictionary<string, XDocument> FooterParts
		{
			get
			{
				return _wordDocument.FooterParts
					.Select(x => new { x.Identifier, x.MainDocumentPart })
					.ToDictionary(x => x.Identifier, y => y.MainDocumentPart);
			}
		}

        public Dictionary<string, IEnumerable<ImagePart>> FooterImagesParts
        {
            get
            {
                return _wordDocument.FooterParts
                    .Select(x => new { x.Identifier, x.ImagesPart })
                    .ToDictionary(x => x.Identifier, y => y.ImagesPart);
            }
        }

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
			var processor = new ContentProcessor(
				new ProcessContext(_wordDocument))
				.SetRemoveContentControls(_isNeedToRemoveContentControls);

			var processResult = processor.FillContent(Document.Root.Element(W.body), content);

			if (_wordDocument.HasFooters)
			{
				foreach (var footer in _wordDocument.FooterParts)
				{
					var footerProcessor = new ContentProcessor(
				        new ProcessContext(footer))
				        .SetRemoveContentControls(_isNeedToRemoveContentControls);

					var footerProcessResult = footerProcessor.FillContent(footer.MainDocumentPart.Element(W.footer), content);
					processResult.Merge(footerProcessResult);
				}
			}

			if (_wordDocument.HasHeaders)
			{
				foreach (var header in _wordDocument.HeaderParts)
				{
					var headerProcessor = new ContentProcessor(
			            new ProcessContext(header))
			            .SetRemoveContentControls(_isNeedToRemoveContentControls);

					var headerProcessResult = headerProcessor.FillContent(header.MainDocumentPart.Element(W.header), content);
					processResult.Merge(headerProcessResult);
				}
			}
			
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
	    private void AddErrors(IList<IError> errors)
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
							    new XElement(W.t, s.Message)))));
	    }

	    public void Dispose()
        {
			if (_wordDocument != null)
				_wordDocument.Dispose();
        }
    }
}
