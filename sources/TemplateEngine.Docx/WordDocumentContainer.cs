using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace TemplateEngine.Docx
{
	internal class WordDocumentContainer : IDisposable, IDocumentContainer
	{
		private readonly WordprocessingDocument _wordDocument;

		public XDocument MainDocumentPart { get; private set; }
		public XDocument NumberingPart { get; private set; }
		public XDocument StylesPart { get; private set; }
		internal List<NestedWordDocumentContainer> HeaderParts { get; private set; }
		internal List<NestedWordDocumentContainer> FooterParts { get; private set; }
		internal IEnumerable<ImagePart> ImagesPart { get; private set; }

		internal bool HasHeaders
		{
			get { return HeaderParts != null && HeaderParts.Any(); }
		}
		internal bool HasFooters
		{
			get { return FooterParts != null && FooterParts.Any(); }
		}

		internal WordDocumentContainer(WordprocessingDocument wordDocument)
		{
			_wordDocument = wordDocument;

			MainDocumentPart = LoadPart(_wordDocument.MainDocumentPart);
			NumberingPart = LoadPart(_wordDocument.MainDocumentPart.NumberingDefinitionsPart);
			StylesPart = LoadPart(_wordDocument.MainDocumentPart.StyleDefinitionsPart);

			ImagesPart = _wordDocument.MainDocumentPart.ImageParts;

			HeaderParts = LoadHeaders(_wordDocument.MainDocumentPart.HeaderParts);
			FooterParts = LoadFooters(_wordDocument.MainDocumentPart.FooterParts);

		}
		internal WordDocumentContainer(XDocument templateSource, XDocument stylesPart = null, XDocument numberingPart = null, IEnumerable<ImagePart> imagesPart = null)
		{
			MainDocumentPart = templateSource;
			NumberingPart = numberingPart;
			StylesPart = stylesPart;
			ImagesPart = imagesPart;
		}

		public OpenXmlPart GetPartById(string partIdentifier)
		{
			if (_wordDocument == null)
				return null;

			try
			{
				return _wordDocument.MainDocumentPart.GetPartById(partIdentifier);
			}
			catch (ArgumentOutOfRangeException)
			{
				return null;
			}
		}
		public void RemovePartById(string partIdentifier)
		{
			if (_wordDocument == null)
				return;

			var part = GetPartById(partIdentifier);
			if (part != null)
			{
				_wordDocument.MainDocumentPart.DeletePart(part);
			}
		}
		public string AddImagePart(byte[] bytes)
		{
			if (_wordDocument == null)
				return null;

			var imagePart = _wordDocument.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);

			using (var stream = new MemoryStream(bytes))
			{
				imagePart.FeedData(stream);
			}
				
			return _wordDocument.MainDocumentPart.GetIdOfPart(imagePart);
		}

		internal void SaveChanges()
		{
			if (MainDocumentPart == null) return;

			// Serialize the XDocument object back to the package.
			using (var xw = XmlWriter.Create(_wordDocument.MainDocumentPart.GetStream(FileMode.Create, FileAccess.Write)))
			{
				MainDocumentPart.Save(xw);
			}

			if (NumberingPart != null)
			{
				// Serialize the XDocument object back to the package.
				using (var xw = XmlWriter.Create(_wordDocument.MainDocumentPart.NumberingDefinitionsPart.GetStream(FileMode.Create,
							FileAccess.Write)))
				{
					NumberingPart.Save(xw);
				}
			}

			foreach (var footer in FooterParts)
			{
				footer.Save();
			}
			foreach (var header in HeaderParts)
			{
				header.Save();
			}

			_wordDocument.Close();
		}

		#region IDisposable
		public void Dispose()
		{
			if (_wordDocument != null)
				_wordDocument.Dispose();
		}

		#endregion

		private XDocument LoadPart(OpenXmlPart source)
		{
			if (source == null) return null;

			var part = source.Annotation<XDocument>();
			if (part != null) return part;

			using (var str = source.GetStream())
			using (var streamReader = new StreamReader(str))
			using (var xr = XmlReader.Create(streamReader))
				part = XDocument.Load(xr);

			return part;
		}

		private List<NestedWordDocumentContainer> LoadHeaders(IEnumerable<OpenXmlPart> partsList)
		{
			return partsList
				.Select(part => 
					new HeaderWordDocumentContainer(
						_wordDocument.MainDocumentPart.GetIdOfPart(part),
						_wordDocument))
                .Cast<NestedWordDocumentContainer>()
				.ToList();
		}

        private List<NestedWordDocumentContainer> LoadFooters(IEnumerable<OpenXmlPart> partsList)
        {
            return partsList
                .Select(part =>
                    new FooterWordDocumentContainer(
                        _wordDocument.MainDocumentPart.GetIdOfPart(part),
                        _wordDocument))
                .Cast<NestedWordDocumentContainer>()
                .ToList();
        }
	}
}
