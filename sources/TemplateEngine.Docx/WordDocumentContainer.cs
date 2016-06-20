using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace TemplateEngine.Docx
{
	internal class WordDocumentContainer:IDisposable
	{
		private readonly WordprocessingDocument _wordDocument;

		internal XDocument MainDocumentPart { get; private set; }
		internal XDocument NumberingPart { get; private set; }
		internal XDocument StylesPart { get; private set; }
		internal IEnumerable<ImagePart> ImagesPart { get; private set; }

		internal WordDocumentContainer(WordprocessingDocument wordDocument)
		{
			_wordDocument = wordDocument;

			MainDocumentPart = LoadPart(_wordDocument.MainDocumentPart);
			NumberingPart = LoadPart(_wordDocument.MainDocumentPart.NumberingDefinitionsPart);
			StylesPart = LoadPart(_wordDocument.MainDocumentPart.StyleDefinitionsPart);

			ImagesPart = _wordDocument.MainDocumentPart.ImageParts;
			
		}
		internal WordDocumentContainer(XDocument templateSource, XDocument stylesPart = null, XDocument numberingPart = null, IEnumerable<ImagePart> imagesPart = null)
		{
			MainDocumentPart = templateSource;
			NumberingPart = numberingPart;
			StylesPart = stylesPart;
			ImagesPart = imagesPart;
		}

		internal OpenXmlPart GetPartById(string partIdentifier)
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
		internal void RemovePartById(string partIdentifier)
		{
			if (_wordDocument == null)
				return;

			var part = GetPartById(partIdentifier);
			if (part != null)
			{
				_wordDocument.MainDocumentPart.DeletePart(part);
			}
		}
		internal string AddImagePart(byte[] bytes)
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
	}
}
