using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace TemplateEngine.Docx
{
	internal abstract class NestedWordDocumentContainer:IDisposable, IDocumentContainer
	{
		internal string Identifier { get; private set; }

		private readonly WordprocessingDocument _mainWordDocument;
	    protected readonly OpenXmlPart _document;
		public XDocument MainDocumentPart { get; private set; }
		public XDocument NumberingPart { get; private set; }
		public XDocument StylesPart { get; private set; }
		public OpenXmlPart GetPartById(string partIdentifier)
		{
			return _document.GetPartById(partIdentifier);
		}

		public void RemovePartById(string partIdentifier)
		{
			_document.DeletePart(GetPartById(partIdentifier));
		}

	    public abstract string AddImagePart(byte[] bytes);

		public IEnumerable<ImagePart> ImagesPart { get; protected set; }

		protected NestedWordDocumentContainer(string identifier, WordprocessingDocument document)
		{
			Identifier = identifier;
			_mainWordDocument = document;
			_document = GetPart();

			MainDocumentPart = LoadPart(_document);
            NumberingPart = LoadPart(document.MainDocumentPart.NumberingDefinitionsPart);
            StylesPart = LoadPart(document.MainDocumentPart.StyleDefinitionsPart);
		}

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

		internal OpenXmlPart GetPart()
		{
			if (_mainWordDocument == null)
				return null;

			try
			{
				return _mainWordDocument.MainDocumentPart.GetPartById(Identifier);
			}
			catch (ArgumentOutOfRangeException)
			{
				return null;
			}
		}

		internal void Save()
		{
			using (var xw = XmlWriter.Create(GetPart().GetStream(FileMode.Create, FileAccess.Write)))
			{
				MainDocumentPart.Save(xw);
			}
		}
		
		#region IDisposable
		public void Dispose()
		{
		}

		#endregion

	}
}
