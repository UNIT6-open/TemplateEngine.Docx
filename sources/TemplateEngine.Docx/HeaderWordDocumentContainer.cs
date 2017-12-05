using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace TemplateEngine.Docx
{
	internal class HeaderWordDocumentContainer:NestedWordDocumentContainer
	{
        internal HeaderWordDocumentContainer(string identifier, WordprocessingDocument document)
            : base(identifier, document)
        {
            ImagesPart = GetImagesPart();
        }

		public override string AddImagePart(byte[] bytes)
		{
			if (_document == null)
				return null;

			var imagePart = (_document as HeaderPart)?.AddImagePart(ImagePartType.Jpeg);
			if (imagePart == null) return string.Empty;

			using (var stream = new MemoryStream(bytes))
			{
				imagePart.FeedData(stream);
			}

			return _document.GetIdOfPart(imagePart);
		}

        private IEnumerable<ImagePart> GetImagesPart()
        {
            return ((HeaderPart)_document).ImageParts;
        }
	}
}
