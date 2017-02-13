using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace TemplateEngine.Docx
{
	internal class FooterWordDocumentContainer:NestedWordDocumentContainer
	{
        internal FooterWordDocumentContainer(string identifier, WordprocessingDocument document)
            : base(identifier, document)
        {
            ImagesPart = GetImagesPart();
        }

		public override string AddImagePart(byte[] bytes)
		{
			if (_document == null)
				return null;

            var imagePart = ((FooterPart)_document).AddImagePart(ImagePartType.Jpeg);

			using (var stream = new MemoryStream(bytes))
			{
				imagePart.FeedData(stream);
			}

			return _document.GetIdOfPart(imagePart);
		}

	    private IEnumerable<ImagePart> GetImagesPart()
	    {
            return ((FooterPart)_document).ImageParts;
	    }
	}
}
