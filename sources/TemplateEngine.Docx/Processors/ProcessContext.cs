using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.Xml.Linq;

namespace TemplateEngine.Docx.Processors
{
	internal class ProcessContext
	{
        internal WordprocessingDocument WordDocument { get; private set; }
        internal Dictionary<int, int> LastNumIds { get; private set; }
		internal XDocument MainPart { get; private set; }
		internal XDocument NumberingPart { get; private set; }
		internal XDocument StylesPart { get; private set; }

		internal ProcessContext(WordprocessingDocument wordDocument, XDocument mainPart, XDocument numberingPart, XDocument stylesPart)
		{
            WordDocument = wordDocument;
            LastNumIds = new Dictionary<int, int>();
			MainPart = mainPart;
			NumberingPart = numberingPart;          
		}
	}
}
