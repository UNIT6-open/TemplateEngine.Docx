using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.Xml.Linq;

namespace TemplateEngine.Docx.Processors
{
	internal class ProcessContext
	{
		internal IDocumentContainer Document { get; private set; }
        internal Dictionary<int, int> LastNumIds { get; private set; }

		internal ProcessContext(IDocumentContainer document)
		{
			Document = document;
            LastNumIds = new Dictionary<int, int>();
			        
		}
	}
}
