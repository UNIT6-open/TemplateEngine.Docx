using System.Collections.Generic;
using System.Xml.Linq;

namespace TemplateEngine.Docx.Processors
{
	internal interface IProcessor
	{

		IProcessor SetRemoveContentControls(bool isNeedToRemove);
		ProcessResult FillContent(XElement contentControl, IEnumerable<IContentItem> items);
	}
}