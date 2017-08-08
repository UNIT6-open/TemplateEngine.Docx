using System;
using Newtonsoft.Json;

namespace TemplateEngine.Docx
{
	public interface IContentItem : IEquatable<IContentItem>
	{
		string Name { get; set; }

	    bool IsHidden { get; set; }
	}
}
