using System;

namespace TemplateEngine.Docx
{
	internal class ContentItemNameAttribute:Attribute
	{
		internal ContentItemNameAttribute(string name)
		{
			Name = name;
		}

		internal string Name { get; private set; }
	}
}
