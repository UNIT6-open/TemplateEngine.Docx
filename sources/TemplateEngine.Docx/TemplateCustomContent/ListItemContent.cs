using System.Collections.Generic;

namespace TemplateEngine.Docx
{
	public class ListItemContent:FieldContent
	{
		public ListItemContent()
        {
            
        }

		public ListItemContent(string name, string value) : base(name, value)
		{
		}
		public ListItemContent(string name, string value, IEnumerable<FieldContent> nestedfields)
			: base(name, value)
		{
			NestedFields = nestedfields;
		}

		public ListItemContent(string name, string value, params FieldContent[] nestedfields)
			: base(name, value)
        {
			NestedFields = nestedfields;
        }


		public IEnumerable<FieldContent> NestedFields { get; set; }
	}
}
