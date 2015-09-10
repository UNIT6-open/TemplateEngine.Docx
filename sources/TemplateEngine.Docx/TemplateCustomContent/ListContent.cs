using System.Collections.Generic;

namespace TemplateEngine.Docx
{
	public class ListContent
	{
		public ListContent()
        {
            
        }

        public ListContent(string name)
        {
            Name = name;
        }

        public ListContent(string name, IEnumerable<FieldContent> items)
            : this(name)
        {
            Items = items;
        }

		public ListContent(string name, params FieldContent[] items)
            : this(name)
        {
            Items = items;
        }

        public string Name { get; set; }
		public IEnumerable<FieldContent> Items { get; set; }

	
	}
}
