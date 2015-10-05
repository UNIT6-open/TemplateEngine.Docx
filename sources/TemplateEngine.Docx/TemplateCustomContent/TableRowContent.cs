using System.Collections.Generic;

namespace TemplateEngine.Docx
{
    public class TableRowContent:Container
    {
        public TableRowContent()
        {
            
        }

		public TableRowContent(params IContentItem[] contentItems)
			: base(contentItems)
		{
			
		}

	    public TableRowContent(List<FieldContent> fields)
	    {
		    Fields = fields;
	    }
    }
}
