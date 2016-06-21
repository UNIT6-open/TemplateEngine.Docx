using System;
using System.Collections.Generic;

namespace TemplateEngine.Docx
{
    public class TableRowContent:Container, IEquatable<TableRowContent>
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

		#region Equals

		public bool Equals(TableRowContent other)
	    {
		    return base.Equals(other);
	    }

	    public override int GetHashCode()
	    {
		    return base.GetHashCode();
		}

		#endregion
	}
}
