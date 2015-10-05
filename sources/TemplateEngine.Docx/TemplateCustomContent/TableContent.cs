using System.Collections.Generic;
using System.Linq;

namespace TemplateEngine.Docx
{
    public class TableContent:IContentItem
	{
		public string Name { get; set; }
		public ICollection<TableRowContent> Rows { get; set; }

		public IEnumerable<string> FieldNames
		{
			get
			{
				return Rows == null ? new List<string>() : Rows.SelectMany(r => r.FieldNames).Distinct().ToList();
			}
		} 

		#region ctors
		public TableContent()
        {
            
        }

        public TableContent(string name)
        {
            Name = name;
        }

        public TableContent(string name, IEnumerable<TableRowContent> rows)
            : this(name)
        {
            Rows = rows.ToList();
        }

        public TableContent(string name, params TableRowContent[] rows)
            : this(name)
        {
            Rows = rows.ToList();
        }
		#endregion

		#region fluent

		public static TableContent Create(string name, params TableRowContent[] rows)
		{
			return new TableContent(name, rows);
		}

		public static TableContent Create(string name, List<TableRowContent> rows)
		{
			return new TableContent(name, rows);
		}

		public TableContent AddRow(params IContentItem[] contentItems)
		{
			if (Rows == null) Rows = new List<TableRowContent>();

			Rows.Add(new TableRowContent(contentItems));
			return this;
		}

		#endregion
		
    }
}
