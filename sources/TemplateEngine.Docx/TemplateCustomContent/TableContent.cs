using System;
using System.Collections.Generic;
using System.Linq;

namespace TemplateEngine.Docx
{
	[ContentItemName("Table")]
    public class TableContent : HiddenContent<TableContent>, IContentItem, IEquatable<TableContent>
    {
        public ICollection<TableRowContent> Rows { get; set; }

		public IEnumerable<string> FieldNames
		{
			get
			{
				return Rows?.SelectMany(r => r.FieldNames).Distinct().ToList() ?? new List<string>();
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

		#region Fluent

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

        #region Equals

        public bool Equals(TableContent other)
	    {
			if (other == null) return false;
            
			return Name.Equals(other.Name) &&
			   Rows.SequenceEqual(other.Rows);
	    }

	    public override bool Equals(IContentItem other)
	    {
		    if (!(other is TableContent)) return false;

		    return Equals((TableContent) other);
	    }

	    public override int GetHashCode()
		{
			var hc = 0;
			if (Rows != null)
				hc = Rows.Aggregate(hc, (current, p) => current ^ p.GetHashCode());

			return new { Name, hc }.GetHashCode();
		}

		#endregion
	}
}
