using System.Collections.Generic;
using System.Linq;

namespace TemplateEngine.Docx
{
	public class ListItemContent:Container
	{
		public ICollection<ListItemContent> NestedFields { get; set; }

		#region ctors
		public ListItemContent(params IContentItem[] contentItems):base(contentItems)
        {
            
        }

		public ListItemContent(IEnumerable<ListItemContent> nestedfields, params IContentItem[] contentItems)
			: base(contentItems)
		{
			NestedFields = nestedfields.ToList();
		}

		public ListItemContent(string name, string value)
		{
			Fields = new List<FieldContent> {new FieldContent {Name = name, Value = value}};
			NestedFields = new List<ListItemContent>();
		}
		
		public ListItemContent(string name, string value, IEnumerable<ListItemContent> nestedfields)
		{
			Fields = new List<FieldContent>{new FieldContent{Name = name, Value = value}};
			NestedFields = nestedfields.ToList();
		}

		public ListItemContent(string name, string value, params ListItemContent[] nestedfields)
		{
			Fields = new List<FieldContent> {new FieldContent {Name = name, Value = value}};
			NestedFields = nestedfields.ToList();

		}

		#endregion

		#region fluent

		public static ListItemContent Create(string name, string value, params ListItemContent[] nestedfields)
		{
			return new ListItemContent(name, value, nestedfields);
		}

		public static ListItemContent Create(string name, string value, List<ListItemContent> nestedfields)
		{
			return new ListItemContent(name, value, nestedfields);
		}
		public ListItemContent AddField(string name, string value)
		{
			if (Fields == null) Fields = new List<FieldContent>();

			Fields.Add(new FieldContent(name, value));
			return this;
		}

		public ListItemContent AddTable(TableContent table)
		{
			if (Tables == null) Tables = new List<TableContent>();

			Tables.Add(table);
			return this;
		}
		public ListItemContent AddList(ListContent list)
		{
			if (Lists == null) Lists = new List<ListContent>();

			Lists.Add(list);
			return this;
		}

		public ListItemContent AddNestedItem(ListItemContent nestedItem)
		{
			if (NestedFields == null) NestedFields = new List<ListItemContent>();

			NestedFields.Add(nestedItem);
			return this;
		}

		#endregion

	}
}
