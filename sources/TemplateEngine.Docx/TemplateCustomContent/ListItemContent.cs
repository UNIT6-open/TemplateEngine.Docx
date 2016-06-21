using System;
using System.Collections.Generic;
using System.Linq;

namespace TemplateEngine.Docx
{
	public class ListItemContent : Container, IEquatable<ListItemContent>
	{
		public ICollection<ListItemContent> NestedFields { get; set; }

		#region ctors

		public ListItemContent()
		{
			
		}

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

		#region Fluent

		public static ListItemContent Create(string name, string value, params ListItemContent[] nestedfields)
		{
			return new ListItemContent(name, value, nestedfields);
		}

		public static ListItemContent Create(string name, string value, List<ListItemContent> nestedfields)
		{
			return new ListItemContent(name, value, nestedfields);
		}
		public new ListItemContent AddField(string name, string value)
		{
			return (ListItemContent)base.AddField(name, value);
		}

		public new ListItemContent AddTable(TableContent table)
		{
			return (ListItemContent)base.AddTable(table);
		}
		public new ListItemContent AddList(ListContent list)
		{
			return (ListItemContent)base.AddList(list);
		}

		public ListItemContent AddNestedItem(ListItemContent nestedItem)
		{
			if (NestedFields == null) NestedFields = new List<ListItemContent>();

			NestedFields.Add(nestedItem);
			return this;
		}

		public ListItemContent AddNestedItem(IContentItem nestedItem)
		{
			if (NestedFields == null) NestedFields = new List<ListItemContent>();

			NestedFields.Add(new ListItemContent(nestedItem));
			return this;
		}

		#endregion

		#region Equals
		public bool Equals(ListItemContent other)
		{
			if (other == null) return false;

			var equals = base.Equals(other);
	
			if (NestedFields != null)
				return equals && NestedFields.SequenceEqual(other.NestedFields);

			return equals;
		}
		public override int GetHashCode()
		{
			var nestedHc = 0;
			
			nestedHc = NestedFields.Aggregate(nestedHc, (current, p) => current ^ p.GetHashCode());
			var baseHc = base.GetHashCode();

			return new { baseHc, nestedHc }.GetHashCode();
		}

		#endregion
	}
}
