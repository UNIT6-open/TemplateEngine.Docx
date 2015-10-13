using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using DocumentFormat.OpenXml.Office.CustomUI;

namespace TemplateEngine.Docx
{
	public class ListContent:IContentItem
	{
		public string Name { get; set; }

		public ICollection<ListItemContent> Items { get; set; }

		public IEnumerable<string> FieldNames
		{
			get
			{
				return Items == null ? new List<string>() : GetFieldNames(Items);
			}
		}


		#region ctors
		public ListContent()
        {
            
        }

        public ListContent(string name)
        {
            Name = name;
        }

		public ListContent(string name, IEnumerable<ListItemContent> items)
            : this(name)
        {
            Items = items.ToList();
        }

		public ListContent(string name, params ListItemContent[] items)
            : this(name)
        {
			Items = items.ToList();
        }

		#endregion

		#region fluent

		public static ListContent Create(string name, params ListItemContent[] items)
        {
			return new ListContent(name, items);
        }
		public static ListContent Create(string name, IEnumerable<ListItemContent> items)
        {
			return new ListContent(name, items);
        }

		public ListContent AddItem(ListItemContent item)
		{
			if (Items == null) Items = new Collection<ListItemContent>();
			Items.Add(item);
			return this;
		}
		public ListContent AddItem(params IContentItem[] contentItems)
		{
			if (Items == null) Items = new Collection<ListItemContent>();
			Items.Add(new ListItemContent(contentItems));
			return this;
		}

		#endregion

		#region IContentItem implementation
		private List<string> GetFieldNames(IEnumerable<ListItemContent> items)
		{
			var result = new List<string>();
			if (items == null) return null;

			foreach (var item in items)
			{
				if (item != null)
				{
					foreach (var fieldName in item.FieldNames.Where(fieldName => !result.Contains(fieldName)))
					{
						result.Add(fieldName);
					}

					if (item.NestedFields != null)
					{
						var listItem = item as ListItemContent;
						if (listItem.NestedFields != null)
						{
							var nestedFieldNames = GetFieldNames(listItem.NestedFields);
							foreach (var nestedFieldName in nestedFieldNames)
							{
								if (!result.Contains(nestedFieldName))
									result.Add(nestedFieldName);
							}
						}
					}

				}
			}			
			return result;
		}
		#endregion
	}
}
