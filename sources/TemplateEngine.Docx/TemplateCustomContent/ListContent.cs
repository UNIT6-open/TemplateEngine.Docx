using System;
using System.Collections.Generic;
using System.Linq;

namespace TemplateEngine.Docx
{
	public class ListContent
	{
		private IEnumerable<FieldContent> _items;

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

		public IEnumerable<FieldContent> Items
		{
			get { return _items; }
			set
			{
				if (value != null && value.Where(i => i != null).Select(i => i.Name).Distinct().Count() > 1)
				{
					throw new Exception(string.Format("Items in the same level must have same names."));
				}
				_items = value;
			}
		}

		public IEnumerable<string> FieldNames
		{
			get
			{
				return Items == null ? new List<string>() : GetFieldNames(Items);
			}
		}

		private List<string> GetFieldNames(IEnumerable<FieldContent> items)
		{
			var result = new List<string>();
			if (items == null) return null;

			foreach (var item in items)
			{
				if (item != null && !result.Contains(item.Name))
					result.Add(item.Name);

				if (item is ListItemContent)
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
			return result;
		}	
	}
}
