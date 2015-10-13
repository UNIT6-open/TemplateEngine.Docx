using System.Collections.Generic;
using System.Linq;
namespace TemplateEngine.Docx
{
	public abstract class Container
	{
		protected Container(params IContentItem[] contentItems)
		{
			if (contentItems != null)
			{
				Lists = contentItems.OfType<ListContent>().ToList();
				Tables = contentItems.OfType<TableContent>().ToList();
				Fields = contentItems.OfType<FieldContent>().ToList();
			}
		}

		private IEnumerable<IContentItem> All
		{
			get
			{
				var result = new List<IContentItem>();
				if (Tables != null) result = result.Concat(Tables).ToList();
				if (Lists != null) result = result.Concat(Lists).ToList();
				if (Fields != null) result = result.Concat(Fields).ToList();

				return result;
			}
		}

		public ICollection<TableContent> Tables { get; set; }
		public ICollection<ListContent> Lists { get; set; }
		public ICollection<FieldContent> Fields { get; set; }

		public IContentItem GetContentItem(string name)
		{
			var allFields = All.ToList();
			if (allFields.Any(t => t.Name == name)) 
				return allFields.First(t => t.Name == name);
			
			return null;
		}

		public IEnumerable<string> FieldNames
		{
			get
			{
				return Tables == null
					? null
					: Tables.Select(t => t.Name).Concat(Tables.SelectMany(t => t.Rows.SelectMany(r => r.FieldNames)))
						.Concat(Lists == null
							? new List<string>()
							: Lists.Select(l => l.Name).Concat(Lists.SelectMany(l => l.FieldNames)))
						.Concat(Fields == null ? new List<string>() : Fields.Select(f => f.Name));
			}
		}

		protected Container AddField(string name, string value)
		{
			if (Fields == null) Fields = new List<FieldContent>();

			Fields.Add(new FieldContent(name, value));
			return this;
		}

		protected Container AddTable(TableContent table)
		{
			if (Tables == null) Tables = new List<TableContent>();

			Tables.Add(table);
			return this;
		}
		protected Container AddList(ListContent list)
		{
			if (Lists == null) Lists = new List<ListContent>();

			Lists.Add(list);
			return this;
		}
	}
}
