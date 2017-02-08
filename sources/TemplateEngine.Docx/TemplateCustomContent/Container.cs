using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

namespace TemplateEngine.Docx
{
	[JsonObject]
	public abstract class Container:IEnumerable<IContentItem>, IEquatable<Container>
	{
		protected Container()
		{

                Repeats = new List<RepeatContent>();
                Lists = new List<ListContent>();
				Tables = new List<TableContent>();
				Fields = new List<FieldContent>();
				Images = new List<ImageContent>();
            
		}
		protected Container(params IContentItem[] contentItems)
		{
			if (contentItems != null)
			{
                Repeats = contentItems.OfType<RepeatContent>().ToList();
                Lists = contentItems.OfType<ListContent>().ToList();
				Tables = contentItems.OfType<TableContent>().ToList();
				Fields = contentItems.OfType<FieldContent>().ToList();
                Images = contentItems.OfType<ImageContent>().ToList();
            }
		}

		protected IEnumerable<IContentItem> All
		{
			get
			{
				var result = new List<IContentItem>();

                if (Repeats != null) result = result.Concat(Repeats).ToList();
                if (Tables != null) result = result.Concat(Tables).ToList();
				if (Lists != null) result = result.Concat(Lists).ToList();
				if (Fields != null) result = result.Concat(Fields).ToList();
                if (Images != null) result = result.Concat(Images).ToList();

                return result;
			}
		}

        public ICollection<RepeatContent> Repeats { get; set; }

        public ICollection<TableContent> Tables { get; set; }

		public ICollection<ListContent> Lists { get; set; }

		public ICollection<FieldContent> Fields { get; set; }

		public ICollection<ImageContent> Images { get; set; }

		
        public IContentItem GetContentItem(string name)
        {
	        return All.FirstOrDefault(t => t.Name == name);
        }
		[JsonIgnore]
		public IEnumerable<string> FieldNames
		{
			get
			{
                var repeatsFieldNames = Repeats == null
                    ? new List<string>()
                    : Repeats.Select(t => t.Name)
                        .Concat(Repeats.SelectMany(t => t.Items.SelectMany(r => r.FieldNames)));

                var tablesFieldNames = Tables == null
					? new List<string>()
					: Tables.Select(t => t.Name)
						.Concat(Tables.SelectMany(t => t.Rows.SelectMany(r => r.FieldNames)));

				var listsFieldNames = Lists == null
							? new List<string>()
							: Lists.Select(l => l.Name)
								.Concat(Lists.SelectMany(l => l.FieldNames));

				var imagesFieldNames = Images == null 
					? new List<string>() 
					: Images.Select(f => f.Name);

				var fieldNames = Fields == null ? new List<string>() : Fields.Select(f => f.Name);

				return repeatsFieldNames
                    .Concat(tablesFieldNames)
                    .Concat(listsFieldNames)
					.Concat(imagesFieldNames)
					.Concat(fieldNames);
			}
		}

		#region Fluent
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
        protected Container AddImage(ImageContent image)
        {
            if (Images == null) Images = new List<ImageContent>();

            Images.Add(image);
            return this;
        }
		#endregion

		#region IEnumerable
		public IEnumerator<IContentItem> GetEnumerator()
		{
			return All.GetEnumerator();
		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			return GetEnumerator();
		}
		#endregion

		#region Equals
		public bool Equals(Container other)
		{
			if (other == null) return false;

			return All.SequenceEqual(other);
		}

		public override int GetHashCode()
		{
			var hc = 0;
			
			hc = All.Aggregate(hc, (current, p) => current ^ p.GetHashCode());

			return hc;
		}
		#endregion
	}
}
