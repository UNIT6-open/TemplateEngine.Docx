using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace TemplateEngine.Docx
{
	[ContentItemName("Repeat")]
	public class RepeatContent : IContentItem, IEquatable<RepeatContent>
	{
        #region properties

        public string Name { get; set; }

		public ICollection<RepeatItemContent> Items { get; set; }

        public IEnumerable<string> FieldNames
        {
            get
            {
                return Items == null ? new List<string>() : Items.SelectMany(r => r.FieldNames).Distinct().ToList();
            }
        }

        #endregion properties

        #region ctors
        public RepeatContent()
        {
            
        }

        public RepeatContent(string name)
        {
            Name = name;
        }

		public RepeatContent(string name, IEnumerable<RepeatItemContent> items)
            : this(name)
        {
            Items = items.ToList();
        }

		public RepeatContent(string name, params RepeatItemContent[] items)
            : this(name)
        {
			Items = items.ToList();
        }

		#endregion

		#region Fluent

		public static RepeatContent Create(string name, params RepeatItemContent[] items)
        {
			return new RepeatContent(name, items);
        }
		public static RepeatContent Create(string name, IEnumerable<RepeatItemContent> items)
        {
			return new RepeatContent(name, items);
        }

		public RepeatContent AddItem(RepeatItemContent item)
		{
			if (Items == null) Items = new Collection<RepeatItemContent>();
			Items.Add(item);
			return this;
		}
		public RepeatContent AddItem(params IContentItem[] contentItems)
		{
			if (Items == null) Items = new Collection<RepeatItemContent>();
			Items.Add(new RepeatItemContent(contentItems));
			return this;
		}

		#endregion

		#region Equals
		public bool Equals(RepeatContent other)
		{
			if (other == null) return false;
			return Name.Equals(other.Name) &&
			       Items.SequenceEqual(other.Items);
		}

		public bool Equals(IContentItem other)
		{
			if (!(other is RepeatContent)) return false;

			return Equals((RepeatContent)other);
		}


		public override int GetHashCode()
		{
			var hc = 0;
			if (Items != null)
				hc = Items.Aggregate(hc, (current, p) => current ^ p.GetHashCode());
			
			return new { Name, hc }.GetHashCode();
		}

		#endregion
	}
}
