using System;

namespace TemplateEngine.Docx
{
	public class Content : Container, IEquatable<Content>
	{
		public Content()
		{
		}
		public Content(params IContentItem[] contentItems):base(contentItems)
		{
		}

		public bool Equals(Content other)
		{
			return base.Equals(other);
		}

		public override int GetHashCode()
		{
			return base.GetHashCode();
		}
	}
}
