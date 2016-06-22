using System;

namespace TemplateEngine.Docx.Errors
{
	internal class ContentControlNotFoundError : IError, IEquatable<ContentControlNotFoundError>, IEquatable<IError>
	{
		private const string ErrorMessageTemplate =
					"{0} Content Control '{1}' not found.";
		internal ContentControlNotFoundError(IContentItem contentItem)
		{
			ContentItem = contentItem;
		}

		public string Message
		{
			get
			{
				return string.Format(ErrorMessageTemplate, ContentItem.GetContentItemName(), ContentItem.Name);
			}
		}

		public IContentItem ContentItem { get; private set; }

		#region Equals
		public bool Equals(ContentControlNotFoundError other)
		{
			if (other == null) return false;

			return other.ContentItem.Equals(ContentItem);
		}

		public bool Equals(IError other)
		{
			if (!(other is ContentControlNotFoundError)) return false;

			return Equals((ContentControlNotFoundError) other);
		}

		public override int GetHashCode()
		{
			return ContentItem.GetHashCode();
		}
		#endregion
	}
}
