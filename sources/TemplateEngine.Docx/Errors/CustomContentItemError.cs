using System;

namespace TemplateEngine.Docx.Errors
{
	internal class CustomContentItemError:IError, IEquatable<CustomContentItemError>
	{
		private const string ErrorMessageTemplate =
					"{0} Content Control '{1}' {2}.";
		internal CustomContentItemError(IContentItem contentItem, string customMessage)
		{
			ContentItem = contentItem;
			_customMessage = customMessage;
		}

		private readonly string _customMessage;
		public string Message
		{
			get
			{
				return string.Format(ErrorMessageTemplate, 
					ContentItem.GetContentItemName(), 
					ContentItem.Name, _customMessage);
			}
		}

		public IContentItem ContentItem { get; private set; }
		#region Equals
		public bool Equals(CustomContentItemError other)
		{
			if (other == null) return false;

			return other.ContentItem.Equals(ContentItem) && other.Message.Equals(Message);
		}

		public bool Equals(IError other)
		{
			if (!(other is CustomContentItemError)) return false;

			return Equals((CustomContentItemError)other);
		}

		public override int GetHashCode()
		{
			var customItemHash = ContentItem.GetHashCode();

			return new { customItemHash, _customMessage }.GetHashCode();
		}
		#endregion
	}
}
