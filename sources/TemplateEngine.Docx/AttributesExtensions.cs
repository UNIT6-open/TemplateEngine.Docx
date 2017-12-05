using System.Linq;

namespace TemplateEngine.Docx
{
	internal static class AttributesExtensions
	{
		public static string GetContentItemName(this IContentItem value)
		{
			var contentItemNameAttribute = value.GetType()
				.GetCustomAttributes(typeof(ContentItemNameAttribute), true)
				   .FirstOrDefault() as ContentItemNameAttribute;

			return contentItemNameAttribute?.Name;
		}
	}
}
