using System.Collections;
using System.Collections.Generic;

namespace TemplateEngine.Docx
{
	public class Content:Container, IEnumerable<IContentItem>
	{
		public Content(params IContentItem[] contentItems):base(contentItems)
		{
		}

		public IEnumerator<IContentItem> GetEnumerator()
		{
			
			foreach (var tableContent in Tables)
			{
				yield return tableContent;
			}

			foreach (var listContent in Lists)
			{
				yield return listContent;
			}
			foreach (var fieldContent in Fields)
			{
				yield return fieldContent;
			}
		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			return GetEnumerator();
		}
	}
}
