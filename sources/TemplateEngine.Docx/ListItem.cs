using System.Xml.Linq;

namespace TemplateEngine.Docx
{
	public class ListItem
	{
		public bool IsListItem { get; set; }
		public XElement Element { get; set; }
		public int? AbstractNumId { get; set; }
		public int? NumId { get; set; }
		public int? Level { get; set; }
		public ListItem(bool isListItem)
		{
			IsListItem = isListItem;
		}

		public ListItem(XElement element, int? abstractNumId, int? numId, int? level, bool isListItem)
		{
			Element = element;
			AbstractNumId = abstractNumId;
			NumId = numId;
			Level = level;
			IsListItem = isListItem;
		}
	}
}
