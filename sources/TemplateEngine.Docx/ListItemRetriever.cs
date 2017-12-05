using System.Linq;
using System.Xml.Linq;

namespace TemplateEngine.Docx
{
	public class ListItemRetriever
	{
		private class ListItemInfo
		{
			public bool IsListItem { get; }
			public XElement Lvl { get; set; }
			public int? Start { get; set; }
			public int? AbstractNumId { get; set; }
			public ListItemInfo(bool isListItem)
			{
				IsListItem = isListItem;
			}
		}

		private static ListItemInfo GetListItemInfoByNumIdAndIlvl(XDocument numbering,
			XDocument styles, int numId, int ilvl)
		{
			if (numId == 0)
				return new ListItemInfo(false);
			var listItemInfo = new ListItemInfo(true);
			var num = numbering.Root
				.Elements(W.num).FirstOrDefault(e => (int)e.Attribute(W.numId) == numId);
			if (num == null)
				return new ListItemInfo(false);

			listItemInfo.AbstractNumId = (int?)num.Elements(W.abstractNumId)
				.Attributes(W.val).FirstOrDefault();
			var lvlOverride = num
				.Elements(W.lvlOverride).FirstOrDefault(e => (int)e.Attribute(W.ilvl) == ilvl);
			// If there is a w:lvlOverride element, and if the w:lvlOverride contains a
			// w:lvl element, then return it.  Otherwise, go look in the abstract numbering
			// definition.
			if (lvlOverride != null)
			{
				// Get the startOverride, if there is one.
				listItemInfo.Start = (int?)num.Elements(W.lvlOverride)
					.Where(o => (int)o.Attribute(W.ilvl) == ilvl).Elements(W.startOverride)
					.Attributes(W.val).FirstOrDefault();
				listItemInfo.Lvl = lvlOverride.Element(W.lvl);
				if (listItemInfo.Lvl != null)
				{
					if (listItemInfo.Start == null)
						listItemInfo.Start = (int?)listItemInfo.Lvl.Elements(W.start)
							.Attributes(W.val).FirstOrDefault();
					return listItemInfo;
				}
			}
			var a = listItemInfo.AbstractNumId;
			var abstractNum = numbering.Root
				.Elements(W.abstractNum).FirstOrDefault(e => (int)e.Attribute(W.abstractNumId) == a);
			var numStyleLink = (string)abstractNum.Elements(W.numStyleLink)
				.Attributes(W.val).FirstOrDefault();
			if (numStyleLink != null)
			{
				var style = styles.Root
					.Elements(W.style)
					.FirstOrDefault(e => (string)e.Attribute(W.styleId) == numStyleLink);
				var numPr = style.Elements(W.pPr).Elements(W.numPr).FirstOrDefault();
				var lNumId = (int)numPr.Elements(W.numId).Attributes(W.val)
					.FirstOrDefault();
				return GetListItemInfoByNumIdAndIlvl(numbering, styles, lNumId, ilvl);
			}
			for (int l = ilvl; l >= 0; --l)
			{
				listItemInfo.Lvl = abstractNum
					.Elements(W.lvl).FirstOrDefault(e => (int)e.Attribute(W.ilvl) == l);
				if (listItemInfo.Lvl == null)
					continue;
				if (listItemInfo.Start == null)
					listItemInfo.Start = (int?)listItemInfo.Lvl.Elements(W.start)
						.Attributes(W.val).FirstOrDefault();
				return listItemInfo;
			}
			return new ListItemInfo(false);
		}

		private static ListItemInfo GetListItemInfoByNumIdAndStyleId(XDocument numbering,
			XDocument styles, int numId, string paragraphStyle)
		{
			// If you have to find the w:lvl by style id, then we can't find it in the
			// w:lvlOverride, as that requires that you have determined the level already.
			var listItemInfo = new ListItemInfo(true);
			var num = numbering.Root
				.Elements(W.num).FirstOrDefault(e => (int)e.Attribute(W.numId) == numId);

			listItemInfo.AbstractNumId = (int)num.Elements(W.abstractNumId)
				.Attributes(W.val).FirstOrDefault();
			var a = listItemInfo.AbstractNumId;
			var abstractNum = numbering.Root
				.Elements(W.abstractNum).FirstOrDefault(e => (int)e.Attribute(W.abstractNumId) == a);
			var numStyleLink = (string)abstractNum.Element(W.numStyleLink)
				.Attributes(W.val).FirstOrDefault();
			if (numStyleLink != null)
			{
				var style = styles.Root
					.Elements(W.style)
					.FirstOrDefault(e => (string)e.Attribute(W.styleId) == numStyleLink);
				var numPr = style.Elements(W.pPr).Elements(W.numPr).FirstOrDefault();
				var lNumId = (int)numPr.Elements(W.numId).Attributes(W.val).FirstOrDefault();
				return GetListItemInfoByNumIdAndStyleId(numbering, styles, lNumId,
					paragraphStyle);
			}
			listItemInfo.Lvl = abstractNum.Elements(W.lvl).FirstOrDefault(e => (string)e.Element(W.pStyle) == paragraphStyle);
			listItemInfo.Start = (int?)listItemInfo.Lvl.Elements(W.start).Attributes(W.val)
				.FirstOrDefault();
			return listItemInfo;
		}

		public static ListItem RetrieveListItem(XDocument numbering, XDocument styles,
			XElement paragraph)
		{
			// The following is an optimization - only determine ListItemInfo once for a
			// paragraph.
			var listItem = paragraph.Annotation<ListItem>();
			if (listItem != null)
				return listItem;

			var paragraphNumberingProperties = paragraph.Elements(W.pPr)
				.Elements(W.numPr).FirstOrDefault();

			var paragraphStyle = (string)paragraph.Elements(W.pPr).Elements(W.pStyle)
				.Attributes(W.val).FirstOrDefault();

			ListItemInfo listItemInfo;
			if (paragraphNumberingProperties != null &&
				paragraphNumberingProperties.Element(W.numId) != null)
			{
				// Paragraph numbering properties must contain a numId.
				var numId = (int)paragraphNumberingProperties.Elements(W.numId)
					.Attributes(W.val).FirstOrDefault();

				var ilvl = (int?)paragraphNumberingProperties.Elements(W.ilvl)
					.Attributes(W.val).FirstOrDefault();

				if (ilvl != null)
				{
					listItemInfo = GetListItemInfoByNumIdAndIlvl(numbering, styles, numId,
						(int)ilvl);
					paragraph.AddAnnotation(listItemInfo);
					return new ListItem(paragraph, listItemInfo.AbstractNumId, numId, ilvl, listItemInfo.IsListItem);
				}
				if (paragraphStyle != null)
				{
					listItemInfo = GetListItemInfoByNumIdAndStyleId(numbering, styles,
						numId, paragraphStyle);
					paragraph.AddAnnotation(listItemInfo);
					return new ListItem(paragraph, listItemInfo.AbstractNumId, numId, ilvl, listItemInfo.IsListItem);
				}
				listItemInfo = new ListItemInfo(false);
				paragraph.AddAnnotation(listItemInfo);
				return new ListItem(paragraph, listItemInfo.AbstractNumId, numId, ilvl, listItemInfo.IsListItem);
			}
			if (paragraphStyle != null)
			{
				var style = styles.Root.Elements(W.style).FirstOrDefault(s => (string)s.Attribute(W.type) == "paragraph" &&
					(string)s.Attribute(W.styleId) == paragraphStyle);

				var styleNumberingProperties = style?.Elements(W.pPr)
					.Elements(W.numPr).FirstOrDefault();
				if (styleNumberingProperties?.Element(W.numId) != null)
				{
					var numId = (int)styleNumberingProperties.Elements(W.numId)
						.Attributes(W.val).FirstOrDefault();

					var ilvl = (int?)styleNumberingProperties.Elements(W.ilvl)
						.Attributes(W.val).FirstOrDefault();

					if (ilvl == null)
						ilvl = 0;

					listItemInfo = GetListItemInfoByNumIdAndIlvl(numbering, styles,
						numId, (int)ilvl);
					paragraph.AddAnnotation(listItemInfo);
					return new ListItem(paragraph, listItemInfo.AbstractNumId, numId, ilvl, listItemInfo.IsListItem);
				}
			}
			listItemInfo = new ListItemInfo(false);
			paragraph.AddAnnotation(listItemInfo);
			return new ListItem(paragraph, listItemInfo.AbstractNumId, null, null, listItemInfo.IsListItem);
		}
	}
}
