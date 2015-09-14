using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace TemplateEngine.Docx
{
	static class XElementExtensions
	{
		// Set content control value th the new value
		public static void ReplaceContentControlWithNewValue(this XElement sdt, string newValue)
		{

			var sdtContentElement = sdt.Element(W.sdtContent);

			if (sdtContentElement != null)
			{
				var firstContentElementWithText =
					sdtContentElement
						.Descendants().FirstOrDefault(d => d.Descendants(W.t).Any());

				if (firstContentElementWithText != null)
				{
					var firstTextElement = firstContentElementWithText
						.Descendants(W.t)
						.First();

					firstTextElement.Value = newValue;

					//remove all text elements with its ancestors from the first contentElement
					var firstElementAncestors = firstTextElement.Ancestors();
					foreach (
						var descendantsWithText in firstContentElementWithText.Descendants().Where(d => d.Descendants(W.t).Any()).ToList()
						)
					{
						descendantsWithText.AncestorsAndSelf().Where(a => !firstElementAncestors.Contains(a)).Remove();
					}

					var contentReplacementElement = new XElement(firstContentElementWithText);

					sdtContentElement.Descendants().Where(d => d.Descendants(W.t).Any() && d != firstContentElementWithText).Remove();

					firstContentElementWithText.AddAfterSelf(contentReplacementElement);
					firstContentElementWithText.Remove();
				}
				else
				{
					if (sdtContentElement.Elements(W.p).Any())
					{
						sdtContentElement.Element(W.p).Add(new XElement(W.r, new XElement(W.t, newValue)));
					}
					else
					{
						sdtContentElement.Add(new XElement(W.p), new XElement(W.r, new XElement(W.t, newValue)));
					}
				}
			}
			else
			{
				sdt.Add(new XElement(W.sdtContent, new XElement(W.p), new XElement(W.r, new XElement(W.t, newValue))));
			}
		}

		public static IEnumerable<XElement> RemoveContentControl(this XElement sdt)
		{

			var sdtContentElement = sdt.Element(W.sdtContent);
			if (sdtContentElement == null)
			{
				sdt.Remove();
				return null;
			}

			// Remove the content control, and replace it with its contents.
			var replacementElements = new List<XElement>(sdt.Element(W.sdtContent).Elements());

			sdt.ReplaceWith(replacementElements);
			return replacementElements;
		}
	}
}
