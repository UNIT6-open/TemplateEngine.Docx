using System.Linq;
using System.Xml.Linq;

namespace TemplateEngine.Docx
{
	static class XElementExtensions
	{
		// Set content control value th the new value
		public static void ReplaceContentControlWithNewValue(this XElement sdt, string newValue, bool isNeedToRemoveContentControls)
		{

			var sdtContentElement = sdt.Element(W.sdtContent);

			if (sdtContentElement != null)
			{
				var firstContentElementWithText =
					sdtContentElement
					.Descendants().FirstOrDefault(d=>d.Descendants(W.t).Any());

				if (firstContentElementWithText != null)
				{
					var firstTextElement = firstContentElementWithText
						.Descendants(W.t)
						.First();
					
					firstTextElement.Value = newValue;

					//remove all text elements with its ancestors from the first contentElement
					var firstElementAncestors = firstTextElement.Ancestors();
					foreach (var descendantsWithText in firstContentElementWithText.Descendants().Where(d => d.Descendants(W.t).Any()).ToList())
					{
						descendantsWithText.AncestorsAndSelf().Where(a => !firstElementAncestors.Contains(a)).Remove();
					}

					var contentReplacementElement = new XElement(firstContentElementWithText);
					
					sdtContentElement.Descendants().Remove();
					sdtContentElement.Add(contentReplacementElement);
					
					if (isNeedToRemoveContentControls)
					{
						// Remove the content control, and replace it with its contents.
						var replacementElement =
							new XElement(sdtContentElement.Elements().First());
						sdt.ReplaceWith(replacementElement);
					}
				}
			}
		}
	}
}
