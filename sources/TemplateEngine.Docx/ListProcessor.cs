using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace TemplateEngine.Docx
{
	internal class ListProcessor
	{
		private readonly XElement _listContentControl;
		private bool _isNeedToRemoveContentControls;

		internal ListProcessor(XElement tableContentControl)
		{
			_listContentControl = tableContentControl;
		}

		internal ListProcessor SetRemoveContentControls(bool isNeedToRemove)
		{
			_isNeedToRemoveContentControls = isNeedToRemove;
			return this;
		}

		internal IEnumerable<string> FillListContent(ListContent list)
		{
			var errors = new List<string>();
			var listName = list.Name;

			// If there isn't a list with that name, add an error to the error string.
			if (_listContentControl == null)
			{
				errors.Add(String.Format("List Content Control '{0}' not found.",
					listName));

				return errors;
			}

			// If the list doesn't contain content controls in items, then error.
			var itemsContentControl = _listContentControl
				.Descendants(W.sdt)
				.FirstOrDefault();
			if (itemsContentControl == null)
			{
				errors.Add(String.Format(
					"List Content Control '{0}' doesn't contain content controls in items.",
					listName));
				return errors;
			}

			var fieldNames = list.FieldNames.ToList();

			var prototypeItems = GetPrototype(_listContentControl, fieldNames).ToList();
			
			if (!prototypeItems.Any())
			{
				errors.Add(String.Format(
					"List Content Control '{0}' doesn't contain items with content controls {1}.",
					listName,
					string.Join(", ", fieldNames)));
				return errors;
			}

			// Create a list of new items to be inserted into the document.  Because this
			// is a document centric transform, this is written in a non-functional
			// style, using tree modification.
			IEnumerable<XElement> newItems;
			try
			{
				newItems = FillPrototype(prototypeItems, list.Items);
			}
			catch (Exception e)
			{
				errors.Add(e.Message);
				return errors;
			}
			prototypeItems.Last().AddAfterSelf(newItems);

			// Remove the prototype row and add all of the newly constructed rows.
			prototypeItems.Remove();

			if (_isNeedToRemoveContentControls)
			{
				foreach (var item in _listContentControl.Descendants(W.sdt).ToList())
				{
					// Remove the content control, and replace it with its contents.
					item.RemoveContentControl();										
				}
				_listContentControl.RemoveContentControl();
			}
			return errors;
		}

		private IEnumerable<XElement> GetPrototype(XContainer listContentControl, IEnumerable<string> fieldNames)
		{
			var itemsWithContentControl = listContentControl
				.Descendants(W.sdt)
				.Where(sdt =>
					fieldNames.Contains(
						sdt.Element(W.sdtPr)
						.Element(W.tag)
						.Attribute(W.val).Value))
				.ToList();


			var tagsInPrototype = itemsWithContentControl.Select(sdt => sdt.Element(W.sdtPr).Element(W.tag).Attribute(W.val).Value);

			// If any field not found return empty list.
			if (tagsInPrototype.Intersect(fieldNames).Count() != fieldNames.Count())
				return new List<XElement>();

			return GetIntermediateItems(itemsWithContentControl.First(), itemsWithContentControl.Last(), _listContentControl);
		}

		// Returns items that there are between first and last content control items.
		private IEnumerable<XElement> GetIntermediateItems(XElement firstItem, XElement lastItem, XContainer listContentControl)
		{
			var resultItems = new List<XElement>();

			var firstItemReached = false;
			
			//find items between first and last item.
			foreach (var listItem in listContentControl.Element(W.sdtContent).Elements())
			{
				if (listItem == firstItem)
				{
					resultItems.Add(listItem);
					firstItemReached = true;
				}
				else if (firstItemReached)
				{
					resultItems.Add(listItem);
					if (listItem == lastItem)
						break;					
				}				
			}

			return resultItems;
		}

		// Fills prototype with values recursive.
		private IEnumerable<XElement> FillPrototype(IEnumerable<XElement> prototypeItems, IEnumerable<FieldContent> content)
		{

			var newRows = new List<XElement>();
			var prototypeItemsCopy = prototypeItems.ToList();

			// If there are items without content control, add them to output.
			foreach (var xElement in prototypeItemsCopy)
			{
				if (xElement.DescendantsAndSelf(W.sdt).Any())
					break;
				
				var newItemEntry = new XElement(xElement);
				newRows.Add(newItemEntry);
				prototypeItemsCopy.Remove(xElement);

			}
			
			foreach (var contentItem in content)
			{
				// Prototype for current content item.
				var currentLevelPrototype = prototypeItemsCopy
					.FirstOrDefault(p => p
						.DescendantsAndSelf(W.sdtPr)
						.Any(d => d.Element(W.tag)
					.Attribute(W.val)
					.Value == contentItem.Name));

				if (currentLevelPrototype == null)
				{
					throw new Exception(string.Format("Prototype for list item '{0}' not found", contentItem.Name));
				}

				// Create new item from the prototype.
				var newItemEntry = new XElement(currentLevelPrototype);

				// Replace item's content control value with the new value.
				var std = newItemEntry.DescendantsAndSelf(W.sdt).First();
				std.ReplaceContentControlWithNewValue(contentItem.Value);
				newRows.Add(newItemEntry);
				
				// If there are nested items fill prototype for them.
				if (contentItem is ListItemContent)
				{
					var listItem = contentItem as ListItemContent;
					if (listItem.NestedFields == null) continue;

					
					var filledNestedFields = FillPrototype(prototypeItemsCopy.Where(itemPrototype => 
						itemPrototype != currentLevelPrototype), listItem.NestedFields);
		
					newRows.AddRange(filledNestedFields);
				}

			}
			return newRows;
		}
	}
}
