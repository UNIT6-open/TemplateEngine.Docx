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

			// If there isn't a list with that name, add an error to the error string
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

			var fieldName = list.Items.First().Name;

			var prototypeItem = GetPrototype(_listContentControl, fieldName);

			//Select content controls tag names
			var contentControlTagName = prototypeItem
				.Element(W.sdtPr)
				.Element(W.tag)
				.Attribute(W.val).Value;


			//If there are not content controls with the one of specified field name we need to add the warning
			if (contentControlTagName != fieldName)
			{
				errors.Add(String.Format(
					"Table Content Control '{0}' doesn't contain items with content controls {1}.",
					listName,
					string.Join(", ", fieldName)));
				return errors;
			}


			// Create a list of new rows to be inserted into the document.  Because this
			// is a document centric transform, this is written in a non-functional
			// style, using tree modification.
			var newRows = new List<XElement>();
			
			foreach (var row in list.Items)
			{
				// Clone the prototypeRows into newRowsEntry.
				var newItemEntry = new XElement(prototypeItem);


				// Get fieldName from the content control tag.
				string tagName = newItemEntry
					.Element(W.sdtPr)
					.Element(W.tag)
					.Attribute(W.val)
					.Value;

				// Get the new value out of contentControlValues.
				var newValueElement = row;


				// Generate error message if the new value doesn't exist.
				if (newValueElement == null || newValueElement.Name != tagName)
				{
					errors.Add(String.Format(
						"List '{0}', Field '{1}' value isn't specified.",
						listName, fieldName));
					continue;
				}

				// Set content control value the new value
				newItemEntry.ReplaceContentControlWithNewValue(newValueElement.Value);
				
				newRows.Add(newItemEntry);
				
				
			}

			prototypeItem.AddAfterSelf(newRows);

			// Remove the prototype row and add all of the newly constructed rows.
			prototypeItem.Remove();

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

		private XElement GetPrototype(XElement listContentControl, string fieldName)
		{
			var itemWithContentControl = listContentControl
				.Descendants(W.sdt)
				.FirstOrDefault(sdt =>
					fieldName == sdt.Element(W.sdtPr)
						.Element(W.tag)
						.Attribute(W.val).Value);


			return itemWithContentControl;
		}
	}
}
