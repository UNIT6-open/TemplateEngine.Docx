using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TemplateEngine.Docx.Errors;

namespace TemplateEngine.Docx.Processors
{
	internal class ListProcessor:IProcessor
	{
		private bool _isNeedToRemoveContentControls;
		private readonly ProcessContext _context;
	

		private class PropagationProcessResult : ProcessResult
		{
			internal IEnumerable<XElement> Result { get; set; }
		}
		/// <summary>
		/// Entire list prototype, includes all list levels.
		/// </summary>
		private class Prototype
		{
			private readonly ProcessContext _context;
			/// <summary>
			/// Creates prototype from list of prototype items.
			/// </summary>
			/// <param name="context">Process context.</param>
			/// <param name="prototypeItems">List of prototype items.</param>
			private Prototype(ProcessContext context, IEnumerable<XElement> prototypeItems)
			{
				_context = context;
				PrototypeItems = prototypeItems.ToList();
			}

			/// <summary>
			/// Creates prototype from list content control and fieldNames.
			/// </summary>
			/// <param name="context">Process context.</param>
			/// <param name="listContentControl">List content control element.</param>
			/// <param name="fieldNames">Names of fields with content.</param>
			public Prototype(ProcessContext context, XElement listContentControl, IEnumerable<string> fieldNames)
			{
				_context = context;
				if (listContentControl.Name != W.sdt)
					throw new Exception("List content control is not a content control element");

				fieldNames = fieldNames.ToList();

				// All elements inside list control content are included to the prototype.
				var listItems = listContentControl
					.Element(W.sdtContent)
					.Elements()
					.ToList();

				var tagsInPrototype = listItems.DescendantsAndSelf(W.sdt)
					.Select(sdt => sdt.SdtTagName());

				// If any field not found return empty list.
				if (fieldNames.Any(fn => !tagsInPrototype.Contains(fn)))
				{
					IsValid = false;
					return;
				}

				IsValid = true;
				PrototypeItems = listItems;
			}

			public bool IsValid { get; private set; }
			public List<XElement> PrototypeItems { get; private set; }

			public Prototype Exclude(LevelPrototype prototypeForExclude)
			{
				return new Prototype(_context, PrototypeItems
					.Where(itemPrototype => !prototypeForExclude
						.PrototypeItems
						.Contains(itemPrototype)));
			}

			/// <summary>
			/// Retrieves prototype for current content item.
			/// </summary>
			/// <param name="fieldNames">Fields names for current content item.</param>
			/// <returns>Prototype for current content item.</returns>
			public LevelPrototype CurrentLevelPrototype(IEnumerable<string> fieldNames)
			{
				return new LevelPrototype(_context, PrototypeItems, fieldNames);
			}
		}

		/// <summary>
		/// Prototype of the concrete level of list.
		/// </summary>
		private class LevelPrototype
		{
			private readonly ProcessContext _context;
			public bool IsValid { get; private set; }
			public LevelPrototype(ProcessContext context, IEnumerable<XElement> prototypeItems, IEnumerable<string> fieldNames)
			{
				_context = context;
				var currentLevelPrototype = new List<XElement>();

				// Items for which no content control.
				// Add this items to the prototype if there are items after them.
				var maybeNeedToAdd = new List<XElement>();
				var numberingElementReached = false;

				foreach (var prototypeItem in prototypeItems)
				{
					//search for first item with numbering
					if (!numberingElementReached)
					{
						var paragraph = prototypeItem.DescendantsAndSelf(W.p).FirstOrDefault();
						if (paragraph != null &&
							ListItemRetriever.RetrieveListItem(
							context.Document.NumberingPart, context.Document.StylesPart, paragraph)
							.IsListItem)
							numberingElementReached = true;
						else
							continue;
					}
					if ((!prototypeItem.FirstLevelDescendantsAndSelf(W.sdt).Any() && prototypeItem.Value != "") ||
						(prototypeItem
						.FirstLevelDescendantsAndSelf(W.sdt)
						.Any(sdt => fieldNames.Contains(sdt.SdtTagName()))))
					{
						currentLevelPrototype.AddRange(maybeNeedToAdd);
						currentLevelPrototype.Add(prototypeItem);
					}

					else
					{
						maybeNeedToAdd.Add(prototypeItem);
					}
				}
				if (!currentLevelPrototype.Any()) return;

				PrototypeItems = currentLevelPrototype;

				if (fieldNames.Any(fn => !SdtTags.Contains(fn)))
				{
					IsValid = false;
					return;
				}

				IsValid = true;
				PrototypeItems = currentLevelPrototype;

			}

			private LevelPrototype(ProcessContext context, IEnumerable<XElement> prototypeItems)
			{
				_context = context;
				PrototypeItems = prototypeItems.ToList();
			}
			public List<XElement> PrototypeItems { get; private set; }

			private List<string> SdtTags
			{
				get
				{
					return PrototypeItems == null
						? new List<string>() 
						: PrototypeItems.DescendantsAndSelf(W.sdt)
						.Select(sdt => sdt.SdtTagName())
						.ToList();
				}
			}

			public LevelPrototype Clone()
			{
				return new LevelPrototype(_context, PrototypeItems.ToList());
			}
		}

		public ListProcessor(ProcessContext context)
		{
			_context = context;
		}
		public IProcessor SetRemoveContentControls(bool isNeedToRemove)
		{
			_isNeedToRemoveContentControls = isNeedToRemove;
			return this;
		}

		public ProcessResult FillContent(XElement contentControl, IEnumerable<IContentItem> items)
		{
			var processResult = ProcessResult.NotHandledResult; 
			var handled = false;

			foreach (var contentItem in items)
			{
				var itemProcessResult = FillContent(contentControl, contentItem);
				processResult.Merge(itemProcessResult);

				if (!itemProcessResult.Handled) continue;

				handled = true;
			}

			if (!handled) return processResult;

		    if (!processResult.Success || !_isNeedToRemoveContentControls) return processResult;

		    foreach (var sdt in contentControl.Descendants(W.sdt).ToList())
		    {
		        // Remove the content control, and replace it with its contents.
		        sdt.RemoveContentControl();
		    }
		    contentControl.RemoveContentControl();
		    return processResult;
		}

		private ProcessResult FillContent(XElement contentControl, IContentItem item)
		{
			var processResult = ProcessResult.NotHandledResult; 
			if (!(item is ListContent))
			{
				return ProcessResult.NotHandledResult;
			}

			var list = item as ListContent;

			// If there isn't a list with that name, add an error to the error string.
			if (contentControl == null)
			{
				processResult.AddError(new ContentControlNotFoundError(list));

				return processResult;
			}

			// If the list doesn't contain content controls in items, then error.
			var itemsContentControl = contentControl
				.Descendants(W.sdt)
				.FirstOrDefault();

			if (itemsContentControl == null)
			{
				processResult.AddError(
					new CustomContentItemError(list, "doesn't contain content controls in items"));
			
				return processResult;
			}			

		    if (list.IsHidden || list.FieldNames == null)
		    {
		        contentControl.Descendants(W.tr).Remove();
		    }
		    else
		    {
		        var fieldNames = list.FieldNames.ToList();

		        // Create a prototype of new items to be inserted into the document.
                var prototype = new Prototype(_context, contentControl, fieldNames);

		        if (!prototype.IsValid)
		        {
		            processResult.AddError(
		                new CustomContentItemError(list,
		                    string.Format("doesn't contain items with content controls {0}",
		                        string.Join(", ", fieldNames))));

		            return processResult;
		        }

		        new NumberingAccessor(_context.Document.NumberingPart, _context.LastNumIds)
		            .ResetNumbering(prototype.PrototypeItems);

		        // Propagates a prototype.
		        if (list.Items != null)

		        {
		            var propagationResult = PropagatePrototype(prototype, list.Items);
		            processResult.Merge(propagationResult);
		            // add all of the newly constructed rows.
		            if (!item.IsHidden) prototype.PrototypeItems.Last().AddAfterSelf(propagationResult.Result);
		        }

		        prototype.PrototypeItems.Remove();
            }                        
				
			processResult.AddItemToHandled(list);
			
			return processResult;
		}

		// Fills prototype with values recursive.
		private PropagationProcessResult PropagatePrototype(Prototype prototype, 
			IEnumerable<ListItemContent> content)
		{
			var processResult = new PropagationProcessResult();
			var newRows = new List<XElement>();

			foreach (var contentItem in content)
			{
				var currentLevelPrototype = prototype.CurrentLevelPrototype(contentItem.FieldNames);
				
				if (currentLevelPrototype == null || !currentLevelPrototype.IsValid)
				{
					processResult.AddError(new CustomError(
						string.Format("Prototype for list item '{0}' not found", 
							string.Join(", ", contentItem.FieldNames))));

					continue;
				}
				
				// Create new item from the prototype.
				var newItemEntry = currentLevelPrototype.Clone();

				foreach (var xElement in newItemEntry.PrototypeItems)
				{
					var newElement = new XElement(xElement);
					if (!newElement.DescendantsAndSelf(W.sdt).Any())
					{
						newRows.Add(newElement);
						continue;
					}

					foreach (var sdt in newElement.FirstLevelDescendantsAndSelf(W.sdt).ToList())
					{
						var fieldContent = contentItem.GetContentItem(sdt.SdtTagName());
						if (fieldContent == null)
						{
							processResult.AddError(new CustomError(
								string.Format("Field content for field '{0}' not found", 
								sdt.SdtTagName())));

							continue;
						}
						
						var contentProcessResult = new ContentProcessor(_context)
							.SetRemoveContentControls(_isNeedToRemoveContentControls)
							.FillContent(sdt, fieldContent);

						processResult.Merge(contentProcessResult);
						
						
					}
					newRows.Add(newElement);
					
				}
				
				// If there are nested items fill prototype for them.
				if (contentItem.NestedFields != null)
				{
					var filledNestedFields = PropagatePrototype(
						prototype.Exclude(currentLevelPrototype), 
						contentItem.NestedFields);

					newRows.AddRange(filledNestedFields.Result);					
				}		
			}
			processResult.Result = newRows;
			return processResult;
		}
	}
}
