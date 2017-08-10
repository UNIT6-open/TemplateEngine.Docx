using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TemplateEngine.Docx.Errors;

namespace TemplateEngine.Docx.Processors
{
	internal class RepeatProcessor : IProcessor
	{
        #region fields

        private bool _isNeedToRemoveContentControls;
		private readonly ProcessContext _context;

        #endregion fields

        #region constructors

        public RepeatProcessor(ProcessContext context)
        {
            _context = context;
        }

        #endregion constructors

        #region methods

        private PropagationProcessResult PropagatePrototype(Prototype prototype, IEnumerable<Content> content)
        {
            var processResult = new PropagationProcessResult();
            var newRows = new List<XElement>();

            foreach (var contentItem in content)
            {
                // Create new item from the prototype.
                var newItemEntry = prototype.Clone();

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
            }
            processResult.Result = newRows;
            return processResult;
        }
        private ProcessResult FillContent(XElement contentControl, IContentItem item)
        {
            var processResult = ProcessResult.NotHandledResult;
            if (!(item is RepeatContent))
            {
                return ProcessResult.NotHandledResult;
            }

            var repeat = item as RepeatContent;

            // If there isn't a list with that name, add an error to the error string.
            if (contentControl == null)
            {
                processResult.AddError(new ContentControlNotFoundError(repeat));

                return processResult;
            }

            // If the list doesn't contain content controls in items, then error.
            var itemsContentControl = contentControl
                .Descendants(W.sdt)
                .FirstOrDefault();

            if (itemsContentControl == null)
            {
                processResult.AddError(
                    new CustomContentItemError(repeat, "doesn't contain content controls in items"));

                return processResult;
            }

            if (repeat.IsHidden || repeat.Items == null)
            {
                contentControl.DescendantsAndSelf(W.sdt).Remove();
            }
            else
            {
                var fieldNames = repeat.FieldNames.ToList();

                // Create a prototype of new items to be inserted into the document.
                var prototype = new Prototype(_context, contentControl, fieldNames);

                if (!prototype.IsValid)
                {
                    processResult.AddError(
                        new CustomContentItemError(repeat,
                            String.Format("doesn't contain items with content controls {0}",
                                string.Join(", ", fieldNames))));

                    return processResult;
                }

                // Propagates a prototype.
                var propagationResult = PropagatePrototype(prototype, repeat.Items);

                processResult.Merge(propagationResult);

                // Remove the prototype row and add all of the newly constructed rows.
                prototype.PrototypeItems.Last().AddAfterSelf(propagationResult.Result);
                prototype.PrototypeItems.Remove();
            }           

            processResult.AddItemToHandled(repeat);

            return processResult;
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

            if (processResult.Success && _isNeedToRemoveContentControls)
            {
                foreach (var sdt in contentControl.Descendants(W.sdt).ToList())
                {
                    // Remove the content control, and replace it with its contents.
                    sdt.RemoveContentControl();
                }
                contentControl.RemoveContentControl();
            }
            return processResult;
        }

        #endregion methods

        #region classes

        private class PropagationProcessResult : ProcessResult
		{
			internal IEnumerable<XElement> Result { get; set; }
		}
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

            public Prototype Clone()
            {
                return new Prototype(_context, PrototypeItems.ToList());
            }
		}

        #endregion classes
	}
}
