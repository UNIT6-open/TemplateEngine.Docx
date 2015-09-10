using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace TemplateEngine.Docx
{
	internal class TableProcessor
	{
		private readonly XElement _tableContentControl;
		private bool _isNeedToRemoveContentControls;

		internal TableProcessor(XElement tableContentControl)
		{
			_tableContentControl = tableContentControl;
		}

		internal TableProcessor SetRemoveContentControls(bool isNeedToRemove)
		{
			_isNeedToRemoveContentControls = isNeedToRemove;
			return this;
		}

		internal IEnumerable<string> FillTableContent(TableContent table)
		{
			var errors = new List<string>();

			// Find the content control with Table Name
			var listName = table.Name;
			
			// If there isn't a table with that name, add an error to the error string,
			// and continue with next table.
			if (_tableContentControl == null)
			{
				errors.Add(String.Format("Table Content Control '{0}' not found.",
					listName));

				return errors;
			}

			// If the table doesn't contain content controls in cells, then error and continue with next table.
			var cellContentControl = _tableContentControl
				.Descendants(W.sdt)
				.FirstOrDefault();
			if (cellContentControl == null)
			{
				errors.Add(String.Format(
					"Table Content Control '{0}' doesn't contain content controls in cells.",
					listName));
				return errors;
			}

			var fieldNames = table.FieldNames.ToList();

			var prototypeRows = GetPrototype(_tableContentControl, fieldNames);

			//Select content controls tag names
			var contentControlTagNames = prototypeRows
				.Descendants(W.sdt)
				.Select(sdt =>
					sdt.Element(W.sdtPr)
						.Element(W.tag)
						.Attribute(W.val).Value)
				.Where(fieldNames.Contains);


			//If there are not content controls with the one of specified field name we need to add the warning
			if (contentControlTagNames.Intersect(fieldNames).Count() != fieldNames.Count())
			{
				errors.Add(String.Format(
					"Table Content Control '{0}' doesn't contain rows with cell content controls {1}.",
					listName,
					string.Join(", ", fieldNames.Select(fn => string.Format("'{0}'", fn)))));
				return errors;
			}


			// Create a list of new rows to be inserted into the document.  Because this
			// is a document centric transform, this is written in a non-functional
			// style, using tree modification.
			var newRows = new List<List<XElement>>();
			foreach (var row in table.Rows)
			{
				// Clone the prototypeRows into newRowsEntry.
				var newRowsEntry = prototypeRows.Select(prototypeRow => new XElement(prototypeRow)).ToList();

				// Create new rows that will contain the data that was passed in to this
				// method in the XML tree.
				foreach (var sdt in newRowsEntry.Descendants(W.sdt).ToList())
				{
					// Get fieldName from the content control tag.
					string fieldName = sdt
						.Element(W.sdtPr)
						.Element(W.tag)
						.Attribute(W.val)
						.Value;

					// Get the new value out of contentControlValues.
					var newValueElement = row
						.Fields
						.Where(f => f.Name == fieldName)
						.FirstOrDefault();

					// Generate error message if the new value doesn't exist.
					if (newValueElement == null)
					{
						errors.Add(String.Format(
							"Table '{0}', Field '{1}' value isn't specified.",
							listName, fieldName));
						continue;
					}

					// Set content control value th the new value
					sdt.ReplaceContentControlWithNewValue(newValueElement.Value);
					if (_isNeedToRemoveContentControls)
						sdt.RemoveContentControl();
				}

				// Add the newRow to the list of rows that will be placed in the newly
				// generated table.
				newRows.Add(newRowsEntry);
			}

			prototypeRows.Last().AddAfterSelf(newRows);

			if (_isNeedToRemoveContentControls)
			{
				// Remove the content control for the table and replace it with its contents.
				var tableElement = prototypeRows.Ancestors(W.tbl).First();
				var tableClone = new XElement(tableElement);
				_tableContentControl.ReplaceWith(tableClone);
			}

			// Remove the prototype row and add all of the newly constructed rows.
			foreach (var newRow in prototypeRows)
			{
				newRow.Remove();
			}
			return errors;
		}

		// Determine the elements that contains the content controls with specified names.
		// This is the prototype for the rows that the code will generate from data.
		private List<XElement> GetPrototype(XContainer tableContentControl, IEnumerable<string> fieldNames)
		{
			var rowsWithContentControl = tableContentControl
				.Descendants(W.tr)
				.Where(tr =>
					tr.Descendants(W.sdt)
						.Any(sdt =>
							fieldNames.Contains(
								sdt.Element(W.sdtPr)
									.Element(W.tag)
									.Attribute(W.val).Value)))
				.ToList();


			return GetIntermediateAndMergedRows(rowsWithContentControl.First(), rowsWithContentControl.Last(),
				tableContentControl);
		}

		private List<XElement> GetIntermediateAndMergedRows(XElement firstRow, XElement lastRow, XContainer tableContentControl)
		{
			var resultRows = new List<XElement>();

			var mergeVector = new bool[lastRow.Descendants(W.tc).Count()];

			var firstRowReached = false;
			var lastRowReached = false;

			//find merged rows and rows between first and last rows
			foreach (var tableRow in tableContentControl.Descendants(W.tr))
			{
				if (tableRow == firstRow)
				{
					resultRows.Add(tableRow);
					firstRowReached = true;
				}
				if (!firstRowReached) continue;

				if (!lastRowReached)
				{
					if (tableRow == lastRow)
					{
						if (firstRow != lastRow)
							resultRows.Add(tableRow);

						var lastRowCells = lastRow.Descendants(W.tc).ToArray();
						for (var i = 0; i < lastRowCells.Count(); i++)
						{
							var cell = lastRowCells[i];
							var cellFormatting = cell.Element(W.tcPr);
							if (cellFormatting != null && cellFormatting.Element(W.vMerge) != null)
							{
								mergeVector[i] = true;
							}
						}
						lastRowReached = true;
						continue;
					}

					if (tableRow != firstRow)
						resultRows.Add(tableRow);
				}

				//if there are any maybe merged rows
				if (mergeVector.Any(r => r))
				{
					var rowCells = tableRow.Descendants(W.tc).ToArray();
					for (var i = 0; i < rowCells.Count(); i++)
					{
						var cell = rowCells[i];
						var cellFormatting = cell.Element(W.tcPr);
						if (cellFormatting != null && cellFormatting.Element(W.vMerge) != null &&
							(cellFormatting.Element(W.vMerge).Attribute(W.val) == null ||
							 cellFormatting.Element(W.vMerge).Attribute(W.val).Value == "continue"))
						{
							resultRows.Add(tableRow);
							mergeVector[i] = true;
						}
						else
						{
							mergeVector[i] = false;
						}
					}
				}
				else if (lastRowReached)
					break;
			}


			return resultRows;
		}
	}
}
