using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace TemplateEngine.Docx
{
	internal class ContentProcessor
	{
		private readonly XElement _content;
		private bool _isNeedToRemoveContentControls;

		internal ContentProcessor(XElement content)
		{
			_content = content;
		}

		internal ContentProcessor SetRemoveContentControls(bool isNeedToRemove)
		{
			_isNeedToRemoveContentControls = isNeedToRemove;
			return this;
		}

		internal List<string> FillContent(Content data)
		{
			var fillFieldsErrors = FillFields(data.Fields);
			var fillTablesErrors = FillTables(data.Tables);
			var fillListsErrors = FillLists(data.Lists);

			var errors = fillFieldsErrors.Concat(fillTablesErrors).Concat(fillListsErrors).ToList();


			return errors;
		}

		// Filling a tables
		private IEnumerable<string> FillTables(IEnumerable<TableContent> content)
		{
			var errors = new List<string>();

			if (content == null) return errors;


			foreach (var table in content)
			{
				var contentControls = FindContentControls(table.Name);
				foreach (var contentControl in contentControls.ToList())
				{
					errors.AddRange(new TableProcessor(contentControl)
							.SetRemoveContentControls(_isNeedToRemoveContentControls)
							.FillTableContent(table));
				}
			}

			return errors;
		}
		// Filling a lists
		private IEnumerable<string> FillLists(IEnumerable<ListContent> content)
		{
			var errors = new List<string>();

			if (content == null) return errors;

			foreach (var list in content)
			{
				var contentControls = FindContentControls(list.Name);
				foreach (var contentControl in contentControls.ToList())
				{
					errors.AddRange(new ListProcessor(contentControl)
						.SetRemoveContentControls(_isNeedToRemoveContentControls)
						.FillListContent(list));
				}
			}

			return errors;
		}

		// Filling a fields
		private IEnumerable<string> FillFields(IEnumerable<FieldContent> content)
		{
			var errors = new List<string>();
			
			if (content != null)
			{
				foreach (var field in content)
				{
					var fieldContentControls = FindContentControls(field.Name).ToList();

					// If there isn't a field with that name, add an error to the error string,
					// and continue with next field.
					if (!fieldContentControls.Any())
					{
						errors.Add(String.Format("Field Content Control '{0}' not found.",
							field.Name));
						continue;
					}

					// Set content control value to the new value
					foreach (var fieldContentControl in fieldContentControls.ToList())
					{
						errors.AddRange(new FieldsProcessor(fieldContentControl)
							.SetRemoveContentControls(_isNeedToRemoveContentControls)
							.FillFieldsContent(field));
					}
				}
			}
			return errors;
		}

		private IEnumerable<XElement> FindContentControls(string tagName)
		{
			return _content
				.Descendants(W.sdt)
				.Where(sdt => tagName == sdt
					.Element(W.sdtPr)
					.Element(W.tag)
					.Attribute(W.val)
					.Value);
		}

	}
}
