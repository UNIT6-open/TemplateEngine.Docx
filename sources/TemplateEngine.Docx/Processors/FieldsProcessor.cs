using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace TemplateEngine.Docx.Processors
{
	internal class FieldsProcessor:IProcessor
	{
		private readonly ProcessContext _context;
		private ProcessResult _processResult;
		public FieldsProcessor(ProcessContext context)
		{
			_context = context;
		}

		private bool _isNeedToRemoveContentControls;

		public IProcessor SetRemoveContentControls(bool isNeedToRemove)
		{
			_isNeedToRemoveContentControls = isNeedToRemove;
			return this;
		}

		public ProcessResult FillContent(XElement contentControl, IEnumerable<IContentItem> items)
		{
			_processResult = new ProcessResult();

			foreach (var contentItem in items)
			{
				FillContent(contentControl, contentItem);
			}


			if (_processResult.Success && _isNeedToRemoveContentControls)
				contentControl.RemoveContentControl();

			return _processResult;
		}

		public void FillContent(XElement contentControl, IContentItem item)
		{
			if (!(item is FieldContent))
			{
				_processResult = ProcessResult.NotHandledResult;
				return;
			}

			var field = item as FieldContent;

			// If there isn't a field with that name, add an error to the error string,
			// and continue with next field.
			if (contentControl == null)
			{
				_processResult.Errors.Add(String.Format("Field Content Control '{0}' not found.",
					field.Name));
				return;
			}
			contentControl.ReplaceContentControlWithNewValue(field.Value);

		}
	}
}
