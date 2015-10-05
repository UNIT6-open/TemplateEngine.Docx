using System;
using System.Xml.Linq;

namespace TemplateEngine.Docx.Processors
{
	internal class FieldsProcessor:IProcessor
	{
		private readonly ProcessContext _context;
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

		public ProcessResult FillContent(XElement contentControl, IContentItem item)
		{
			if (!(item is FieldContent)) return ProcessResult.NotHandledResult;

			var processResult = new ProcessResult();
			var field = item as FieldContent;

			// If there isn't a field with that name, add an error to the error string,
			// and continue with next field.
			if (contentControl == null)
			{
				processResult.Errors.Add(String.Format("Field Content Control '{0}' not found.",
					field.Name));
				return processResult;
			}
			contentControl.ReplaceContentControlWithNewValue(field.Value);


			if (_isNeedToRemoveContentControls)
				contentControl.RemoveContentControl();


			return processResult;
		}
	}
}
