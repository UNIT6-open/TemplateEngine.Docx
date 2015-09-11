using System;
using System.Collections.Generic;
using System.Xml.Linq;
namespace TemplateEngine.Docx
{
	internal class FieldsProcessor
	{
		private readonly XElement _fieldContentControl;
		private bool _isNeedToRemoveContentControls;

		internal FieldsProcessor(XElement tableContentControl)
		{
			_fieldContentControl = tableContentControl;
		}

		internal FieldsProcessor SetRemoveContentControls(bool isNeedToRemove)
		{
			_isNeedToRemoveContentControls = isNeedToRemove;
			return this;
		}

		internal IEnumerable<string> FillFieldsContent(FieldContent field)
		{
			var errors = new List<string>();
			
			// If there isn't a field with that name, add an error to the error string,
			// and continue with next field.
			if (_fieldContentControl == null)
			{
				errors.Add(String.Format("Field Content Control '{0}' not found.",
					field.Name));

			}

			_fieldContentControl.ReplaceContentControlWithNewValue(field.Value);
			if (_isNeedToRemoveContentControls)
				_fieldContentControl.RemoveContentControl();

			return errors;

		}
	}
}
