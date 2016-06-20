using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace TemplateEngine.Docx.Processors
{
	internal class ImagesProcessor:IProcessor
	{
		private readonly ProcessContext _context;
		private ProcessResult _processResult;
		public ImagesProcessor(ProcessContext context)
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
			if (!(item is ImageContent))
			{
				_processResult = ProcessResult.NotHandledResult;
				return;
			}

			var field = item as ImageContent;

			// If there isn't a field with that name, add an error to the error string,
			// and continue with next field.
			if (contentControl == null)
			{
				_processResult.Errors.Add(String.Format("Field Content Control '{0}' not found.",
					field.Name));
				return;
			}


            var blip = contentControl.DescendantsAndSelf(A.blip).First();
            if (blip == null)
            {
                _processResult.Errors.Add(String.Format("Image to replace for '{0}' not found.",
                    field.Name));
                return;
            }

            var imageId = blip.Attribute(R.embed).Value;

            var imagePart = (ImagePart)_context.Document.GetPartById(imageId);

			if (imagePart == null)
			{
				_processResult.Errors.Add(String.Format("Image to replace for '{0}' not found.",
				   field.Name));
				return;
			}

			imagePart.GetStream().SetLength(0);
            using (var writer = new BinaryWriter(imagePart.GetStream()))
            {
                writer.Write(field.Binary);               
            }                           
        }
	}
}
