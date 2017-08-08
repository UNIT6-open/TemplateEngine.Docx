using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using TemplateEngine.Docx.Errors;

namespace TemplateEngine.Docx.Processors
{
	internal class ImagesProcessor:IProcessor
	{
		private readonly ProcessContext _context;
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
			var processResult = ProcessResult.NotHandledResult;

			foreach (var contentItem in items)
			{
				processResult.Merge(FillContent(contentControl, contentItem));
			}


			if (processResult.Success && _isNeedToRemoveContentControls)
				contentControl.RemoveContentControl();

			return processResult;
		}

		public ProcessResult FillContent(XElement contentControl, IContentItem item)
		{
			var processResult = ProcessResult.NotHandledResult; 

			if (!(item is ImageContent))
			{
				processResult = ProcessResult.NotHandledResult;
				return processResult;
			}

			var field = item as ImageContent;

			// If there isn't a field with that name, add an error to the error string,
			// and continue with next field.
			if (contentControl == null)
			{
				processResult.AddError(new ContentControlNotFoundError(field));
				return processResult;
			}

		    if (item.IsHidden)
		    {
		        var graphic = contentControl.DescendantsAndSelf(W.drawing).First();
		        graphic.Remove();
            }
		    else
		    {
		        var blip = contentControl.DescendantsAndSelf(A.blip).First();
		        if (blip == null)
		        {
		            processResult.AddError(new CustomContentItemError(field, "doesn't contain an image for replace"));
		            return processResult;
		        }

		        var imageId = blip.Attribute(R.embed).Value;

		        var imagePart = (ImagePart)_context.Document.GetPartById(imageId);

		        if (imagePart != null)
		        {
		            _context.Document.RemovePartById(imageId);
		        }

		        var imagePartId = _context.Document.AddImagePart(field.Binary);

		        blip.Attribute(R.embed).SetValue(imagePartId);
            }                    

			processResult.AddItemToHandled(item);
			return processResult;
		}
	}
}
