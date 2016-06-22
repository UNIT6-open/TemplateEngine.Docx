using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using TemplateEngine.Docx.Errors;

namespace TemplateEngine.Docx.Processors
{
	internal class ProcessResult
	{
		protected ProcessResult(bool handled=true)
		{
			_errors = new List<IError>();
			HandledItems = new Collection<IContentItem>();
			Handled = handled;
		}

		public static ProcessResult SuccessResult
		{
			get
			{
				return new ProcessResult
				{
					Handled = true
				};
			}
		}
		public static ProcessResult ErrorResult(IEnumerable<IError> errors)
		{
			return new ProcessResult
			{
				Handled = true,
				_errors = errors.ToList()
			};
		}

		public static ProcessResult NotHandledResult
		{
			get
			{
				return new ProcessResult
				{
					Handled = false
				};
			}
		}

		private List<IError> _errors; 
		public ReadOnlyCollection<IError> Errors { get{return new ReadOnlyCollection<IError>(_errors);}}
		public bool Success { get { return Handled && !Errors.Any(); } }
		public bool Handled { get; private set; }

		public ICollection<IContentItem> HandledItems { get; private set; }

		public ProcessResult AddItemToHandled(IContentItem handledItem)
		{
			if (!HandledItems.Contains(handledItem))
				HandledItems.Add(handledItem);

			var contentControlNotFoundErrors = Errors.OfType<ContentControlNotFoundError>()
				.Where(x => x.ContentItem.Equals(handledItem))
				.ToList();

			foreach (var error in contentControlNotFoundErrors)
			{
				_errors.Remove(error);
			}
			
			Handled = true;

			return this;
		}

		public ProcessResult AddError(IError error)
		{
			if (_errors.Contains(error))
				return this;

			var foundError = error as ContentControlNotFoundError;
			if (foundError != null)
			{
				if (HandledItems.Contains(foundError.ContentItem))
					return this;
			}

			_errors.Add(error);
			return this;
		}

		public ProcessResult Merge(ProcessResult another)
		{
			if (another == null)
			{
				return this;
			}

			if (!another.Success)
			{
				foreach (var error in another.Errors)
				{
					AddError(error);
				}
			}

			foreach (var handledItem in another.HandledItems)
			{
				AddItemToHandled(handledItem);
			}

			Handled = Handled || another.Handled;

			return this;
		}
	}
}
