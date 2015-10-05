using System.Collections.Generic;
using System.Linq;

namespace TemplateEngine.Docx.Processors
{
	public class ProcessResult
	{
		public ProcessResult(bool handled=true)
		{
			Errors = new List<string>();
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
		public static ProcessResult ErrorResult(IEnumerable<string> errors)
		{
			return new ProcessResult
			{
				Handled = true,
				Errors = errors.ToList()
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
		public List<string> Errors { get; set; }
		public bool Success { get { return Handled && !Errors.Any(); } }
		public bool Handled { get; private set; }
	}
}
