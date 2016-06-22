using System;

namespace TemplateEngine.Docx.Errors
{
	internal interface IError:IEquatable<IError>
	{
		string Message { get; }
	}
}
