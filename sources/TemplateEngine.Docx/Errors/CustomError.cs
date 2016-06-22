using System;

namespace TemplateEngine.Docx.Errors
{
	internal class CustomError:IError, IEquatable<CustomError>
	{
		internal CustomError(string customMessage)
		{
			_customMessage = customMessage;
		}

		private readonly string _customMessage;
		public string Message
		{
			get
			{
				return _customMessage;
			}
		}

		#region Equals
		public bool Equals(IError other)
		{
			if (!(other is CustomError))
				return false;

			return Equals((CustomError) other);
		}
		
		public bool Equals(CustomError other)
		{
			return Message.Equals(other.Message);
		}

		public override int GetHashCode()
		{
			return Message.GetHashCode();
		}
		#endregion
	}
}
