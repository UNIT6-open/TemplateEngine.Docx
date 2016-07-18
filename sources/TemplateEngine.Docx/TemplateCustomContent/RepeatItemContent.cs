using System;

namespace TemplateEngine.Docx
{
    public class RepeatItemContent : Container, IEquatable<RepeatItemContent>
    {
        public RepeatItemContent()
        {
            
        }

		public RepeatItemContent(params IContentItem[] contentItems)
			: base(contentItems)
		{
			
		}

		#region Equals

		public bool Equals(RepeatItemContent other)
	    {
		    return base.Equals(other);
	    }

	    public override int GetHashCode()
	    {
		    return base.GetHashCode();
		}

		#endregion
	}
}
