using System;

namespace TemplateEngine.Docx
{
	[ContentItemName("Field")]
	public class FieldContent : HiddenContent<FieldContent>, IEquatable<FieldContent>
	{
        public FieldContent()
        {
            
        }

        public FieldContent(string name, string value)
        {
            Name = name;
            Value = value;
        }
   
	    public string Value { get; set; }

	    #region Equals

        public bool Equals(FieldContent other)
		{
			if (other == null) return false;

			return Name.Equals(other.Name) &&
			       Value.Equals(other.Value);
		}

		public override bool Equals(IContentItem other)
		{
			if (!(other is FieldContent)) return false;

			return Equals((FieldContent)other);
		}

		public override int GetHashCode()
		{
			return new { Name, Value }.GetHashCode();
		}

        #endregion
    }
}
