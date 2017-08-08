using System;
using System.Linq;

namespace TemplateEngine.Docx
{
	[ContentItemName("Image")]
	public class ImageContent : HiddenContent<ImageContent>, IEquatable<ImageContent>
    {
        public ImageContent()
        {
            
        }

        public ImageContent(string name, byte[] binary)
        {
            Name = name;
            Binary = binary;
        }

        public byte[] Binary { get; set; }

        #region Equals

        public bool Equals(ImageContent other)
		{
			if (other == null) return false;

			return Name.Equals(other.Name, StringComparison.InvariantCultureIgnoreCase) &&
			       Binary.SequenceEqual(other.Binary);
		}

		public override bool Equals(IContentItem other)
		{
			if (!(other is ImageContent)) return false;

			return Equals((ImageContent)other);
		}

		public override int GetHashCode()
		{
			return new {Name, Binary}.GetHashCode();
		}

		#endregion
	}
}
