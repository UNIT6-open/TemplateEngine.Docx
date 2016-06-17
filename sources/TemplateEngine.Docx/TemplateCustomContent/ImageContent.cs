namespace TemplateEngine.Docx
{
	public class ImageContent : IContentItem
    {
        public ImageContent()
        {
            
        }

        public ImageContent(string name, byte[] binary)
        {
            Name = name;
            Binary = binary;
        }
   
        public string Name { get; set; }
        public byte[] Binary { get; set; }
    }
}
