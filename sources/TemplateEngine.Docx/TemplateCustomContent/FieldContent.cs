namespace TemplateEngine.Docx
{
	public class FieldContent : IContentItem
    {
        public FieldContent()
        {
            
        }

        public FieldContent(string name, string value)
        {
            Name = name;
            Value = value;
        }
   
        public string Name { get; set; }
        public string Value { get; set; }
    }
}
