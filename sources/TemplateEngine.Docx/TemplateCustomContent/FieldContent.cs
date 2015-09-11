namespace TemplateEngine.Docx
{
    public class FieldContent
    {
        public FieldContent()
        {
            
        }

        public FieldContent(string name, string value)
        {
            Name = name;
            Value = value;
        }
        public FieldContent(string name, string value, Content content)
        {
            Name = name;
            Value = value;

	        Content = content;
        }

        public string Name { get; set; }
        public string Value { get; set; }

		public Content Content { get; set; }
    }
}
