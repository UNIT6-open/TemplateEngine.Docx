using System.Collections.Generic;

namespace TemplateEngine.Docx
{
    public class TableRowContent
    {
        public TableRowContent()
        {
            
        }

        public TableRowContent(IEnumerable<FieldContent> fields)
        {
            Fields = fields;
        }

        public TableRowContent(params FieldContent[] fields)
        {
            Fields = fields;
        }

        public IEnumerable<FieldContent> Fields { get; set; }
    }
}
