using System.Collections.Generic;

namespace TemplateEngine.Docx
{
    public class TableContent
    {
        public TableContent()
        {
            
        }

        public TableContent(string name)
        {
            Name = name;
        }

        public TableContent(string name, IEnumerable<TableRowContent> rows)
            : this(name)
        {
            Rows = rows;
        }

        public TableContent(string name, params TableRowContent[] rows)
            : this(name)
        {
            Rows = rows;
        }

        public string Name { get; set; }
        public IEnumerable<TableRowContent> Rows { get; set; }
    }
}
