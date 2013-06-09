using System.Collections.Generic;

namespace TemplateEngine.Docx.TemplateCustomContent
{
    public class TableContent
    {
        public string Name { get; set; }
        public IEnumerable<TableRowContent> Rows { get; set; }
    }
}
