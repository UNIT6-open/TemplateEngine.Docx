using System.Collections.Generic;

namespace TemplateEngine.Docx.TemplateCustomContent
{
    public class Content
    {
        public IEnumerable<TableContent> Tables { get; set; }
        public IEnumerable<FieldContent> Fields { get; set; }
    }
}
