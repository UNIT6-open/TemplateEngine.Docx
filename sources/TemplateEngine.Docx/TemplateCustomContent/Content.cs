using System.Collections.Generic;

namespace TemplateEngine.Docx
{
    public class Content
    {
        public IEnumerable<TableContent> Tables { get; set; }
        public IEnumerable<ListContent> Lists { get; set; }
        public IEnumerable<FieldContent> Fields { get; set; }
    }
}
