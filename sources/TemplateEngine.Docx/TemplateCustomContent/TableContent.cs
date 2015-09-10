using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

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

		public IEnumerable<string> FieldNames { get
		{
			return Rows == null ? new List<string>() : Rows.SelectMany(r => r.Fields.Select(f => f.Name)).Distinct().ToList();
		} } 
    }
}
