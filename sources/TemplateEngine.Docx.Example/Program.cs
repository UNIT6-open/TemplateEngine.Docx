using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using TemplateEngine.Docx.TemplateCustomContent;

namespace TemplateEngine.Docx.Example
{
    class Program
    {
        static void Main(string[] args)
        {
            var valuesToFill = new Content
            {
                Fields = new List<FieldContent>
                {
                    new FieldContent { Name = "ReportDate", Value = DateTime.Now.ToShortDateString() },
                    new FieldContent { Name = "Count", Value = "2" },
                },
                Tables = new List<TableContent>
                {
                    new TableContent 
                    {
                        Name = "Team Members",
                        Rows = new List<TableRowContent>
                        {
                            new TableRowContent
                            {
                                Fields = new List<FieldContent>
                                    {
                                        new FieldContent { Name = "Name", Value = "Eric" },
                                        new FieldContent { Name = "Title", Value = "Program Manager" }
                                    }
                            },
                            new TableRowContent
                            {
                                Fields = new List<FieldContent>
                                    {
                                        new FieldContent { Name = "Name", Value = "Bob" },
                                        new FieldContent { Name = "Title", Value = "Developer" }
                                    }
                            },
                        }
                    }
                }
            };

            File.Delete("OutputDocument.docx");
            File.Copy("InputTemplate.docx", "OutputDocument.docx");
            new TemplateProcessor("OutputDocument.docx")
                .FillContent(valuesToFill)
                .SaveChanges();
        }
    }
}
