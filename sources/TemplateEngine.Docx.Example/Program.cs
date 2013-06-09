using System;
using System.Collections.Generic;
using System.IO;
using TemplateEngine.Docx;

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
                    new FieldContent("ReportDate", DateTime.Now.ToShortDateString()),
                    new FieldContent("Count", "2"),
                },
                Tables = new List<TableContent>
                {
                    new TableContent
                    (
                        "Team Members",
                        new TableRowContent
                        (
                            new FieldContent("Name", "Eric"),
                            new FieldContent("Title", "Program Manager")
                        ),
                        new TableRowContent
                        (
                            new FieldContent("Name", "Bob"),
                            new FieldContent("Title", "Developer")
                        )
                    )
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
