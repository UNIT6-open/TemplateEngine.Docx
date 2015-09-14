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
                Lists = new List<ListContent>
				{
					new ListContent("Team List", 
						new FieldContent("Team", "First team", 
							new Content
							{
								Tables = new List<TableContent>
								{
									new TableContent
									(
										"Team Members",
										new TableRowContent
										(
											new FieldContent("Name", "Eric"),
											new FieldContent("Role", "Program Manager")
										),
										new TableRowContent
										(
											new FieldContent("Name", "Bob"),
											new FieldContent("Role", "Developer")
										)
									)
								}
							}), 
						new FieldContent("Team", "Second team", 
							new Content
							{
								Tables = new List<TableContent>
								{
									new TableContent
									(
										"Team Members",
										new TableRowContent
										(
											new FieldContent("Name", "Mark"),
											new FieldContent("Role", "Team Leed")
										),
										new TableRowContent
										(
											new FieldContent("Name", "David"),
											new FieldContent("Role", "Developer")
										)
									)
								}
							}))
				},
                
            };

            File.Delete("OutputDocument.docx");
            File.Copy("InputTemplate.docx", "OutputDocument.docx");

            using(var outputDocument = new TemplateProcessor("OutputDocument.docx").SetRemoveContentControls(true))
            {
                outputDocument.FillContent(valuesToFill);
                outputDocument.SaveChanges();
            }
        }
    }
}
