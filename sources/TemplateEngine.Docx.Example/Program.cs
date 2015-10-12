using System;
using System.IO;

namespace TemplateEngine.Docx.Example
{
    class Program
    {
        static void Main(string[] args)
        {				
            File.Delete("OutputDocument.docx");
            File.Copy("InputTemplate.docx", "OutputDocument.docx");

	        var valuesToFill = new Content(
		        // Add field.
		        new FieldContent("Report date", DateTime.Now.ToString()),

		        // Add table.
		        new TableContent("Team Members Table")
			        .AddRow(
				        new FieldContent("Name", "Eric"),
				        new FieldContent("Role", "Program Manager"))
			        .AddRow(
				        new FieldContent("Name", "Bob"),
				        new FieldContent("Role", "Developer")),

				// Add field inside table that not to propagate.
		        new FieldContent("Count", "2"),

				// Add list.	
				new ListContent("Team Members List")
					.AddItem(new ListItemContent("Name", "Eric").AddField("Role", "Program Manager"))
					.AddItem(new ListItemContent("Name", "Bob").AddField("Role", "Developer")),

				// Add nested list.	
				new ListContent("Team Members Nested List")
					.AddItem(new ListItemContent("Role", "Program Manager")
						.AddNestedItem(new ListItemContent("Name", "Eric"))
						.AddNestedItem(new ListItemContent("Name", "Ann")))
					.AddItem(new ListItemContent("Role", "Developer")
						.AddNestedItem(new ListItemContent("Name", "Bob"))
						.AddNestedItem(new ListItemContent("Name", "Richard"))),

				// Add list inside table.	
				new TableContent("Projects Table")
					.AddRow(
						new FieldContent("Name", "Eric"), 
						new FieldContent("Role", "Program Manager"), 
						new ListContent("Projects")
							.AddItem(new ListItemContent("Project", "Project one"))
							.AddItem(new ListItemContent("Project", "Project two")))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer"),
						new ListContent("Projects")
							.AddItem(new ListItemContent("Project", "Project one"))
							.AddItem(new ListItemContent("Project", "Project three"))),
		      
				// Add table inside list.	
				new ListContent("Projects List")
					.AddItem(new ListItemContent("Project", "Project one")
						.AddTable(TableContent.Create("Team members")
							.AddRow(
								new FieldContent("Name", "Eric"), 
								new FieldContent("Role", "Program Manager"))
							.AddRow(
								new FieldContent("Name", "Bob"), 
								new FieldContent("Role", "Developer"))))
					.AddItem(new ListItemContent("Project", "Project two")
						.AddTable(TableContent.Create("Team members")
							.AddRow(
								new FieldContent("Name", "Eric"),
								new FieldContent("Role", "Program Manager"))))
					.AddItem(new ListItemContent("Project", "Project three")
						.AddTable(TableContent.Create("Team members")
							.AddRow(
								new FieldContent("Name", "Bob"),
								new FieldContent("Role", "Developer")))));

            using(var outputDocument = new TemplateProcessor("OutputDocument.docx")
				.SetRemoveContentControls(true))
            {
                outputDocument.FillContent(valuesToFill);
                outputDocument.SaveChanges();
            }
        }
    }
}
