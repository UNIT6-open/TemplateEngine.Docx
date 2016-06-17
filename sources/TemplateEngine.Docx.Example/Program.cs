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
					.AddItem(
						new FieldContent("Name", "Eric"), 
						new FieldContent("Role", "Program Manager"))
					.AddItem(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer")),

				// Add nested list.	
				new ListContent("Team Members Nested List")
					.AddItem(new ListItemContent("Role", "Program Manager")
						.AddNestedItem(new FieldContent("Name", "Eric"))
						.AddNestedItem(new FieldContent("Name", "Ann")))
					.AddItem(new ListItemContent("Role", "Developer")
						.AddNestedItem(new FieldContent("Name", "Bob"))
						.AddNestedItem(new FieldContent("Name", "Richard"))),

				// Add list inside table.	
				new TableContent("Projects Table")
					.AddRow(
						new FieldContent("Name", "Eric"), 
						new FieldContent("Role", "Program Manager"), 
						new ListContent("Projects")
							.AddItem(new FieldContent("Project", "Project one"))
							.AddItem(new FieldContent("Project", "Project two")))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer"),
						new ListContent("Projects")
							.AddItem(new FieldContent("Project", "Project one"))
							.AddItem(new FieldContent("Project", "Project three"))),
		      
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
								new FieldContent("Role", "Developer")))),


				// Add table with several blocks.	
				new TableContent("Team Members Statistics")
					.AddRow(
						new FieldContent("Name", "Eric"),
						new FieldContent("Role", "Program Manager"))
					.AddRow(
					    new FieldContent("Name", "Richard"),
						new FieldContent("Role", "Program Manager"))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer")),

					new TableContent("Team Members Statistics")
					.AddRow(
						new FieldContent("Statistics Role", "Program Manager"),
						new FieldContent("Statistics Role Count", "2"))						
					.AddRow(
						new FieldContent("Statistics Role", "Developer"),
						new FieldContent("Statistics Role Count", "1")),
						
				// Add table with merged rows
				new TableContent("Team members info")
					.AddRow(
						new FieldContent("Name", "Eric"),
						new FieldContent("Role", "Program Manager"),
						new FieldContent("Age", "37"),
						new FieldContent("Gender", "Male"))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer"),
						new FieldContent("Age", "33"),
						new FieldContent("Gender", "Male"))
					.AddRow(
						new FieldContent("Name", "Ann"),
						new FieldContent("Role", "Developer"),
						new FieldContent("Age", "34"),
						new FieldContent("Gender", "Female")),
						
				// Add table with merged columns
				new TableContent("Team members projects")
					.AddRow(
						new FieldContent("Name", "Eric"),
						new FieldContent("Role", "Program Manager"),
						new FieldContent("Age", "37"),
						new FieldContent("Projects", "Project one, Project two"))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer"),
						new FieldContent("Age", "33"),
						new FieldContent("Projects", "Project one"))
					.AddRow(
						new FieldContent("Name", "Ann"),
						new FieldContent("Role", "Developer"),
						new FieldContent("Age", "34"),
						new FieldContent("Projects", "Project two")),
                new ImageContent("photo", File.ReadAllBytes("Tesla.jpg"))
            );

            using(var outputDocument = new TemplateProcessor("OutputDocument.docx")
				.SetRemoveContentControls(true))
            {
                outputDocument.FillContent(valuesToFill);
                outputDocument.SaveChanges();
            }
        }
    }
}
