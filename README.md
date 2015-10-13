# TemplateEngine.Docx

Template engine for generating Word docx documents on the server-side with a human-created Word templates, based on content controls Word feature.

Based on [Eric White code sample](http://msdn.microsoft.com/en-us/magazine/ee532473.aspx).

## Installation

You can get it with [NuGet package](https://nuget.org/packages/TemplateEngine.Docx/).

![nuget install command](http://unit6.ru/img/template-engine/NuGet-Install.png)

## Example of usage

[Here](https://bitbucket.org/unit6ru/templateengine/src/a3d49e1a2840b4c04939761901b50f2b8e6dc4ac/sources/TemplateEngine.Docx.Example?at=master) is source code of sample usage.

### How to create template

* Create Word document with prototype of content:  
![document example](file:///D:/Desktop/TemplateEngineDocx/Table-Example.PNG)  

* Switch to developer ribbon tab. If you can't see it, go to settings of ribbon menu and enable developer tab:  
![developer mode](file:///D:/Desktop/TemplateEngineDocx/Developer-Mode.PNG)

* All data binds on content control tag value:  
![create content control](http://unit6.ru/img/template-engine/Create-Content-Control.PNG)


###Filling fields

![docx template](http://unit6.ru/img/template-engine/Fill-Field.PNG)

```
#!CSharp
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
		        new FieldContent("Report date", DateTime.Now.ToString()));

		    using(var outputDocument = new TemplateProcessor("OutputDocument.docx")
				.SetRemoveContentControls(true))
            {
                outputDocument.FillContent(valuesToFill);
                outputDocument.SaveChanges();
            } 
		}
	}
}			
```
###Filling tables

![filling table](http://unit6.ru/img/template-engine/Fill-Table.PNG)

```
#!CSharp
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
		         new TableContent("Team Members Table")
			        .AddRow(
				        new FieldContent("Name", "Eric"),
				        new FieldContent("Role", "Program Manager"))
			        .AddRow(
				        new FieldContent("Name", "Bob"),
				        new FieldContent("Role", "Developer")));
		
		    using(var outputDocument = new TemplateProcessor("OutputDocument.docx")
				.SetRemoveContentControls(true))
            {
                outputDocument.FillContent(valuesToFill);
                outputDocument.SaveChanges();
            } 
		}
	}
}			
```
###Filling lists

![filling list](http://unit6.ru/img/template-engine/Fill-List.PNG)

```
#!CSharp
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
		        new ListContent("Team Members List")
					.AddItem(
						new FieldContent("Name", "Eric"), 
						new FieldContent("Role", "Program Manager"))
					.AddItem(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer")));

		    using(var outputDocument = new TemplateProcessor("OutputDocument.docx")
				.SetRemoveContentControls(true))
            {
                outputDocument.FillContent(valuesToFill);
                outputDocument.SaveChanges();
            } 
		}
	}
}			
```
###Filling nested lists

![filling nested list](http://unit6.ru/img/template-engine/Fill-Nested-List.PNG)

```
#!CSharp
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
		       	new ListContent("Team Members Nested List")
					.AddItem(new ListItemContent("Role", "Program Manager")
						.AddNestedItem(new FieldContent("Name", "Eric"))
						.AddNestedItem(new FieldContent("Name", "Ann")))
					.AddItem(new ListItemContent("Role", "Developer")
						.AddNestedItem(new FieldContent("Name", "Bob"))
						.AddNestedItem(new FieldContent("Name", "Richard"))));


		    using(var outputDocument = new TemplateProcessor("OutputDocument.docx")
				.SetRemoveContentControls(true))
            {
                outputDocument.FillContent(valuesToFill);
                outputDocument.SaveChanges();
            } 
		}
	}
}			
```
###Filling list inside table

![filling list inside table](http://unit6.ru/img/template-engine/Fill-List-Inside-Table.PNG)

```
#!CSharp
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
							.AddItem(new FieldContent("Project", "Project three"))));


		    using(var outputDocument = new TemplateProcessor("OutputDocument.docx")
				.SetRemoveContentControls(true))
            {
                outputDocument.FillContent(valuesToFill);
                outputDocument.SaveChanges();
            } 
		}
	}
}			
```
###Filling table inside list

![filling table inside list](http://unit6.ru/img/template-engine/Fill-Table-Inside-List.PNG)

```
#!CSharp
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
```

###Filling table with several blocks

![filling table with several blocks](http://unit6.ru/img/template-engine/Fill-Table-With-Several-Blocks.PNG)

```
#!CSharp
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
						new FieldContent("Statistics Role Count", "1")));


		    using(var outputDocument = new TemplateProcessor("OutputDocument.docx")
				.SetRemoveContentControls(true))
            {
                outputDocument.FillContent(valuesToFill);
                outputDocument.SaveChanges();
            } 
		}
	}
}			
```

###Filling table with merged rows

![filling table with merged rows](http://unit6.ru/img/template-engine/Fill-Table-With-Merged-Rows.PNG)

```
#!CSharp
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
						new FieldContent("Gender", "Female")));


		    using(var outputDocument = new TemplateProcessor("OutputDocument.docx")
				.SetRemoveContentControls(true))
            {
                outputDocument.FillContent(valuesToFill);
                outputDocument.SaveChanges();
            } 
		}
	}
}			
```

###Filling table with merged columns

![table with merged columns](http://unit6.ru/img/template-engine/Fill-Table-With-Merged-Columns.PNG)

```
#!CSharp
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
						new FieldContent("Projects", "Project two")));


		    using(var outputDocument = new TemplateProcessor("OutputDocument.docx")
				.SetRemoveContentControls(true))
            {
                outputDocument.FillContent(valuesToFill);
                outputDocument.SaveChanges();
            } 
		}
	}
}			
```

###Content structure
![content structure](http://unit6.ru/img/template-engine/Structure.PNG)

Have fun!

![yandex counter](//mc.yandex.ru/watch/9260296)