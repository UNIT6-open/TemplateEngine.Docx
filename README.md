# TemplateEngine.Docx

Template engine for generating Word docx documents on the server-side with a human-created Word templates, based on content controls Word feature.

Based on [Eric White code sample](http://msdn.microsoft.com/en-us/magazine/ee532473.aspx).

## Installation

You can get it with [NuGet package](https://nuget.org/packages/TemplateEngine.Docx/).
![nuget install command](http://unit6.ru/img/template-engine/NuGet-Install.png)

## Example of usage

[Here](https://bitbucket.org/unit6ru/templateengine/src/a3d49e1a2840b4c04939761901b50f2b8e6dc4ac/sources/TemplateEngine.Docx.Example?at=master) is source code of sample usage.
```
#!CSharp
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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
                    new FieldContent
					{
						Name = "ReportDate",
						Value = DateTime.Now.ToShortDateString()
					},
                    new FieldContent
					{
						Name = "Count",
						Value = "2"
					},
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
									new FieldContent
									{
										Name = "Name",
										Value = "Eric"
									},
									new FieldContent
									{
										Name = "Title",
										Value = "Program Manager"
									}
								}
                            },
                            new TableRowContent
                            {
                                Fields = new List<FieldContent>
								{
									new FieldContent
									{
										Name = "Name",
										Value = "Bob"
									},
									new FieldContent
									{
										Name = "Title",
										Value = "Developer"
									}
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
```

This example shows how you can transform this template:
![docx template](http://unit6.ru/img/template-engine/Word-Template-0.png)

Into this document:
![generated docx document](http://unit6.ru/img/template-engine/Word-Template-1.png)

All data binds on content control tag value (you should switch to developer ribbon tab, if you can't see it, go to settings of ribbon menu and enable developer tab):
![docx template settings](http://unit6.ru/img/template-engine/Word-Template-2.png)

Have fun!