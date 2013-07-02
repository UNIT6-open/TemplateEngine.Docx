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
```

This example shows how you can transform this template:
![docx template](http://unit6.ru/img/template-engine/Word-Template-0.png)

Into this document:
![generated docx document](http://unit6.ru/img/template-engine/Word-Template-1.png)

All data binds on content control tag value (you should switch to developer ribbon tab, if you can't see it, go to settings of ribbon menu and enable developer tab):
![docx template settings](http://unit6.ru/img/template-engine/Word-Template-2.png)

Have fun!

<!-- Yandex.Metrika counter -->
<div style="display:none;"><script type="text/javascript">
(function(w, c) {
	(w[c] = w[c] || []).push(function() {
		try {
			w.yaCounter9260296 = new Ya.Metrika({id:9260296, enableAll: true, webvisor:true});
		}
		catch(e) { }
	});
})(window, "yandex_metrika_callbacks");
</script></div>
<script src="//mc.yandex.ru/metrika/watch.js" type="text/javascript" defer="defer"></script>
<noscript><div><img src="//mc.yandex.ru/watch/9260296" style="position:absolute; left:-9999px;" alt="" /></div></noscript>
<!-- /Yandex.Metrika counter -->
