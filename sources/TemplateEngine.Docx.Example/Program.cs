using System.IO;

namespace TemplateEngine.Docx.Example
{
    class Program
    {
        static void Main(string[] args)
        {
			var valuesToFill = new Content(
				new TableContent("Products")
				.AddRow(
					new FieldContent("Category", "Fruits"),
					new ListContent("Items")
						.AddItem(new ListItemContent("Item", "Orange")
							.AddNestedItem(new ListItemContent("Color", "Orange")))
						.AddItem(new ListItemContent("Item", "Apple")
							.AddNestedItem(new ListItemContent("Color", "Green"))
							.AddNestedItem(new ListItemContent("Color", "Red"))))					
				.AddRow(
					new FieldContent("Category", "Vegetables"),
					new ListContent("Items")
						.AddItem(new ListItemContent("Item", "Tomato")
							.AddNestedItem(new ListItemContent("Color", "Yellow"))
							.AddNestedItem(new ListItemContent("Color", "Red")))
						.AddItem(new ListItemContent("Item", "Cabbage"))));
				
							
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
