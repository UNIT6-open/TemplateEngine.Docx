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
				new ListContent("Scientists")
					.AddItem(new FieldContent("Name", "Nicola Tesla"),
						new ImageContent("Photo", File.ReadAllBytes("Tesla.jpg")),
						new FieldContent("Dates of life", string.Format("{0}-{1}", 
							1856, 1943)),
						new FieldContent("Info",
							"Serbian American inventor, electrical engineer, mechanical engineer, physicist, and futurist best known for his contributions to the design of the modern alternating current (AC) electricity supply system"))
					.AddItem(new FieldContent("Name", "Thomas Edison"),
						new ImageContent("Photo", File.ReadAllBytes("Edison.jpg")),
						new FieldContent("Dates of life", string.Format("{0}-{1}",
							1847, 1931)),
						new FieldContent("Info",
							"American inventor and businessman. He developed many devices that greatly influenced life around the world, including the phonograph, the motion picture camera, and the long-lasting, practical electric light bulb."))
					.AddItem(new FieldContent("Name", "Albert Einstein"),
						new ImageContent("Photo", File.ReadAllBytes("Einstein.jpg")),
						new FieldContent("Dates of life", string.Format("{0}-{1}",
							1879, 1955)),
						new FieldContent("Info",
							"German-born theoretical physicist. He developed the general theory of relativity, one of the two pillars of modern physics (alongside quantum mechanics). Einstein's work is also known for its influence on the philosophy of science. Einstein is best known in popular culture for his mass–energy equivalence formula E = mc2 (which has been dubbed 'the world's most famous equation')."))
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
