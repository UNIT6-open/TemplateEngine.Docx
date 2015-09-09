using System.Collections.Generic;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TemplateEngine.Docx.Tests.Properties;

namespace TemplateEngine.Docx.Tests
{
    [TestClass]
    public class TemplateProcessorTests
    {
        [TestMethod]
        public void FillingOneTableWithTwoRowsAndPreserveContentControls()
        {
            var templateDocument = XDocument.Parse(Resources.TemplateWithSingleTable);
            var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleTableFilledWithTwoRows);

            var valuesToFill = new Content
            {
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

            var template = new TemplateProcessor(templateDocument)
                .FillContent(valuesToFill);

            var documentXml = template.Document.ToString();

            Assert.AreEqual(expectedDocument.Document.ToString(), documentXml);
        }

		[TestMethod]
		public void FillingOneTableWithTwoRowsAndRemoveContentControls()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithSingleTable);
			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleTableFilledWithTwoRowsAndRemovedCC);

			var valuesToFill = new Content
			{
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

			var template = new TemplateProcessor(templateDocument).SetRemoveContentControls(true)
				.FillContent(valuesToFill);

			var documentXml = template.Document.ToString();

			Assert.AreEqual(expectedDocument.Document.ToString(), documentXml);
		}

        [TestMethod]
        public void FillingOneFieldWithValue()
        {
            var templateDocument = XDocument.Parse(Resources.TemplateWithSingleField);
            var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleFieldFilled);

            var valuesToFill = new Content
            {
                Fields = new List<FieldContent>
                {
                    new FieldContent { Name = "ReportDate", Value = "09.06.2013" }
                }
            };

            var template = new TemplateProcessor(templateDocument)
                .FillContent(valuesToFill);

            var documentXml = template.Document.ToString();

            Assert.AreEqual(expectedDocument.Document.ToString(), documentXml);
        }

		[TestMethod]
		public void FillingOneFieldWithValueAndRemoveContentControl()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithSingleField);
			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleFieldAndRemovedCC);

			var valuesToFill = new Content
			{
				Fields = new List<FieldContent>
                {
                    new FieldContent { Name = "ReportDate", Value = "09.06.2013" }
                }
			};

			var template = new TemplateProcessor(templateDocument)
							.SetRemoveContentControls(true)
							.FillContent(valuesToFill);

			var documentXml = template.Document.ToString();

			Assert.AreEqual(expectedDocument.Document.ToString(), documentXml);
		}

        [TestMethod]
        public void FillingOneFieldWithWrongValue_WillNoticeWithWarning()
        {
            var templateDocument = XDocument.Parse(Resources.TemplateWithSingleField);
            var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleFieldWrongFilled);

            var valuesToFill = new Content
            {
                Fields = new List<FieldContent>
                {
                    new FieldContent { Name = "WrongReportDate", Value = "09.06.2013" }
                }
            };

            var template = new TemplateProcessor(templateDocument)
                            .FillContent(valuesToFill);

            var documentXml = template.Document.ToString();

            Assert.AreEqual(expectedDocument.Document.ToString(), documentXml);
        }
		[TestMethod]
		public void FillingFieldInTableHeaderWithValue()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithFieldInTableHeader);
			var expectedDocument = XDocument.Parse(Resources.DocumentWithFieldInTableHeaderFilled);

			var valuesToFill = new Content
			{
				Fields = new List<FieldContent>
                {
                    new FieldContent { Name = "Count", Value = "2" }
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

			var template = new TemplateProcessor(templateDocument)
				.FillContent(valuesToFill);

			var documentXml = template.Document.ToString();

			Assert.AreEqual(expectedDocument.Document.ToString(), documentXml);
		}

		[TestMethod]
		public void FillingOneTableWithTwoRowsWithWrongValues_WillNoticeWithWarning()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithSingleTable);
			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleTableWrongFilled);

			var valuesToFill = new Content
			{
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
                                        new FieldContent { Name = "Title", Value = "Program Manager" },
										new FieldContent {Name = "WrongFieldName", Value = "Value"}
                                    }
                            },
                            new TableRowContent
                            {
                                Fields = new List<FieldContent>
                                    {
                                        new FieldContent { Name = "Name", Value = "Bob" },
                                        new FieldContent { Name = "Title", Value = "Developer" },
										new FieldContent {Name = "WrongFieldName", Value = "Value"}
                                    }
                            },
                        }
                    }
                }
			};

			var template = new TemplateProcessor(templateDocument)
				.FillContent(valuesToFill);

			var documentXml = template.Document.ToString();

			Assert.AreEqual(expectedDocument.Document.ToString(), documentXml);
		}
    }
}
