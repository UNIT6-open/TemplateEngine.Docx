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


		[TestMethod]
		public void FillingOneFieldWithSeveralTextEntries()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithSingleFieldWithSeveralTextEntries);
			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleFieldWithSeveralTextEntriesFilled);

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
		public void FillingOneTableWithAdjacentRows()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithSingleTableWithAdjacentRows);
			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleTableWithAdjacentRowsFilled);

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
										new FieldContent { Name = "Age", Value = "33" },
										new FieldContent { Name = "Gender", Value = "Male" },
										new FieldContent { Name = "Comment", Value = "" }
                                    }
                            },
                            new TableRowContent
                            {
                                 Fields = new List<FieldContent>
                                    {
                                        new FieldContent { Name = "Name", Value = "Bob" },
                                        new FieldContent { Name = "Title", Value = "Developer" },
										new FieldContent { Name = "Age", Value = "51" },
										new FieldContent { Name = "Gender", Value = "Male" },
										new FieldContent { Name = "Comment", Value = "Retiral" }
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
		public void FillingOneTableWithMergedVerticallyRows()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithSingleTableWithMergedVerticallyRows);
			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleTableWithMergedVerticallyRowsFilled);

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
                                        new FieldContent { Name = "Name", Value = "Eric" }
                                    }
                            },
                            new TableRowContent
                            {
                                 Fields = new List<FieldContent>
                                    {
                                        new FieldContent { Name = "Name", Value = "Bob" }                                     
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
		public void FillingOneListAndPreserveContentControl()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithSingleList);
			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleListFilled);

			var valuesToFill = new Content
			{
				Lists = new List<ListContent>
                {
                    new ListContent 
                    {
                        Name = "Food Items",
                        Items = new List<FieldContent>
                        {                   
                             new FieldContent { Name = "Category", Value = "Fruit" },
                             new FieldContent { Name = "Category", Value = "Vegetables" }   
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
		public void FillingOneListAndRemoveContentControl()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithSingleList);
			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleListFilledAndRemovedCC);

			var valuesToFill = new Content
			{
				Lists = new List<ListContent>
                {
                    new ListContent 
                    {
                        Name = "Food Items",
                        Items = new List<FieldContent>
                        {                   
                              new FieldContent { Name = "Category", Value = "Fruit" },
                              new FieldContent { Name = "Category", Value = "Vegetables" }   
                        }
                    }
                }
			};

			var template = new TemplateProcessor(templateDocument)
				.SetRemoveContentControls(true)
				.FillContent(valuesToFill);

			var documentXml = template.Document.ToString();

			Assert.AreEqual(expectedDocument.Document.ToString(), documentXml);
		}		
		[TestMethod]
		public void FillingOneNestedListAndPreserveContentControl()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithSingleNestedList);
			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleNestedListFilled);

			var valuesToFill = new Content
			{
				Lists = new List<ListContent>
               {
	               new ListContent("Document", 
					   new FieldContent ("Header", "Introduction"),
					   new ListItemContent
					   {
						   Name = "Header",Value = "Chapter 1 - The new start screen", 
						   NestedFields= new []
						   {
								new ListItemContent
								{
									Name = "Subheader", Value ="What's new in Windows 8?",	
									NestedFields = new List<FieldContent>
									{
										new FieldContent("Paragraph", "Chapter content")
									}
								},
								new FieldContent("Subheader", "Starting Windows 8")
													
						   }
					   },
					   new ListItemContent
					   {
						   Name = "Header",Value = "Chapter 2 - The traditional Desktop", 
						   NestedFields= new []
						   {
								new FieldContent("Subheader","Browsing the File Explorer"),
								new FieldContent("Subheader","Getting the Lowdown on Folders and Libraries")
						   }
					   })					  	 					  
               }

			};

			var template = new TemplateProcessor(templateDocument).FillContent(valuesToFill);

			var documentXml = template.Document.ToString();

			Assert.AreEqual(expectedDocument.Document.ToString(), documentXml);
		}
		[TestMethod]
		public void FillingOneNestedListAndRemoveContentControl()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithSingleNestedList);
			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleNestedListFilledAndRemovedCC);

			var valuesToFill = new Content
			{
				Lists = new List<ListContent>
               {
	               new ListContent("Document", 
					   new FieldContent ("Header", "Introduction"),
					   new ListItemContent
					   {
						   Name = "Header",Value = "Chapter 1 - The new start screen", 
						   NestedFields= new []
						   {
								new ListItemContent
								{
									Name = "Subheader", Value ="What's new in Windows 8?",	
									NestedFields = new List<FieldContent>
									{
										new FieldContent("Paragraph", "Chapter content")
									}
								},
								new FieldContent("Subheader", "Starting Windows 8")
													
						   }
					   },
					   new ListItemContent
					   {
						   Name = "Header",Value = "Chapter 2 - The traditional Desktop", 
						   NestedFields= new []
						   {
								new FieldContent("Subheader","Browsing the File Explorer"),
								new FieldContent("Subheader","Getting the Lowdown on Folders and Libraries")
						   }
					   })					  	 					  
               }

			};

			var template = new TemplateProcessor(templateDocument).SetRemoveContentControls(true).FillContent(valuesToFill);

			var documentXml = template.Document.ToString();

			Assert.AreEqual(expectedDocument.Document.ToString(), documentXml);
		}
    }
}
