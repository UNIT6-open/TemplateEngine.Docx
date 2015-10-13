using System.Collections.Generic;
using System.Text.RegularExpressions;
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
		public void FillingOneFieldWithWrongValueAndDisabledErrorsNotifications_NotWillNoticeWithWarning()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithSingleField);
			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleFieldWrongFilledWithoutErrorsNotifications);

			var valuesToFill = new Content
			{
				Fields = new List<FieldContent>
                {
                    new FieldContent { Name = "WrongReportDate", Value = "09.06.2013" }
                }
			};

			var template = new TemplateProcessor(templateDocument)
				.SetNoticeAboutErrors(false)
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

			var valuesToFill = new Content(
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
                );

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

			var valuesToFill = new Content(
				new TableContent("Team Members")
					.AddRow(new FieldContent ("Name", "Eric"))
					.AddRow(new FieldContent("Name", "Bob")));
			
			var template = new TemplateProcessor(templateDocument)
				.FillContent(valuesToFill);

			var documentXml = template.Document.ToString();

			Assert.AreEqual(expectedDocument.Document.ToString(), documentXml);
		}		
		
		[TestMethod]
		public void FillingOneListAndPreserveContentControl()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithSingleList_document);
			var templateStyles = XDocument.Parse(Resources.TemplateWithSingleList_styles);
			var templateNumbering = XDocument.Parse(Resources.TemplateWithSingleList_numbering);

			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleListFilled_document);
			var expectedStyles = XDocument.Parse(Resources.DocumentWithSingleListFilled_styles);
			var expectedNumbering = XDocument.Parse(Resources.DocumentWithSingleListFilled_numbering);

			var valuesToFill = new Content
			{
				Lists = new List<ListContent>
                {
                    new ListContent 
                    {
                        Name = "Food Items",
                        Items = new List<ListItemContent>
                        {                   
                             new ListItemContent ("Category", "Fruit"),
                             new ListItemContent ("Category", "Vegetables")
                        }
                    }
                }
			};

			var filledDocument = new TemplateProcessor(templateDocument, templateStyles, templateNumbering)
				.FillContent(valuesToFill);

			Assert.AreEqual(expectedDocument.ToString(), filledDocument.Document.ToString());
			Assert.AreEqual(expectedStyles.ToString(), filledDocument.StylesPart.ToString());
			Assert.AreEqual(expectedNumbering.ToString(), filledDocument.NumberingPart.ToString());
		}		
		[TestMethod]
		public void FillingOneListAndRemoveContentControl()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithSingleList_document);
			var templateStyles = XDocument.Parse(Resources.TemplateWithSingleList_styles);
			var templateNumbering = XDocument.Parse(Resources.TemplateWithSingleList_numbering);

			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleListFilledAndRemovedCC_document);
			var expectedStyles = XDocument.Parse(Resources.DocumentWithSingleListFilledAndRemovedCC_styles);
			var expectedNumbering = XDocument.Parse(Resources.DocumentWithSingleListFilledAndRemovedCC_numbering);

			var valuesToFill = new Content
			{
				Lists = new List<ListContent>
                {
                    new ListContent 
                    {
                        Name = "Food Items",
                       Items = new List<ListItemContent>
                        {                   
                             new ListItemContent ("Category", "Fruit"),
                             new ListItemContent ("Category", "Vegetables")
                        }
                    }
                }
			};

			var filledDocument = new TemplateProcessor(templateDocument, templateStyles, templateNumbering)
				.SetRemoveContentControls(true)
				.FillContent(valuesToFill);

			Assert.AreEqual(expectedDocument.ToString(), filledDocument.Document.ToString());
			Assert.AreEqual(expectedStyles.ToString(), filledDocument.StylesPart.ToString());
			Assert.AreEqual(expectedNumbering.ToString(), filledDocument.NumberingPart.ToString());
		}

		[TestMethod]
		public void FillingOneListWithWrongValues_WillNoticeWithWarning()
		{

			var templateDocument = XDocument.Parse(Resources.TemplateWithSingleList_document);
			var templateStyles = XDocument.Parse(Resources.TemplateWithSingleList_styles);
			var templateNumbering = XDocument.Parse(Resources.TemplateWithSingleList_numbering);

			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleListWrongFilled_document);
			var expectedStyles = XDocument.Parse(Resources.DocumentWithSingleListWrongFilled_styles);
			var expectedNumbering = XDocument.Parse(Resources.DocumentWithSingleListWrongFilled_numbering);

			var valuesToFill = new Content
			{
				Lists = new List<ListContent>
                {
                    new ListContent 
                    {
                        Name = "Food Items",
						Items = new List<ListItemContent>
                        {                   
                             new ListItemContent ("WrongListItem", "Fruit")
                        }
                    }
                }
			};

			var filledDocument = new TemplateProcessor(templateDocument, templateStyles, templateNumbering)
				.SetRemoveContentControls(true)
				.FillContent(valuesToFill);

			Assert.AreEqual(expectedDocument.ToString(), filledDocument.Document.ToString());
			Assert.AreEqual(expectedStyles.ToString(), filledDocument.StylesPart.ToString());
			Assert.AreEqual(expectedNumbering.ToString(), filledDocument.NumberingPart.ToString());
		}




		[TestMethod]
		public void FillingOneNestedListAndPreserveContentControl()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithSingleNestedList_document);
			var templateStyles = XDocument.Parse(Resources.TemplateWithSingleNestedList_styles);
			var templateNumbering = XDocument.Parse(Resources.TemplateWithSingleNestedList_numbering);

			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleNestedListFIlled_document);
			var expectedStyles = XDocument.Parse(Resources.DocumentWithSingleNestedListFIlled_styles);
			var expectedNumbering = XDocument.Parse(Resources.DocumentWithSingleNestedListFIlled_numbering);

			var valuesToFill = new Content(
				ListContent.Create("Document",
					ListItemContent.Create("Header", "Introduction"), 

					ListItemContent.Create("Header", "Chapter 1 - The new start screen")
						.AddField("Header text", "Header 2 paragraph text")
						.AddNestedItem(ListItemContent.Create("Subheader", "What's new in Windows 8?")
							.AddField("Subheader text", "Subheader 2.1 paragraph text"))
						.AddNestedItem(ListItemContent.Create("Subheader", "Starting Windows 8")),

					ListItemContent.Create("Header", "Chapter 2 - The traditional Desktop")
						.AddNestedItem(ListItemContent.Create("Subheader", "Browsing the File Explorer"))
						.AddNestedItem(ListItemContent.Create("Subheader", "Getting the Lowdown on Folders and Libraries"))));

			
			var filledDocument = new TemplateProcessor(templateDocument, templateStyles, templateNumbering)
				.SetRemoveContentControls(false)
				.FillContent(valuesToFill);

			Assert.AreEqual(expectedDocument.ToString(), filledDocument.Document.ToString());
			Assert.AreEqual(expectedStyles.ToString(), filledDocument.StylesPart.ToString());
			Assert.AreEqual(expectedNumbering.ToString(), filledDocument.NumberingPart.ToString());
		}
		[TestMethod]
		public void FillingOneNestedListAndRemoveContentControl()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithSingleNestedList_document);
			var templateStyles = XDocument.Parse(Resources.TemplateWithSingleNestedList_styles);
			var templateNumbering = XDocument.Parse(Resources.TemplateWithSingleNestedList_numbering);

			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleNestedListFilledAndRemovedCC_document);
			var expectedStyles = XDocument.Parse(Resources.DocumentWithSingleNestedListFilledAndRemovedCC_styles);
			var expectedNumbering = XDocument.Parse(Resources.DocumentWithSingleNestedListFilledAndRemovedCC_numbering);

			var valuesToFill = new Content(
				ListContent.Create("Document")
					.AddItem(ListItemContent.Create("Header", "Introduction"))

					.AddItem(ListItemContent.Create("Header", "Chapter 1 - The new start screen")
						.AddField("Header text", "Header 2 paragraph text")
						.AddNestedItem(ListItemContent.Create("Subheader", "What's new in Windows 8?")
							.AddField("Subheader text", "Subheader 2.1 paragraph text"))
						.AddNestedItem(ListItemContent.Create("Subheader", "Starting Windows 8")))

					.AddItem(ListItemContent.Create("Header", "Chapter 2 - The traditional Desktop")
						.AddNestedItem(ListItemContent.Create("Subheader", "Browsing the File Explorer"))
						.AddNestedItem(ListItemContent.Create("Subheader", "Getting the Lowdown on Folders and Libraries"))));


			var filledDocument = new TemplateProcessor(templateDocument, templateStyles, templateNumbering)
				.SetRemoveContentControls(true)
				.FillContent(valuesToFill);

			Assert.AreEqual(expectedDocument.ToString(), filledDocument.Document.ToString());
			Assert.AreEqual(expectedStyles.ToString(), filledDocument.StylesPart.ToString());
			Assert.AreEqual(expectedNumbering.ToString(), filledDocument.NumberingPart.ToString());
		}

		[TestMethod]
		public void FillingOneNestedListInsideTableAndRemoveContentControl()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithNestedListInsideTable_document);
			var templateStyles = XDocument.Parse(Resources.TemplateWithNestedListInsideTable_styles);
			var templateNumbering = XDocument.Parse(Resources.TemplateWithNestedListInsideTable_numbering);

			var expectedDocument = XDocument.Parse(Resources.DocumentWithNestedListInsideTableAndRemovedCC_document);
			var expectedStyles = XDocument.Parse(Resources.DocumentWithNestedListInsideTableAndRemovedCC_styles);
			var expectedNumbering = XDocument.Parse(Resources.DocumentWithNestedListInsideTableAndRemovedCC_numbering);

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


			var filledDocument = new TemplateProcessor(templateDocument, templateStyles, templateNumbering)
				.SetRemoveContentControls(true)
				.FillContent(valuesToFill);

			Assert.AreEqual(expectedDocument.ToString(), filledDocument.Document.ToString());
			Assert.AreEqual(expectedStyles.ToString(), filledDocument.StylesPart.ToString());
			Assert.AreEqual(RemoveNsid(expectedNumbering.ToString()), RemoveNsid(filledDocument.NumberingPart.ToString()));
		}

		[TestMethod]
		public void FillingOneNestedListInsideTableAndPreserveContentControl()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithNestedListInsideTable_document);
			var templateStyles = XDocument.Parse(Resources.TemplateWithNestedListInsideTable_styles);
			var templateNumbering = XDocument.Parse(Resources.TemplateWithNestedListInsideTable_numbering);

			var expectedDocument = XDocument.Parse(Resources.DocumentWithNestedListInsideTable_document);
			var expectedStyles = XDocument.Parse(Resources.DocumentWithNestedListInsideTable_styles);
			var expectedNumbering = XDocument.Parse(Resources.DocumentWithNestedListInsideTable_numbering);

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


			var filledDocument = new TemplateProcessor(templateDocument, templateStyles, templateNumbering)
				.SetRemoveContentControls(false)
				.FillContent(valuesToFill);

			Assert.AreEqual(expectedDocument.ToString(), filledDocument.Document.ToString());
			Assert.AreEqual(expectedStyles.ToString(), filledDocument.StylesPart.ToString());
			Assert.AreEqual(RemoveNsid(expectedNumbering.ToString()), RemoveNsid(filledDocument.NumberingPart.ToString()));
		}

		[TestMethod]
		public void FillingOneTableInsideListAndPreserveContentControl()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithTableInsideList_document);
			var templateStyles = XDocument.Parse(Resources.TemplateWithTableInsideList_styles);
			var templateNumbering = XDocument.Parse(Resources.TemplateWithTableInsideList_numbering);

			var expectedDocument = XDocument.Parse(Resources.DocumentWithTableInsideList_document);
			var expectedStyles = XDocument.Parse(Resources.DocumentWithTableInsideList_styles);
			var expectedNumbering = XDocument.Parse(Resources.DocumentWithTableInsideList_numbering);

			var valuesToFill = new Content(
				new ListContent("Products")
					.AddItem(new ListItemContent("Category", "Fruits")
						.AddTable(TableContent.Create("Items")
							.AddRow(new FieldContent("Name", "Orange"), new FieldContent("Count", "10"))
							.AddRow(new FieldContent("Name", "Apple"), new FieldContent("Count", "15"))))
					.AddItem(new ListItemContent("Category", "Vegetables")
						.AddTable(TableContent.Create("Items")
							.AddRow(new FieldContent("Name", "Tomato"), new FieldContent("Count", "8"))
							.AddRow(new FieldContent("Name", "Cabbage"), new FieldContent("Count", "17")))));


			var filledDocument = new TemplateProcessor(templateDocument, templateStyles, templateNumbering)
				.SetRemoveContentControls(false)
				.FillContent(valuesToFill);

			Assert.AreEqual(expectedDocument.ToString(), filledDocument.Document.ToString());
			Assert.AreEqual(expectedStyles.ToString(), filledDocument.StylesPart.ToString());
			Assert.AreEqual(RemoveNsid(expectedNumbering.ToString()), RemoveNsid(filledDocument.NumberingPart.ToString()));
		}
		[TestMethod]
		public void FillingOneTableInsideListAndRemovedContentControl()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithTableInsideList_document);
			var templateStyles = XDocument.Parse(Resources.TemplateWithTableInsideList_styles);
			var templateNumbering = XDocument.Parse(Resources.TemplateWithTableInsideList_numbering);

			var expectedDocument = XDocument.Parse(Resources.DocumentWithTableInsideListAndRemovedCC_document);
			var expectedStyles = XDocument.Parse(Resources.DocumentWithTableInsideListAndRemovedCC_styles);
			var expectedNumbering = XDocument.Parse(Resources.DocumentWithTableInsideListAndRemovedCC_numbering);

			var valuesToFill = new Content(
				new ListContent("Products")
					.AddItem(new ListItemContent("Category", "Fruits")
						.AddTable(TableContent.Create("Items")
							.AddRow(new FieldContent("Name", "Orange"), new FieldContent("Count", "10"))
							.AddRow(new FieldContent("Name", "Apple"), new FieldContent("Count", "15"))))
					.AddItem(new ListItemContent("Category", "Vegetables")
						.AddTable(TableContent.Create("Items")
							.AddRow(new FieldContent("Name", "Tomato"), new FieldContent("Count", "8"))
							.AddRow(new FieldContent("Name", "Cabbage"), new FieldContent("Count", "17")))));


			var filledDocument = new TemplateProcessor(templateDocument, templateStyles, templateNumbering)
				.SetRemoveContentControls(true)
				.FillContent(valuesToFill);

			Assert.AreEqual(expectedDocument.ToString(), filledDocument.Document.ToString());
			Assert.AreEqual(expectedStyles.ToString(), filledDocument.StylesPart.ToString());
			Assert.AreEqual(RemoveNsid(expectedNumbering.ToString()), RemoveNsid(filledDocument.NumberingPart.ToString()));
		}

		[TestMethod]
		public void FillingOneListAndFieldInsideNestedListAndPreserveContentControl()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithFieldAndListInsideNestedList_document);
			var templateStyles = XDocument.Parse(Resources.TemplateWithFieldAndListInsideNestedList_styles);
			var templateNumbering = XDocument.Parse(Resources.TemplateWithFieldAndListInsideNestedList_numbering);

			var expectedDocument = XDocument.Parse(Resources.DocumentWithFieldAndListInsideNestedListFilled_document);
			var expectedStyles = XDocument.Parse(Resources.DocumentWithFieldAndListInsideNestedListFilled_styles);
			var expectedNumbering = XDocument.Parse(Resources.DocumentWithFieldAndListInsideNestedListFilled_numbering);

			var valuesToFill = new Content(
				   new ListContent("Document")
					   .AddItem(new ListItemContent("Header", "First classification")
						   .AddNestedItem(new ListItemContent("Subheader", "Food classification")
							   .AddField("Paragraph", "Text about food classification")
							   .AddList(
								   ListContent.Create("Products")
									   .AddItem(new ListItemContent("Category", "Fruits")
										   .AddNestedItem(new ListItemContent("Name", "Apple"))
										   .AddNestedItem(new ListItemContent("Name", "Orange")))
									   .AddItem(new ListItemContent("Category", "Vegetables")
										   .AddNestedItem(new ListItemContent("Name", "Tomato"))
										   .AddNestedItem(new ListItemContent("Name", "Cabbage"))))))
					   .AddItem(new ListItemContent("Header", "Second classification")
						   .AddNestedItem(new ListItemContent("Subheader", "Animals classification")
							   .AddField("Paragraph", "Text about animal classification")
							   .AddList(
								   ListContent.Create("Products")
									   .AddItem(new ListItemContent("Category", "Vertebrate")
										   .AddNestedItem(new ListItemContent("Name", "Fish"))
										   .AddNestedItem(new ListItemContent("Name", "Mammal")))
									   .AddItem(new ListItemContent("Category", "Invertebrate")
										   .AddNestedItem(new ListItemContent("Name", "Crustacean"))
										   .AddNestedItem(new ListItemContent("Name", "Insect")))))));

			var filledDocument = new TemplateProcessor(templateDocument, templateStyles, templateNumbering)
				.SetRemoveContentControls(false)
				.FillContent(valuesToFill);

			Assert.AreEqual(expectedDocument.ToString(), filledDocument.Document.ToString());
			Assert.AreEqual(expectedStyles.ToString(), filledDocument.StylesPart.ToString());
			Assert.AreEqual(RemoveNsid(expectedNumbering.ToString()), RemoveNsid(filledDocument.NumberingPart.ToString()));
		}


		[TestMethod]
		public void FillingOneListAndFieldInsideNestedListAndRemoveContentControl()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithFieldAndListInsideNestedList_document);
			var templateStyles = XDocument.Parse(Resources.TemplateWithFieldAndListInsideNestedList_styles);
			var templateNumbering = XDocument.Parse(Resources.TemplateWithFieldAndListInsideNestedList_numbering);

			var expectedDocument = XDocument.Parse(Resources.DocumentWithFieldAndListInsideNestedListFilledAndRemovedCC_document);
			var expectedStyles = XDocument.Parse(Resources.DocumentWithFieldAndListInsideNestedListFilledAndRemovedCC_styles);
			var expectedNumbering = XDocument.Parse(Resources.DocumentWithFieldAndListInsideNestedListFilledAndRemovedCC_numbering);

			var valuesToFill = new Content(
				   new ListContent("Document")
					   .AddItem(new ListItemContent("Header", "First classification")
						   .AddNestedItem(new ListItemContent("Subheader", "Food classification")
							   .AddField("Paragraph", "Text about food classification")
							   .AddList(
								   ListContent.Create("Products")
									   .AddItem(new ListItemContent("Category", "Fruits")
										   .AddNestedItem(new ListItemContent("Name", "Apple"))
										   .AddNestedItem(new ListItemContent("Name", "Orange")))
									   .AddItem(new ListItemContent("Category", "Vegetables")
										   .AddNestedItem(new ListItemContent("Name", "Tomato"))
										   .AddNestedItem(new ListItemContent("Name", "Cabbage"))))))
					   .AddItem(new ListItemContent("Header", "Second classification")
						   .AddNestedItem(new ListItemContent("Subheader", "Animals classification")
							   .AddField("Paragraph", "Text about animal classification")
							   .AddList(
								   ListContent.Create("Products")
									   .AddItem(new ListItemContent("Category", "Vertebrate")
										   .AddNestedItem(new ListItemContent("Name", "Fish"))
										   .AddNestedItem(new ListItemContent("Name", "Mammal")))
									   .AddItem(new ListItemContent("Category", "Invertebrate")
										   .AddNestedItem(new ListItemContent("Name", "Crustacean"))
										   .AddNestedItem(new ListItemContent("Name", "Insect")))))));

			var filledDocument = new TemplateProcessor(templateDocument, templateStyles, templateNumbering)
				.SetRemoveContentControls(true)
				.FillContent(valuesToFill);

			Assert.AreEqual(expectedDocument.ToString(), filledDocument.Document.ToString());
			Assert.AreEqual(expectedStyles.ToString(), filledDocument.StylesPart.ToString());
			Assert.AreEqual(RemoveNsid(expectedNumbering.ToString()), RemoveNsid(filledDocument.NumberingPart.ToString()));
		}

		[TestMethod]
		public void FillingTwoTablesWithListsInsideAndPreverseContentControl()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithTwoTablesWithListsInside_document);
			var templateStyles = XDocument.Parse(Resources.TemplateWithTwoTablesWithListsInside_styles);
			var templateNumbering = XDocument.Parse(Resources.TemplateWithTwoTablesWithListsInside_numbering);

			var expectedDocument = XDocument.Parse(Resources.DocumentWithTwoTablesWithListsInsideFilled_document);
			var expectedStyles = XDocument.Parse(Resources.DocumentWithTwoTablesWithListsInsideFilled_styles);
			var expectedNumbering = XDocument.Parse(Resources.DocumentWithTwoTablesWithListsInsideFilled_numbering);

			var valuesToFill = new Content(
			  new TableContent("Peoples")
				  .AddRow(
					  new FieldContent("Name", "Eric"),
					  new FieldContent("Age", "34"),
					  new ListContent("Childs")
						  .AddItem(new ListItemContent("ChildName", "Robbie"))
						  .AddItem(new ListItemContent("ChildName", "Trisha")))
				  .AddRow(
					  new FieldContent("Name", "Poll"),
					  new FieldContent("Age", "40"),
					  new ListContent("Childs")
						  .AddItem(new ListItemContent("ChildName", "Ann"))
						  .AddItem(new ListItemContent("ChildName", "Richard"))),
			  new TableContent("Team Members")
				  .AddRow(
					  new FieldContent("Name", "Eric"),
					  new ListContent("Roles")
						  .AddItem(new ListItemContent("Role", "Developer"))
						  .AddItem(new ListItemContent("Role", "Tester")))
				  .AddRow(
					  new FieldContent("Name", "Poll"),
					  new ListContent("Roles")
						  .AddItem(new ListItemContent("Role", "Admin"))
						  .AddItem(new ListItemContent("Role", "Developer"))));

			var filledDocument = new TemplateProcessor(templateDocument, templateStyles, templateNumbering)
				.SetRemoveContentControls(false)
				.FillContent(valuesToFill);

			Assert.AreEqual(expectedDocument.ToString(), filledDocument.Document.ToString());
			Assert.AreEqual(expectedStyles.ToString(), filledDocument.StylesPart.ToString());
			Assert.AreEqual(RemoveNsid(expectedNumbering.ToString()), RemoveNsid(filledDocument.NumberingPart.ToString()));
		}

		[TestMethod]
		public void FillingTwoTablesWithListsInsideAndRemoveContentControl()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithTwoTablesWithListsInside_document);
			var templateStyles = XDocument.Parse(Resources.TemplateWithTwoTablesWithListsInside_styles);
			var templateNumbering = XDocument.Parse(Resources.TemplateWithTwoTablesWithListsInside_numbering);

			var expectedDocument = XDocument.Parse(Resources.DocumentWithTwoTablesWithListsInsideFilledAndRemovedCC_document);
			var expectedStyles = XDocument.Parse(Resources.DocumentWithTwoTablesWithListsInsideFilledAndRemovedCC_styles);
			var expectedNumbering = XDocument.Parse(Resources.DocumentWithTwoTablesWithListsInsideFilledAndRemovedCC_numbering);

			var valuesToFill = new Content(
			  new TableContent("Peoples")
				  .AddRow(
					  new FieldContent("Name", "Eric"),
					  new FieldContent("Age", "34"),
					  new ListContent("Childs")
						  .AddItem(new ListItemContent("ChildName", "Robbie"))
						  .AddItem(new ListItemContent("ChildName", "Trisha")))
				  .AddRow(
					  new FieldContent("Name", "Poll"),
					  new FieldContent("Age", "40"),
					  new ListContent("Childs")
						  .AddItem(new ListItemContent("ChildName", "Ann"))
						  .AddItem(new ListItemContent("ChildName", "Richard"))),
			  new TableContent("Team Members")
				  .AddRow(
					  new FieldContent("Name", "Eric"),
					  new ListContent("Roles")
						  .AddItem(new ListItemContent("Role", "Developer"))
						  .AddItem(new ListItemContent("Role", "Tester")))
				  .AddRow(
					  new FieldContent("Name", "Poll"),
					  new ListContent("Roles")
						  .AddItem(new ListItemContent("Role", "Admin"))
						  .AddItem(new ListItemContent("Role", "Developer"))));

			var filledDocument = new TemplateProcessor(templateDocument, templateStyles, templateNumbering)
				.SetRemoveContentControls(true)
				.FillContent(valuesToFill);

			Assert.AreEqual(expectedDocument.ToString(), filledDocument.Document.ToString());
			Assert.AreEqual(expectedStyles.ToString(), filledDocument.StylesPart.ToString());
			Assert.AreEqual(RemoveNsid(expectedNumbering.ToString()), RemoveNsid(filledDocument.NumberingPart.ToString()));
		}

		[TestMethod]
		public void FillingTableWithTwoBlocksAndRemoveContentControl()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithSingleTableWithTwoBlocks);

			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleTableWithTwoBlocksFilledAndRemovedCC);

			var valuesToFill = new Content(
						new TableContent("Team Members")
								.AddRow(
									new FieldContent("Name", "Eric"),
									new FieldContent("Role", "Program Manager"))
								.AddRow(
									new FieldContent("Name", "Bob"),
									new FieldContent("Role", "Developer")),

								new TableContent("Team Members")
								.AddRow(
									new FieldContent("Statistics Role", "Program Manager"),
									new FieldContent("Statistics Role Count", "1"))
								.AddRow(
									new FieldContent("Statistics Role", "Developer"),
									new FieldContent("Statistics Role Count", "1")));

			var filledDocument = new TemplateProcessor(templateDocument)
				.SetRemoveContentControls(true)
				.FillContent(valuesToFill);

			Assert.AreEqual(expectedDocument.ToString(), filledDocument.Document.ToString());
		}

		[TestMethod]
		public void FillingTableWithTwoBlocksAndPreverseContentControl()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithSingleTableWithTwoBlocks);

			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleTableWithTwoBlocksFilled);

			var valuesToFill = new Content(
						new TableContent("Team Members")
								.AddRow(
									new FieldContent("Name", "Eric"),
									new FieldContent("Role", "Program Manager"))
								.AddRow(
									new FieldContent("Name", "Bob"),
									new FieldContent("Role", "Developer")),

								new TableContent("Team Members")
								.AddRow(
									new FieldContent("Statistics Role", "Program Manager"),
									new FieldContent("Statistics Role Count", "1"))
								.AddRow(
									new FieldContent("Statistics Role", "Developer"),
									new FieldContent("Statistics Role Count", "1")));

			var filledDocument = new TemplateProcessor(templateDocument)
				.SetRemoveContentControls(false)
				.FillContent(valuesToFill);

			Assert.AreEqual(expectedDocument.ToString(), filledDocument.Document.ToString());
		}

	    private string RemoveNsid(string source)
	    {
			const string nsidRegexp = "nsid w:val=\"[0-9a-fA-F]+\"";
			return Regex.Replace(source, nsidRegexp, "");
	    }
    }
}
