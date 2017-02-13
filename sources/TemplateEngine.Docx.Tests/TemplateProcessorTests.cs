using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
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

            Assert.IsNotNull(expectedDocument.Document);
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

			Assert.IsNotNull(expectedDocument.Document);
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

			Assert.IsNotNull(expectedDocument.Document);
            Assert.AreEqual(expectedDocument.Document.ToString(), documentXml);
        }

		[TestMethod]
		public void FillingOneFieldWithValue_ValueContainsLineBreak_ShouldInsertLineBreakToResultDocument()
		{
			var templateDocument = XDocument.Parse(Resources.TemplateWithSingleField);
			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleFieldFilledWithLinebreaks);

			var dateTime = new DateTime(2013, 09, 06);
			var valuesToFill = new Content
			{
				Fields = new List<FieldContent>
                {
                    new FieldContent { Name = "ReportDate",
						Value = string.Format("{0}\r\n{1}\n{2}",
						dateTime.ToString("d", CultureInfo.InvariantCulture), 
						dateTime.ToString("D", CultureInfo.InvariantCulture),
						dateTime.ToString("y", CultureInfo.InvariantCulture)) 
					}
                }
			};

			var template = new TemplateProcessor(templateDocument)
				.SetRemoveContentControls(true)
				.FillContent(valuesToFill);

			var documentXml = template.Document.ToString();

			Assert.IsNotNull(expectedDocument.Document);
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

			Assert.AreEqual(expectedDocument.ToString(), documentXml);
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

            Assert.AreEqual(expectedDocument.ToString(), documentXml);
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

			Assert.AreEqual(expectedDocument.ToString(), documentXml);
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

			Assert.AreEqual(expectedDocument.ToString(), documentXml);
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

			Assert.AreEqual(expectedDocument.ToString(), documentXml);
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

			Assert.AreEqual(expectedDocument.ToString(), documentXml);
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

			Assert.AreEqual(expectedDocument.ToString(), documentXml);
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

			Assert.AreEqual(expectedDocument.ToString(), documentXml);
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
						  .AddItem(new FieldContent("ChildName", "Ann"))
						  .AddItem(new FieldContent("ChildName", "Richard"))),
			  new TableContent("Team Members")
				  .AddRow(
					  new FieldContent("Name", "Eric"),
					  new ListContent("Roles")
						  .AddItem(new ListItemContent("Role", "Developer"))
						  .AddItem(new ListItemContent("Role", "Tester")))
				  .AddRow(
					  new FieldContent("Name", "Poll"),
					  new ListContent("Roles")
						  .AddItem(new FieldContent("Role", "Admin"))
						  .AddItem(new FieldContent("Role", "Developer"))));

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
						  .AddItem(new FieldContent("ChildName", "Robbie"))
						  .AddItem(new FieldContent("ChildName", "Trisha")))
				  .AddRow(
					  new FieldContent("Name", "Poll"),
					  new FieldContent("Age", "40"),
					  new ListContent("Childs")
						  .AddItem(new FieldContent("ChildName", "Ann"))
						  .AddItem(new FieldContent("ChildName", "Richard"))),
			  new TableContent("Team Members")
				  .AddRow(
					  new FieldContent("Name", "Eric"),
					  new ListContent("Roles")
						  .AddItem(new FieldContent("Role", "Developer"))
						  .AddItem(new FieldContent("Role", "Tester")))
				  .AddRow(
					  new FieldContent("Name", "Poll"),
					  new ListContent("Roles")
						  .AddItem(new FieldContent("Role", "Admin"))
						  .AddItem(new FieldContent("Role", "Developer"))));

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

		[TestMethod]
		public void FillingSingleImageAndRemoveContentControl()
		{
			var templateDocumentDocx = Resources.TemplateWithSingleImage;
			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleImage_AndRemovedCC); 

			var newFile = File.ReadAllBytes("Tesla.jpg");

			var valuesToFill = new Content(
						new ImageContent("TeslaPhoto", newFile));

			TemplateProcessor processor;
			byte[] resultImage;
			using (var ms = new MemoryStream())
			{
				ms.Write(templateDocumentDocx, 0, templateDocumentDocx.Length);

				processor = new TemplateProcessor(ms)
					.SetRemoveContentControls(true)
					.FillContent(valuesToFill);

				resultImage = GetImageFromPart(processor.ImagesPart, 0);
			}

			Assert.AreEqual(processor.ImagesPart.Count(), 1);
			Assert.IsNotNull(resultImage);
			Assert.IsTrue(resultImage.SequenceEqual(newFile));
			
			Assert.AreEqual(RemoveRembed(expectedDocument.ToString().Trim()), 
				RemoveRembed(processor.Document.ToString().Trim()));
		}

		[TestMethod]
		public void FillingSingleImageAndPreverseContentControl()
		{
			var templateDocumentDocx = Resources.TemplateWithSingleImage;
			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleImage);

			var newFile = File.ReadAllBytes("Tesla.jpg");

			var valuesToFill = new Content(
						new ImageContent("TeslaPhoto", newFile));

			TemplateProcessor processor;
			byte[] resultImage;
			using (var ms = new MemoryStream())
			{
				ms.Write(templateDocumentDocx, 0, templateDocumentDocx.Length);
				processor = new TemplateProcessor(ms)
					.SetRemoveContentControls(false)
					.FillContent(valuesToFill);

				resultImage = GetImageFromPart(processor.ImagesPart, 0);
			}

			Assert.AreEqual(processor.ImagesPart.Count(), 1);
			Assert.IsNotNull(resultImage);
			Assert.IsTrue(resultImage.SequenceEqual(newFile));

			Assert.AreEqual(RemoveRembed(expectedDocument.ToString().Trim()), 
				RemoveRembed(processor.Document.ToString().Trim()));
		}

		[TestMethod]
		public void FillingSingleImage_ImageContentControlNotFound_ShowError()
		{
			var templateDocumentDocx = Resources.TemplateEmpty;
			var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleImageNotFoundError);

			var newFile = File.ReadAllBytes("Tesla.jpg");

			var valuesToFill = new Content(
						new ImageContent("TeslaPhoto", newFile));

			TemplateProcessor processor;

			using (var ms = new MemoryStream(templateDocumentDocx))
			{
				processor = new TemplateProcessor(ms)
					.SetRemoveContentControls(false)
					.FillContent(valuesToFill);
			}

			Assert.AreEqual(processor.ImagesPart.Count(), 0);
		
			Assert.AreEqual(expectedDocument.ToString().Trim(), processor.Document.ToString().Trim());
		}

		[TestMethod]
		public void FillingImageInsideTable_CorrectFiledItems_Success()
		{
			var templateDocumentDocx = Resources.TemplateWithImagesInsideTable;
			var expectedDocument = XDocument.Parse(Resources.DocumentWithImagesInsideTable);

			var valuesToFill = new Content(
				new TableContent("Scientists")
					.AddRow(new FieldContent("Name", "Nicola Tesla"),
						new FieldContent("Born", new DateTime(1856, 7, 10).ToShortDateString()),
						new ImageContent("Photo", File.ReadAllBytes("Tesla.jpg")),
						new FieldContent("Info",
							"Serbian American inventor, electrical engineer, mechanical engineer, physicist, and futurist best known for his contributions to the design of the modern alternating current (AC) electricity supply system"))
					.AddRow(new FieldContent("Name", "Thomas Edison"),
						new FieldContent("Born", new DateTime(1847, 2, 11).ToShortDateString()),
						new ImageContent("Photo", File.ReadAllBytes("Edison.jpg")),
						new FieldContent("Info",
							"American inventor and businessman. He developed many devices that greatly influenced life around the world, including the phonograph, the motion picture camera, and the long-lasting, practical electric light bulb."))
					.AddRow(new FieldContent("Name", "Albert Einstein"),
						new FieldContent("Born", new DateTime(1879, 3, 14).ToShortDateString()),
						new ImageContent("Photo", File.ReadAllBytes("Einstein.jpg")),
						new FieldContent("Info",
							"German-born theoretical physicist. He developed the general theory of relativity, one of the two pillars of modern physics (alongside quantum mechanics). Einstein's work is also known for its influence on the philosophy of science. Einstein is best known in popular culture for his mass–energy equivalence formula E = mc2 (which has been dubbed 'the world's most famous equation')."))
				);


			TemplateProcessor processor;

			using (var ms = new MemoryStream())
			{
				ms.Write(templateDocumentDocx, 0, templateDocumentDocx.Length);
				processor = new TemplateProcessor(ms)
					.SetRemoveContentControls(true)
					.FillContent(valuesToFill);
			}

			Assert.AreEqual(3, processor.ImagesPart.Count());

			Assert.AreEqual(RemoveRembed(expectedDocument.ToString().Trim()), 
				RemoveRembed(processor.Document.ToString().Trim()));
		}

		[TestMethod]
		public void FillingImageInsideAList_CorrectFiledItems_Success()
		{
			var templateDocumentDocx = Resources.TemplateWithImagesInsideList;
			var expectedDocument = XDocument.Parse(Resources.DocumentWithImagesInsideListFilledAndRemovedCC);

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

			TemplateProcessor processor;

			using (var ms = new MemoryStream())
			{
				ms.Write(templateDocumentDocx, 0, templateDocumentDocx.Length);
				processor = new TemplateProcessor(ms)
					.SetRemoveContentControls(true)
					.FillContent(valuesToFill);
			}

			Assert.AreEqual(3, processor.ImagesPart.Count());

			Assert.AreEqual(RemoveRembed(expectedDocument.ToString().Trim()), 
				RemoveRembed(processor.Document.ToString().Trim()));
		}

		[TestMethod]
		public void FillingFieldsInHeaderAndFooter_WithCorrectValues_Success()
		{
			var templateDocumentDocx = Resources.TemplateEmptyWithFieldsInHeaderAndFooter;
			var expectedHeader = XDocument.Parse(Resources.DocumentWithFieldFilledInHeaderAndFooter_header);
			var expectedFooter = XDocument.Parse(Resources.DocumentWithFieldFilledInHeaderAndFooter_footer);


			var valuesToFill = new Content(
				new FieldContent("Company name", "Spiderwasp Communications"),
				new FieldContent("Copyright", "© All rights reserved"));

			TemplateProcessor processor;

			using (var ms = new MemoryStream(templateDocumentDocx))
			{
				processor = new TemplateProcessor(ms)
					.SetRemoveContentControls(false)
					.FillContent(valuesToFill);
			}

			Assert.AreEqual(processor.HeaderParts.Count, 1);
			Assert.AreEqual(processor.FooterParts.Count, 1);

			Assert.AreEqual(expectedHeader.ToString().Trim(), processor.HeaderParts.First().Value.ToString().Trim());
			Assert.AreEqual(expectedFooter.ToString().Trim(), processor.FooterParts.First().Value.ToString().Trim());
		}
		[TestMethod]
		public void FillingFieldsInHeaderAndFooter_WithCorrectValuesAndRemoveContentControls_Success()
		{
			var templateDocumentDocx = Resources.TemplateEmptyWithFieldsInHeaderAndFooter;
			var expectedHeader = XDocument.Parse(Resources.DocumentWithFieldFilledInHeaderAndFooterAndRemovedCC_header);
			var expectedFooter = XDocument.Parse(Resources.DocumentWithFieldFilledInHeaderAndFooterAndRemovedCC_footer);


			var valuesToFill = new Content(
				new FieldContent("Company name", "Spiderwasp Communications"),
				new FieldContent("Copyright", "© All rights reserved"));

			TemplateProcessor processor;

			using (var ms = new MemoryStream(templateDocumentDocx))
			{
				processor = new TemplateProcessor(ms)
					.SetRemoveContentControls(true)
					.FillContent(valuesToFill);
			}

			Assert.AreEqual(processor.HeaderParts.Count, 1);
			Assert.AreEqual(processor.FooterParts.Count, 1);

			Assert.AreEqual(expectedHeader.ToString().Trim(), processor.HeaderParts.First().Value.ToString().Trim());
			Assert.AreEqual(expectedFooter.ToString().Trim(), processor.FooterParts.First().Value.ToString().Trim());
		}

		[TestMethod]
		public void FillingTwoLists_InMainDocumentAndInFooter_Success()
		{
			var templateDocumentDocx = Resources.TemplateWithTwoListsInMainDocumentAndInFooter;
			var expectedDocument = XDocument.Parse(Resources.DocumentWithTwoListsInMainDocumentAndInFooter_document);
			var expectedFooter = XDocument.Parse(Resources.DocumentWithTwoListsInMainDocumentAndInFooter_footer);


			var valuesToFill = new Content(new ListContent
			{
				Name = "Footer",
				Items = new List<ListItemContent>
                        {                   
                             new ListItemContent ("Footer item", "Spiderwasp Communications"),
                             new ListItemContent ("Footer item", "© All rights reserved")
                        }
			},
					ListContent.Create("Document",
					ListItemContent.Create("Header", "Introduction"),

					ListItemContent.Create("Header", "Chapter 1 - The new start screen")
						.AddNestedItem(ListItemContent.Create("Subheader", "What's new in Windows 8?"))
						.AddNestedItem(ListItemContent.Create("Subheader", "Starting Windows 8")),

					ListItemContent.Create("Header", "Chapter 2 - The traditional Desktop")
						.AddNestedItem(ListItemContent.Create("Subheader", "Browsing the File Explorer"))
						.AddNestedItem(ListItemContent.Create("Subheader", "Getting the Lowdown on Folders and Libraries")))
				);

			TemplateProcessor processor;

			using (var ms = new MemoryStream(templateDocumentDocx))
			{
				processor = new TemplateProcessor(ms)
					.SetRemoveContentControls(false)
					.FillContent(valuesToFill);
			}

			Assert.AreEqual(processor.FooterParts.Count, 1);

			Assert.AreEqual(expectedDocument.ToString().Trim(), processor.Document.ToString().Trim());
			Assert.AreEqual(expectedFooter.ToString().Trim(), processor.FooterParts.First().Value.ToString().Trim());
		}

        [TestMethod]
        public void FillingRepeatContent_WithCorrectValues_Success()
        {
            var templateDocumentDocx = Resources.TemplateWithRepeatContentWithImagesAndFields;
            var expectedDocument = XDocument.Parse(Resources.DocumentWithRepeatContentWithImagesAndFields_document);

            var valuesToFill = new Content(new RepeatContent("Repeats")
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
                        "German-born theoretical physicist. He developed the general theory of relativity, one of the two pillars of modern physics (alongside quantum mechanics). Einstein's work is also known for its influence on the philosophy of science. Einstein is best known in popular culture for his mass–energy equivalence formula E = mc2 (which has been dubbed 'the world's most famous equation').")));


            TemplateProcessor processor;

            using (var ms = new MemoryStream())
            {
                ms.Write(templateDocumentDocx, 0, templateDocumentDocx.Length);
                processor = new TemplateProcessor(ms)
                    .FillContent(valuesToFill);
            }

            Assert.AreEqual(RemoveRembed(expectedDocument.ToString().Trim()), RemoveRembed(processor.Document.ToString().Trim()));
        }

        [TestMethod]
        public void FillingRepeatContent_WithCorrectValuesAndRemoveContentControls_Success()
        {
            var templateDocumentDocx = Resources.TemplateWithRepeatContentWithImagesAndFields;
            var expectedDocument = XDocument.Parse(Resources.DocumentWithRepeatContentWithImagesAndFieldsAndRemovedCC_document);
            
            var valuesToFill = new Content(new RepeatContent("Repeats")
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
                        "German-born theoretical physicist. He developed the general theory of relativity, one of the two pillars of modern physics (alongside quantum mechanics). Einstein's work is also known for its influence on the philosophy of science. Einstein is best known in popular culture for his mass–energy equivalence formula E = mc2 (which has been dubbed 'the world's most famous equation').")));
            

            TemplateProcessor processor;

            using (var ms = new MemoryStream())
            {
                ms.Write(templateDocumentDocx, 0, templateDocumentDocx.Length);
                processor = new TemplateProcessor(ms)
                    .SetRemoveContentControls(true)
                    .FillContent(valuesToFill);
            }

            Assert.AreEqual(RemoveRembed(expectedDocument.ToString().Trim()), RemoveRembed(processor.Document.ToString().Trim()));
        }
        
        [TestMethod]
        public void FillingRepeatContent_WithImageTableAndList_Success()
        {
            var templateDocumentDocx = Resources.TemplateWithRepeatContentWithImageTableAndList;
           
            var expectedDocument = XDocument.Parse(Resources.DocumentWithRepeatContentWithImageTableAndList_document);

            var valuesToFill = new Content(new RepeatContent("Scientists")
                .AddItem(new FieldContent("Name", "Nicola Tesla"),
                    new ImageContent("Photo", File.ReadAllBytes("Tesla.jpg")),
                    new FieldContent("Info",
                        "Serbian American inventor, electrical engineer, mechanical engineer, physicist, and futurist best known for his contributions to the design of the modern alternating current (AC) electricity supply system"),
                    new TableContent("Awards")
                        .AddRow(new FieldContent("Award Name", "Order of St. Sava, II Class"), new FieldContent("Award Date", "1892"))
                        .AddRow(new FieldContent("Award Name", "Elliott Cresson Medal"), new FieldContent("Award Date", "1894"))
                        .AddRow(new FieldContent("Award Name", "Order of Prince Danilo I"), new FieldContent("Award Date", "1895"))
                        .AddRow(new FieldContent("Award Name", "..."), new FieldContent("Award Date", "")),
                    new ListContent("Inventions")
                        .AddItem(new ListItemContent("Invention", "Rotating Magnetic Field"))
                        .AddItem(new ListItemContent("Invention", "AC Motor"))
                        .AddItem(new ListItemContent("Invention", "Tesla coil"))
                        .AddItem(new ListItemContent("Invention", "...")))
                 .AddItem(new FieldContent("Name", "Thomas Edison"),
                    new ImageContent("Photo", File.ReadAllBytes("Edison.jpg")),
                    new FieldContent("Info",
                       "American inventor and businessman. He developed many devices that greatly influenced life around the world, including the phonograph, the motion picture camera, and the long-lasting, practical electric light bulb."),
                    new TableContent("Awards")
                        .AddRow(new FieldContent("Award Name", "Matteucci Medal"), new FieldContent("Award Date", "1887"))
                        .AddRow(new FieldContent("Award Name", "Edward Longstreth Medal"), new FieldContent("Award Date", "1899"))
                        .AddRow(new FieldContent("Award Name", "..."), new FieldContent("Award Date", "")),
                    new ListContent("Inventions")
                        .AddItem(new ListItemContent("Invention", "The Phonograph"))
                        .AddItem(new ListItemContent("Invention", "The Carbon Microphone"))
                        .AddItem(new ListItemContent("Invention", "The Incandescent Light Bulb"))
                        .AddItem(new ListItemContent("Invention", "...")))
                 .AddItem(new FieldContent("Name", "Albert Einstein"),
                    new ImageContent("Photo", File.ReadAllBytes("Einstein.jpg")),
                    new FieldContent("Info",
                       "German-born theoretical physicist. He developed the general theory of relativity, one of the two pillars of modern physics (alongside quantum mechanics). Einstein's work is also known for its influence on the philosophy of science. Einstein is best known in popular culture for his mass–energy equivalence formula E = mc2 (which has been dubbed 'the world's most famous equation')."),
                    new TableContent("Awards")
                        .AddRow(new FieldContent("Award Name", "Copley Medal"), new FieldContent("Award Date", "1925"))
                        .AddRow(new FieldContent("Award Name", "Franklin Medal"), new FieldContent("Award Date", "1936"))
                        .AddRow(new FieldContent("Award Name", "..."), new FieldContent("Award Date", "")),
                    new ListContent("Inventions")
                        .AddItem(new ListItemContent("Invention", "Brownian movement"))
                        .AddItem(new ListItemContent("Invention", "The quantum theory of light"))
                        .AddItem(new ListItemContent("Invention", "The special theory of relativity"))
                        .AddItem(new ListItemContent("Invention", "The link between mass and energy"))
                        .AddItem(new ListItemContent("Invention", "..."))));


            TemplateProcessor processor;

            using (var ms = new MemoryStream())
            {
                ms.Write(templateDocumentDocx, 0, templateDocumentDocx.Length);
                processor = new TemplateProcessor(ms)
                    .SetRemoveContentControls(true)
                    .FillContent(valuesToFill);
            }


            Assert.AreEqual(3, processor.ImagesPart.Count());

            Assert.AreEqual(RemoveRembed(expectedDocument.ToString().Trim()),
                RemoveRembed(processor.Document.ToString().Trim()));
        }

        [TestMethod]
        public void FillingSingleImageInHeader_Success()
        {
            var templateDocumentDocx = Resources.TemplateWithSingleImageInHeader;
            var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleImageInHeaderAndRemovedCC_document);
            var expectedHeader = XDocument.Parse(Resources.DocumentWithSingleImageInHeaderAndRemovedCC_header);

            var newFile = File.ReadAllBytes("Logo.jpg");

            var valuesToFill = new Content(
                        new ImageContent("Logo", newFile));

            TemplateProcessor processor;
            byte[] resultImage;
            using (var ms = new MemoryStream())
            {
                ms.Write(templateDocumentDocx, 0, templateDocumentDocx.Length);

                processor = new TemplateProcessor(ms)
                    .SetRemoveContentControls(true)
                    .FillContent(valuesToFill);

                resultImage = GetImageFromPart(processor.HeaderImagesParts.First().Value, 0);
            }

            Assert.AreEqual(1, processor.HeaderParts.Count);
            Assert.AreEqual(1, processor.HeaderImagesParts.Count);
            Assert.AreEqual(processor.HeaderParts.First().Key, processor.HeaderImagesParts.First().Key);
            Assert.IsNotNull(resultImage);
            Assert.IsTrue(resultImage.SequenceEqual(newFile));

            Assert.AreEqual(expectedDocument.ToString().Trim(),
                processor.Document.ToString().Trim());

            Assert.AreEqual(RemoveRembed(expectedHeader.ToString().Trim()),
               RemoveRembed(processor.HeaderParts.First().Value.ToString().Trim()));
        }

        [TestMethod]
        public void FillingSingleImageInHeader_PreverseContentControl_Success()
        {
            var templateDocumentDocx = Resources.TemplateWithSingleImageInHeader;
            var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleImageInHeader_document);
            var expectedHeader = XDocument.Parse(Resources.DocumentWithSingleImageInHeader_header);

            var newFile = File.ReadAllBytes("Logo.jpg");

            var valuesToFill = new Content(
                        new ImageContent("Logo", newFile));

            TemplateProcessor processor;
            byte[] resultImage;
            using (var ms = new MemoryStream())
            {
                ms.Write(templateDocumentDocx, 0, templateDocumentDocx.Length);

                processor = new TemplateProcessor(ms)
                    .SetRemoveContentControls(false)
                    .FillContent(valuesToFill);

                resultImage = GetImageFromPart(processor.HeaderImagesParts.First().Value, 0);
            }

            Assert.AreEqual(1, processor.HeaderParts.Count);
            Assert.AreEqual(1, processor.HeaderImagesParts.Count);
            Assert.AreEqual(processor.HeaderParts.First().Key, processor.HeaderImagesParts.First().Key);
            Assert.IsNotNull(resultImage);
            Assert.IsTrue(resultImage.SequenceEqual(newFile));

            Assert.AreEqual(RemoveRembed(expectedDocument.ToString().Trim()),
                RemoveRembed(processor.Document.ToString().Trim()));

            Assert.AreEqual(RemoveRembed(expectedHeader.ToString().Trim()),
               RemoveRembed(processor.HeaderParts.First().Value.ToString().Trim()));
        }


        [TestMethod]
        public void FillingSingleImageInFooter_Success()
        {
            var templateDocumentDocx = Resources.TemplateWithSingleImageInFooter;
            var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleImageInFooterAndRemovedCC_document);
            var expectedFooter = XDocument.Parse(Resources.DocumentWithSingleImageInFooterAndRemovedCC_footer);

            var newFile = File.ReadAllBytes("Logo.jpg");

            var valuesToFill = new Content(
                        new ImageContent("Logo", newFile));

            TemplateProcessor processor;
            byte[] resultImage;
            using (var ms = new MemoryStream())
            {
                ms.Write(templateDocumentDocx, 0, templateDocumentDocx.Length);

                processor = new TemplateProcessor(ms)
                    .SetRemoveContentControls(true)
                    .FillContent(valuesToFill);

                resultImage = GetImageFromPart(processor.FooterImagesParts.First().Value, 0);
            }

            Assert.AreEqual(1, processor.FooterParts.Count);
            Assert.AreEqual(1, processor.FooterImagesParts.Count);
            Assert.AreEqual(processor.FooterParts.First().Key, processor.FooterImagesParts.First().Key);
            Assert.IsNotNull(resultImage);
            Assert.IsTrue(resultImage.SequenceEqual(newFile));

            Assert.AreEqual(RemoveRembed(expectedDocument.ToString().Trim()),
                RemoveRembed(processor.Document.ToString().Trim()));

            Assert.AreEqual(RemoveRembed(expectedFooter.ToString().Trim()),
              RemoveRembed(processor.FooterParts.First().Value.ToString().Trim()));
        }


        [TestMethod]
        public void FillingImagesInHeaderBodyAndFooter_Success()
        {
            var templateDocumentDocx = Resources.TemplateWithImagesInHeaderBodyAndFooter;
            var expectedDocument = XDocument.Parse(Resources.DocumentWithImagesInHeaderBodyAndFooterAndRemovedCC_document);
            var expectedHeader = XDocument.Parse(Resources.DocumentWithImagesInHeaderBodyAndFooterAndRemovedCC_header);
            var expectedFooter = XDocument.Parse(Resources.DocumentWithImagesInHeaderBodyAndFooterAndRemovedCC_footer);

            var headerFile = File.ReadAllBytes("Edison.jpg");
            var bodyFile = File.ReadAllBytes("Einstein.jpg");
            var footerFile = File.ReadAllBytes("Tesla.jpg");

            var valuesToFill = new Content(
                   new ImageContent("HeaderImage", headerFile),
                   new ImageContent("BodyImage", bodyFile),
                   new ImageContent("FooterImage", footerFile)
               );

            TemplateProcessor processor;
            byte[] resultHeaderImage;
            byte[] resultBodyImage;
            byte[] resultFooterImage;

            using (var ms = new MemoryStream())
            {
                ms.Write(templateDocumentDocx, 0, templateDocumentDocx.Length);

                processor = new TemplateProcessor(ms)
                    .SetRemoveContentControls(true)
                    .FillContent(valuesToFill);

                resultHeaderImage = GetImageFromPart(processor.HeaderImagesParts.First().Value, 0);
                resultBodyImage = GetImageFromPart(processor.ImagesPart, 0);
                resultFooterImage = GetImageFromPart(processor.FooterImagesParts.First().Value, 0);
            }

            Assert.AreEqual(1, processor.HeaderParts.Count);
            Assert.AreEqual(1, processor.HeaderImagesParts.Count);
            Assert.AreEqual(processor.HeaderParts.First().Key, processor.HeaderImagesParts.First().Key);
            Assert.IsNotNull(resultHeaderImage);
            Assert.IsTrue(resultHeaderImage.SequenceEqual(headerFile));

            Assert.AreEqual(1, processor.FooterParts.Count);
            Assert.AreEqual(1, processor.FooterImagesParts.Count);
            Assert.AreEqual(processor.FooterParts.First().Key, processor.FooterImagesParts.First().Key);
            Assert.IsNotNull(resultFooterImage);
            Assert.IsTrue(resultFooterImage.SequenceEqual(footerFile));

            Assert.AreEqual(1, processor.ImagesPart.Count());
            Assert.IsNotNull(resultBodyImage);
            Assert.IsTrue(resultBodyImage.SequenceEqual(bodyFile));

            Assert.AreEqual(RemoveRembed(expectedDocument.ToString().Trim()),
                RemoveRembed(processor.Document.ToString().Trim()));

            Assert.AreEqual(RemoveRembed(expectedFooter.ToString().Trim()),
             RemoveRembed(processor.FooterParts.First().Value.ToString().Trim()));

            Assert.AreEqual(RemoveRembed(expectedHeader.ToString().Trim()),
             RemoveRembed(processor.HeaderParts.First().Value.ToString().Trim()));
        }


	    private static byte[] GetImageFromPart(IEnumerable<ImagePart> imageParts, int partIndex)
	    {
		    byte[] resultImage = null;
            if (imageParts.Any())
		    {
                var stream = imageParts.ToArray()[partIndex].GetStream();

			    resultImage = new byte[stream.Length];
                using (var reader = new BinaryReader(stream))
			    {
				    reader.Read(resultImage, 0, (int) stream.Length);
			    }
		    }
		    return resultImage;
	    }

	    private string RemoveNsid(string source)
	    {
			const string nsidRegexp = "nsid w:val=\"[0-9a-fA-F]+\"";
			return Regex.Replace(source, nsidRegexp, "");
	    }
	    private string RemoveRembed(string source)
	    {
			const string rembedRegexp = "r:embed=\"[0-9a-zA-Z]+\"";
			return Regex.Replace(source, rembedRegexp, "");
	    }
    }
}
