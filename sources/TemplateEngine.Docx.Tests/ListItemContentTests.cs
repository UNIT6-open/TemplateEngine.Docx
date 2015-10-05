using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TemplateEngine.Docx.Tests
{
	[TestClass]
	public class ListItemContentTests
	{
		[TestMethod]
		public void ListItemContentConstructorWithEnumerable_FillNameAndValue()
		{
			var listItemContent = new ListItemContent("Name", "Value");

			Assert.AreEqual(listItemContent.Fields.Count(), 1);
			Assert.AreEqual("Name", listItemContent.Fields.First().Name);
			Assert.AreEqual("Value", listItemContent.Fields.First().Value);
		}

		[TestMethod]
		public void ListItemContentConstructorWithEnumerable_FillNameAndValueAndFields()
		{
			var listItemContent = new ListItemContent("Name", "Value", new List<ListItemContent>());

			Assert.AreEqual(listItemContent.Fields.Count(), 1);
			Assert.AreEqual("Name", listItemContent.Fields.First().Name);
			Assert.AreEqual("Value", listItemContent.Fields.First().Value);
			Assert.IsNotNull(listItemContent.NestedFields);
		}

		[TestMethod]
		public void ListItemContentFluentConstructorWithEnumerable_FillNameAndValue()
		{
			var listItemContent = ListItemContent.Create("Name", "Value");

			Assert.AreEqual(listItemContent.Fields.Count(), 1);
			Assert.AreEqual("Name", listItemContent.Fields.First().Name);
			Assert.AreEqual("Value", listItemContent.Fields.First().Value);
		}
		[TestMethod]
		public void ListItemContentFluentConstructorWithEnumerable_FillNameAndValueAndFields()
		{
			var listItemContent = ListItemContent.Create("Name", "Value", new List<ListItemContent>());

			Assert.AreEqual(listItemContent.Fields.Count(), 1);
			Assert.AreEqual("Name", listItemContent.Fields.First().Name);
			Assert.AreEqual("Value", listItemContent.Fields.First().Value);
			Assert.IsNotNull(listItemContent.NestedFields);
		}

		[TestMethod]
		public void ListItemContentFluentAddItem_FillsField()
		{
			var listItemContent = new ListItemContent("Name1", "Value1").AddField("Name2", "Value2");

			Assert.AreEqual(listItemContent.Fields.Count(), 2);
			Assert.AreEqual("Name1", listItemContent.Fields.First().Name);
			Assert.AreEqual("Value1", listItemContent.Fields.First().Value);
			Assert.AreEqual("Name2", listItemContent.Fields.Last().Name);
			Assert.AreEqual("Value2", listItemContent.Fields.Last().Value);
			Assert.IsNotNull(listItemContent.NestedFields);
		}
		[TestMethod]
		public void ListItemContentFluentAddNestedItem_FillsNestedField()
		{
			var listItemContent = new ListItemContent("Name1", "Value1")
				.AddNestedItem(ListItemContent.Create("NestedName", "NestedValue"));

			Assert.AreEqual(listItemContent.Fields.Count(), 1);
			Assert.AreEqual("Name1", listItemContent.Fields.First().Name);
			Assert.AreEqual("Value1", listItemContent.Fields.First().Value);
			Assert.AreEqual(listItemContent.NestedFields.Count, 1);
			Assert.AreEqual(listItemContent.NestedFields.First().Fields.Count, 1);
			Assert.AreEqual(listItemContent.NestedFields.First().Fields.First().Name, "NestedName");
			Assert.AreEqual(listItemContent.NestedFields.First().Fields.First().Value, "NestedValue");
		}
	}
}
