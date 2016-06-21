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

		[TestMethod]
		public void EqualsTest_ValuesAreEqual_Equals()
		{
			var firstItemContent = new ListItemContent("Name", "Value")
				.AddNestedItem(ListItemContent.Create("NestedName", "NestedValue"));

			var secondItemContent = new ListItemContent("Name", "Value")
				.AddNestedItem(ListItemContent.Create("NestedName", "NestedValue"));
			
			Assert.IsTrue(firstItemContent.Equals(secondItemContent));
		}

		[TestMethod]
		public void EqualsTest_ValuesDifferByName_NotEquals()
		{
			var firstItemContent = new ListItemContent("Name1", "Value")
				.AddNestedItem(ListItemContent.Create("NestedName", "NestedValue"));

			var secondItemContent = new ListItemContent("Name2", "Value")
				.AddNestedItem(ListItemContent.Create("NestedName", "NestedValue"));
			
			Assert.IsFalse(firstItemContent.Equals(secondItemContent));
		}

		[TestMethod]
		public void EqualsTest_ValuesDifferByNestedValueName_NotEquals()
		{
			var firstItemContent = new ListItemContent("Name", "Value")
				.AddNestedItem(ListItemContent.Create("NestedName1", "NestedValue"));

			var secondItemContent = new ListItemContent("Name", "Value")
				.AddNestedItem(ListItemContent.Create("NestedName2", "NestedValue"));
			
			Assert.IsFalse(firstItemContent.Equals(secondItemContent));
		}
		[TestMethod]
		public void EqualsTest_ValuesDifferByNestedValuesCounts_NotEquals()
		{
			var firstItemContent = new ListItemContent("Name", "Value")
				.AddNestedItem(ListItemContent.Create("NestedName1", "NestedValue"));

			var secondItemContent = new ListItemContent("Name", "Value");
			
			Assert.IsFalse(firstItemContent.Equals(secondItemContent));
		}

		[TestMethod]
		public void EqualsTest_CompareWithNull_NotEquals()
		{
			var firstItemContent = new ListItemContent("Name", "Value")
				.AddNestedItem(ListItemContent.Create("NestedName1", "NestedValue"));
			
			Assert.IsFalse(firstItemContent.Equals(null));
		}
	}
}
