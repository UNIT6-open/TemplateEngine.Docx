using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TemplateEngine.Docx.Tests
{
	[TestClass]
	public class ListItemContentTests
	{
		[TestMethod]
		public void TableRowContentConstructorWithEnumerable_FillNameAndValue()
		{
			var listItemContent = new ListItemContent("Name", "Value");

			Assert.AreEqual("Name", listItemContent.Name);
			Assert.AreEqual("Value", listItemContent.Value);
		}

		[TestMethod]
		public void TableRowContentConstructorWithEnumerableFieldContent_FillNameAndValueAndFields()
		{
			var listItemContent = new ListItemContent("Name", "Value", new List<FieldContent>());

			Assert.AreEqual("Name", listItemContent.Name);
			Assert.AreEqual("Value", listItemContent.Value);
			Assert.IsNotNull(listItemContent.NestedFields);
		}

		[TestMethod]
		public void TableRowContentConstructorWithEnumerableListItemContent_FillNameAndValueAndFields()
		{
			var listItemContent = new ListItemContent("Name", "Value", new List<ListItemContent>());

			Assert.AreEqual("Name", listItemContent.Name);
			Assert.AreEqual("Value", listItemContent.Value);
			Assert.IsNotNull(listItemContent.NestedFields);
		}

		[TestMethod]
		public void TableRowContentConstructorWithFields_FillNameAndValueAndFields()
		{
			var listItemContent = new ListItemContent("Name", "Value", new ListItemContent(), new FieldContent());

			Assert.AreEqual("Name", listItemContent.Name);
			Assert.AreEqual("Value", listItemContent.Value);
			Assert.IsNotNull(listItemContent.NestedFields);
			Assert.AreEqual(2, listItemContent.NestedFields.Count());
		}
	}
}
