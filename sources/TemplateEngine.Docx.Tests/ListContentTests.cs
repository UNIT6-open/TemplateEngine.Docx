using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TemplateEngine.Docx.Tests
{
	[TestClass]
	public class ListContentTests
	{

		[TestMethod]
		public void ListContentConstructorWithName_FillsName()
		{
			var listContent = new ListContent("Name");

			Assert.AreEqual("Name", listContent.Name);
		}

		[TestMethod]
		public void ListContentConstructorWithNameAndEnumerableFieldContent_FillsNameAndItems()
		{
			var listContent = new ListContent("Name", new List<FieldContent>());

			Assert.IsNotNull(listContent.Items);
			Assert.AreEqual("Name", listContent.Name);
		}

		[TestMethod]
		public void ListContentConstructorWithNameAndItems_FillsNameAndItems()
		{
			var listContent = new ListContent("Name", new ListItemContent(), new FieldContent());

			Assert.AreEqual(2, listContent.Items.Count());
			Assert.AreEqual("Name", listContent.Name);
		}

		[TestMethod]
		public void ListContentConstructorWithNameAndEnumerableListItemContent_FillsNameAndItems()
		{
			var listContent = new ListContent("Name", new List<ListItemContent>());

			Assert.IsNotNull(listContent.Items);
			Assert.AreEqual("Name", listContent.Name);
		}

		[TestMethod]
		public void ListContentGetFieldnames()
		{
			var listContent = new ListContent("Name", 
				new ListItemContent("Header", "value", 
					new FieldContent("Subheader", "value")),
				new ListItemContent("Header", "value",
					new FieldContent("Subheader", "value"),
					new ListItemContent("Subheader", "value2", 
						new FieldContent("Subsubheader", "value"))));

			Assert.IsNotNull(listContent.Items);
			Assert.AreEqual("Name", listContent.Name);
			Assert.AreEqual(3, listContent.FieldNames.Count());
			Assert.AreEqual("Header", listContent.FieldNames.ToArray()[0]);
			Assert.AreEqual("Subheader", listContent.FieldNames.ToArray()[1]);
			Assert.AreEqual("Subsubheader", listContent.FieldNames.ToArray()[2]);
		}		
	}
}
