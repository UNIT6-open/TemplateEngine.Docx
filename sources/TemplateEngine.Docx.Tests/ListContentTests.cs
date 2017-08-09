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
			var listContent = new ListContent("Name", new ListItemContent[]{});

			Assert.IsNotNull(listContent.Items);
			Assert.AreEqual("Name", listContent.Name);
		}

		[TestMethod]
		public void ListContentConstructorWithNameAndItems_FillsNameAndItems()
		{
			var listContent = new ListContent("Name", new ListItemContent(), new ListItemContent());

			Assert.AreEqual(2, listContent.Items.Count);
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
					new ListItemContent("Subheader", "value")),
				new ListItemContent("Header", "value",
					new ListItemContent("Subheader", "value"),
					new ListItemContent("Subheader", "value2",
						new ListItemContent("Subsubheader", "value"))));

			Assert.IsNotNull(listContent.Items);
			Assert.AreEqual("Name", listContent.Name);
			Assert.AreEqual(3, listContent.FieldNames.Count());
			Assert.AreEqual("Header", listContent.FieldNames.ToArray()[0]);
			Assert.AreEqual("Subheader", listContent.FieldNames.ToArray()[1]);
			Assert.AreEqual("Subsubheader", listContent.FieldNames.ToArray()[2]);
		}

		[TestMethod]
		public void ListContentFluentConstructorWithNameAndItems_FillsNameAndItems()
		{
			var listContent = ListContent.Create("Name", new ListItemContent(), new ListItemContent());

			Assert.AreEqual(2, listContent.Items.Count);
			Assert.AreEqual("Name", listContent.Name);
		}

		[TestMethod]
		public void ListContentFluentConstructorWithNameAndEnumerableListItemContent_FillsNameAndItems()
		{
			var listContent = ListContent.Create("Name", new List<ListItemContent>());

			Assert.IsNotNull(listContent.Items);
			Assert.AreEqual("Name", listContent.Name);
		}


		[TestMethod]
		public void ListContentFluentAddItem_FillsNameAndItems()
		{
			var listContent = ListContent.Create("Name", new List<ListItemContent>())
				.AddItem(ListItemContent.Create("ItemName", "Name"));
			
			Assert.IsNotNull(listContent.Items);
			Assert.AreEqual("Name", listContent.Name);
			Assert.AreEqual(listContent.Items.Count, 1);
			Assert.AreEqual(listContent.Items.First().Fields.Count, 1);
			Assert.AreEqual(listContent.Items.First().Fields.First().Name, "ItemName");
			Assert.AreEqual(listContent.Items.First().Fields.First().Value, "Name");
		}

		[TestMethod]
		public void EqualsTest_ValuesAreEqual_Equals()
		{
			var firstListContent = new ListContent("Name",
			new ListItemContent("Header", "value",
				new ListItemContent("Subheader", "value")),
			new ListItemContent("Header", "value",
				new ListItemContent("Subheader", "value"),
				new ListItemContent("Subheader", "value2",
					new ListItemContent("Subsubheader", "value"))));

			var secondListContent = new ListContent("Name",
			new ListItemContent("Header", "value",
				new ListItemContent("Subheader", "value")),
			new ListItemContent("Header", "value",
				new ListItemContent("Subheader", "value"),
				new ListItemContent("Subheader", "value2",
					new ListItemContent("Subsubheader", "value"))));


			Assert.IsTrue(firstListContent.Equals(secondListContent));
		}


		[TestMethod]
		public void EqualsTest_ValuesDifferByName_NotEquals()
		{
			var firstListContent = new ListContent("Name1",
			new ListItemContent("Header", "value",
				new ListItemContent("Subheader", "value")),
			new ListItemContent("Header", "value",
				new ListItemContent("Subheader", "value"),
				new ListItemContent("Subheader", "value2",
					new ListItemContent("Subsubheader", "value"))));

			var secondListContent = new ListContent("Name2",
			new ListItemContent("Header", "value",
				new ListItemContent("Subheader", "value")),
			new ListItemContent("Header", "value",
				new ListItemContent("Subheader", "value"),
				new ListItemContent("Subheader", "value2",
					new ListItemContent("Subsubheader", "value"))));


			Assert.IsFalse(firstListContent.Equals(secondListContent));
		}

		[TestMethod]
		public void EqualsTest_ValuesDifferByListItemValue_NotEquals()
		{
			var firstListContent = new ListContent("Name",
			new ListItemContent("Header", "value",
				new ListItemContent("Subheader", "value")),
			new ListItemContent("Header", "value",
				new ListItemContent("Subheader", "value"),
				new ListItemContent("Subheader", "value1",
					new ListItemContent("Subsubheader", "value"))));

			var secondListContent = new ListContent("Name",
			new ListItemContent("Header", "value",
				new ListItemContent("Subheader", "value")),
			new ListItemContent("Header", "value",
				new ListItemContent("Subheader", "value"),
				new ListItemContent("Subheader", "value2",
					new ListItemContent("Subsubheader", "value"))));


			Assert.IsFalse(firstListContent.Equals(secondListContent));
		}

		[TestMethod]
		public void EqualsTest_ValuesDifferByNestedValue_NotEquals()
		{
			var firstListContent = new ListContent("Name",
			new ListItemContent("Header", "value",
				new ListItemContent("Subheader", "value")),
			new ListItemContent("Header", "value",
				new ListItemContent("Subheader", "value"),
				new ListItemContent("Subheader", "value",
					new ListItemContent("Subsubheader1", "value"))));

			var secondListContent = new ListContent("Name",
			new ListItemContent("Header", "value",
				new ListItemContent("Subheader", "value")),
			new ListItemContent("Header", "value",
				new ListItemContent("Subheader", "value"),
				new ListItemContent("Subheader", "value",
					new ListItemContent("Subsubheader2", "value"))));


			Assert.IsFalse(firstListContent.Equals(secondListContent));
		}
		[TestMethod]
		public void EqualsTest_CompareWithNull_NotEquals()
		{
			var firstListContent = new ListContent("Name",
			new ListItemContent("Header", "value",
				new ListItemContent("Subheader", "value")),
			new ListItemContent("Header", "value",
				new ListItemContent("Subheader", "value"),
				new ListItemContent("Subheader", "value",
					new ListItemContent("Subsubheader1", "value"))));
			
			Assert.IsFalse(firstListContent.Equals(null));
		}
	}
}
