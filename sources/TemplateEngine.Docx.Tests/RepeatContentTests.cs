using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TemplateEngine.Docx.Tests
{
	[TestClass]
	public class RepeatContentTests
	{

		[TestMethod]
		public void RepeatContentConstructorWithName_FillsName()
		{
			var repeatContent = new RepeatContent("Name");

            Assert.AreEqual("Name", repeatContent.Name);
		}

		[TestMethod]
        public void RepeatContentConstructorWithNameAndEnumerableFieldContent_FillsNameAndItems()
		{
            var repeatContent = new RepeatContent("Name", new Content[]{});

            Assert.IsNotNull(repeatContent.Items);
            Assert.AreEqual("Name", repeatContent.Name);
		}

		[TestMethod]
        public void RepeatContentConstructorWithNameAndItems_FillsNameAndItems()
		{
            var repeatContent = new RepeatContent("Name", new Content(), new Content());

            Assert.AreEqual(2, repeatContent.Items.Count);
            Assert.AreEqual("Name", repeatContent.Name);
		}

		[TestMethod]
        public void RepeatContentConstructorWithNameAndEnumerableListItemContent_FillsNameAndItems()
		{
            var repeatContent = new RepeatContent("Name", new List<Content>());

            Assert.IsNotNull(repeatContent.Items);
            Assert.AreEqual("Name", repeatContent.Name);
		}

		[TestMethod]
        public void RepeatContentGetFieldnames()
		{
            var repeatContent = new RepeatContent("Name",
		        new Content(new FieldContent("Field1", "value"), new FieldContent("Field2", "value")),
		        new Content(new FieldContent("Field1", "value"), new FieldContent("Field2", "value")));

		    var fieldNames = repeatContent.FieldNames.ToArray();

            Assert.IsNotNull(repeatContent.Items);
            Assert.AreEqual("Name", repeatContent.Name);
            Assert.AreEqual(2, fieldNames.Length);
            Assert.AreEqual("Field1", fieldNames[0]);
            Assert.AreEqual("Field2", fieldNames[1]);
		}

		[TestMethod]
		public void ListContentFluentConstructorWithNameAndItems_FillsNameAndItems()
		{
            var repeatContent = RepeatContent.Create("Name", new Content(), new Content());

            Assert.AreEqual(2, repeatContent.Items.Count);
            Assert.AreEqual("Name", repeatContent.Name);
		}

		[TestMethod]
		public void ListContentFluentConstructorWithNameAndEnumerableListItemContent_FillsNameAndItems()
		{
            var repeatContent = RepeatContent.Create("Name", new List<Content>());

            Assert.IsNotNull(repeatContent.Items);
            Assert.AreEqual("Name", repeatContent.Name);
		}


		[TestMethod]
		public void ListContentFluentAddItem_FillsNameAndItems()
		{
            var repeatContent = RepeatContent.Create("Name", new List<Content>())
                .AddItem(new Content(new FieldContent("ItemName", "Name")));

            Assert.IsNotNull(repeatContent.Items);
            Assert.AreEqual("Name", repeatContent.Name);
            Assert.AreEqual(repeatContent.Items.Count, 1);
            Assert.AreEqual(repeatContent.Items.First().Fields.Count, 1);
            Assert.AreEqual(repeatContent.Items.First().Fields.First().Name, "ItemName");
            Assert.AreEqual(repeatContent.Items.First().Fields.First().Value, "Name");
		}

		[TestMethod]
		public void EqualsTest_ValuesAreEqual_Equals()
		{
			var firstRepeatContent = new RepeatContent("Name",
                new Content(new FieldContent("Field1", "value")),
                new Content(new FieldContent("Field1", "value2")));

            var secondRepeatContent = new RepeatContent("Name",
                new Content(new FieldContent("Field1", "value")),
                new Content(new FieldContent("Field1", "value2")));


            Assert.IsTrue(firstRepeatContent.Equals(secondRepeatContent));
		}

		[TestMethod]
		public void EqualsTest_ValuesDifferByName_NotEquals()
		{
            var firstRepeatContent = new RepeatContent("Name",
                new Content(new FieldContent("Field1", "value")),
                new Content(new FieldContent("Field1", "value2")));

            var secondRepeatContent = new RepeatContent("Name",
                new Content(new FieldContent("Field2", "value")),
                new Content(new FieldContent("Field1", "value2")));


            Assert.IsFalse(firstRepeatContent.Equals(secondRepeatContent));
		}

		[TestMethod]
		public void EqualsTest_ValuesDifferByItemValue_NotEquals()
		{
            var firstRepeatContent = new RepeatContent("Name",
                new Content(new FieldContent("Field1", "value")),
                new Content(new FieldContent("Field1", "value2")));

            var secondRepeatContent = new RepeatContent("Name",
                new Content(new FieldContent("Field1", "value")),
                new Content(new FieldContent("Field1", "anotherValue")));


            Assert.IsFalse(firstRepeatContent.Equals(secondRepeatContent));
		}

        [TestMethod]
        public void EqualsTest_ContentsDifferByName_NotEquals()
        {
            var firstRepeatContent = new RepeatContent("Name",
                new Content(new FieldContent("Field1", "value")),
                new Content(new FieldContent("Field1", "value2")));

            var secondRepeatContent = new RepeatContent("AnotherName",
                new Content(new FieldContent("Field1", "value")),
                new Content(new FieldContent("Field1", "value2")));


            Assert.IsFalse(firstRepeatContent.Equals(secondRepeatContent));
        }

		[TestMethod]
		public void EqualsTest_CompareWithNull_NotEquals()
		{
            var firstRepeatContent = new RepeatContent("Name",
                new Content(new FieldContent("Field1", "value")),
                new Content(new FieldContent("Field1", "value2")));

            Assert.IsFalse(firstRepeatContent.Equals(null));
		}
	}
}
