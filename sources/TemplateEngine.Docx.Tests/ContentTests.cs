using System;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;

namespace TemplateEngine.Docx.Tests
{
	[TestClass]
	public class ContentTests
	{
		[TestMethod]
		public void ContentSerializationTest_SerializeToJson_Success()
		{
			var valuesToFill = new Content(
				// Add field.
				new FieldContent("Report date", new DateTime(2000, 01, 01).ToShortDateString()),
				
				// Add table.
				new TableContent("Team Members Table")
					.AddRow(
						new FieldContent("Name", "Eric"),
						new FieldContent("Role", "Program Manager"))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer")),
				// Add nested list.	
				new ListContent("Team Members Nested List")
					.AddItem(new ListItemContent("Role", "Program Manager")
						.AddNestedItem(new FieldContent("Name", "Eric"))
						.AddNestedItem(new FieldContent("Name", "Ann")))
					.AddItem(new ListItemContent("Role", "Developer")
						.AddNestedItem(new FieldContent("Name", "Bob"))
						.AddNestedItem(new FieldContent("Name", "Richard"))),
				// Add image
				new ImageContent("photo", new byte[]{1, 2, 3}),

                // Add repeat content
				new RepeatContent("Repeat")
                    .AddItem(new FieldContent("Weekend", "Saturday"))
                    .AddItem(new FieldContent("Weekend", "Sunday"))
				
				);


			var serialized = JsonConvert.SerializeObject(valuesToFill);

			const string expectedSerialized =
				"{\"Repeats\":[{\"Items\":[{\"Repeats\":[],\"Tables\":[],\"Lists\":[],\"Fields\":[{\"Value\":\"Saturday\",\"Name\":\"Weekend\",\"IsHidden\":false}],\"Images\":[]},{\"Repeats\":[],\"Tables\":[],\"Lists\":[],\"Fields\":[{\"Value\":\"Sunday\",\"Name\":\"Weekend\",\"IsHidden\":false}],\"Images\":[]}],\"FieldNames\":[\"Weekend\"],\"Name\":\"Repeat\",\"IsHidden\":false}],\"Tables\":[{\"Rows\":[{\"Repeats\":[],\"Tables\":[],\"Lists\":[],\"Fields\":[{\"Value\":\"Eric\",\"Name\":\"Name\",\"IsHidden\":false},{\"Value\":\"Program Manager\",\"Name\":\"Role\",\"IsHidden\":false}],\"Images\":[]},{\"Repeats\":[],\"Tables\":[],\"Lists\":[],\"Fields\":[{\"Value\":\"Bob\",\"Name\":\"Name\",\"IsHidden\":false},{\"Value\":\"Developer\",\"Name\":\"Role\",\"IsHidden\":false}],\"Images\":[]}],\"FieldNames\":[\"Name\",\"Role\"],\"Name\":\"Team Members Table\",\"IsHidden\":false}],\"Lists\":[{\"Items\":[{\"NestedFields\":[{\"NestedFields\":null,\"Repeats\":[],\"Tables\":[],\"Lists\":[],\"Fields\":[{\"Value\":\"Eric\",\"Name\":\"Name\",\"IsHidden\":false}],\"Images\":[]},{\"NestedFields\":null,\"Repeats\":[],\"Tables\":[],\"Lists\":[],\"Fields\":[{\"Value\":\"Ann\",\"Name\":\"Name\",\"IsHidden\":false}],\"Images\":[]}],\"Repeats\":[],\"Tables\":[],\"Lists\":[],\"Fields\":[{\"Value\":\"Program Manager\",\"Name\":\"Role\",\"IsHidden\":false}],\"Images\":[]},{\"NestedFields\":[{\"NestedFields\":null,\"Repeats\":[],\"Tables\":[],\"Lists\":[],\"Fields\":[{\"Value\":\"Bob\",\"Name\":\"Name\",\"IsHidden\":false}],\"Images\":[]},{\"NestedFields\":null,\"Repeats\":[],\"Tables\":[],\"Lists\":[],\"Fields\":[{\"Value\":\"Richard\",\"Name\":\"Name\",\"IsHidden\":false}],\"Images\":[]}],\"Repeats\":[],\"Tables\":[],\"Lists\":[],\"Fields\":[{\"Value\":\"Developer\",\"Name\":\"Role\",\"IsHidden\":false}],\"Images\":[]}],\"FieldNames\":[\"Role\",\"Name\"],\"Name\":\"Team Members Nested List\",\"IsHidden\":false}],\"Fields\":[{\"Value\":\"01.01.2000\",\"Name\":\"Report date\",\"IsHidden\":false}],\"Images\":[{\"Binary\":\"AQID\",\"Name\":\"photo\",\"IsHidden\":false}]}";

            Assert.AreEqual(expectedSerialized, serialized);

		}

		[TestMethod]
		public void ContentDeserializationTest_DeserializeFromJson_Success()
		{
			var valuesToFill = new Content(
				// Add field.
				new FieldContent("Report date", new DateTime(2000, 01, 01).ToShortDateString()),
				// Add table.
				new TableContent("Team Members Table")
					.AddRow(
						new FieldContent("Name", "Eric"),
						new FieldContent("Role", "Program Manager"))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer")),
				// Add nested list.	
				new ListContent("Team Members Nested List")
					.AddItem(new ListItemContent("Role", "Program Manager")
						.AddNestedItem(new FieldContent("Name", "Eric"))
						.AddNestedItem(new FieldContent("Name", "Ann")))
					.AddItem(new ListItemContent("Role", "Developer")
						.AddNestedItem(new FieldContent("Name", "Bob"))
						.AddNestedItem(new FieldContent("Name", "Richard"))),
				// Add image
				new ImageContent("photo", new byte[] { 1, 2, 3 })
				);


			const string serialized = "{\"Tables\":[{\"Name\":\"Team Members Table\",\"Rows\":[{\"Tables\":[],\"Lists\":[],\"Fields\":[{\"Name\":\"Name\",\"Value\":\"Eric\"},{\"Name\":\"Role\",\"Value\":\"Program Manager\"}],\"Images\":[]},{\"Tables\":[],\"Lists\":[],\"Fields\":[{\"Name\":\"Name\",\"Value\":\"Bob\"},{\"Name\":\"Role\",\"Value\":\"Developer\"}],\"Images\":[]}],\"FieldNames\":[\"Name\",\"Role\"]}],\"Lists\":[{\"Name\":\"Team Members Nested List\",\"Items\":[{\"NestedFields\":[{\"NestedFields\":null,\"Tables\":[],\"Lists\":[],\"Fields\":[{\"Name\":\"Name\",\"Value\":\"Eric\"}],\"Images\":[]},{\"NestedFields\":null,\"Tables\":[],\"Lists\":[],\"Fields\":[{\"Name\":\"Name\",\"Value\":\"Ann\"}],\"Images\":[]}],\"Tables\":[],\"Lists\":[],\"Fields\":[{\"Name\":\"Role\",\"Value\":\"Program Manager\"}],\"Images\":[]},{\"NestedFields\":[{\"NestedFields\":null,\"Tables\":[],\"Lists\":[],\"Fields\":[{\"Name\":\"Name\",\"Value\":\"Bob\"}],\"Images\":[]},{\"NestedFields\":null,\"Tables\":[],\"Lists\":[],\"Fields\":[{\"Name\":\"Name\",\"Value\":\"Richard\"}],\"Images\":[]}],\"Tables\":[],\"Lists\":[],\"Fields\":[{\"Name\":\"Role\",\"Value\":\"Developer\"}],\"Images\":[]}],\"FieldNames\":[\"Role\",\"Name\"]}],\"Fields\":[{\"Name\":\"Report date\",\"Value\":\"01.01.2000\"}],\"Images\":[{\"Name\":\"photo\",\"Binary\":\"AQID\"}]}";

			var deserialized = JsonConvert.DeserializeObject<Content>(serialized);

			Assert.IsTrue(valuesToFill.Equals(deserialized));
		}

		[TestMethod]
		public void ContentSerializationTest_EmptyContentSerializeToJson_Success()
		{
			var valuesToFill = new Content();

			var serialized = JsonConvert.SerializeObject(valuesToFill);

			Assert.AreEqual("{\"Repeats\":[],\"Tables\":[],\"Lists\":[],\"Fields\":[],\"Images\":[]}", serialized);
		}

		[TestMethod]
		public void ContentDeserializationTest_EmptyContentDeserializeFromJson_Success()
		{
			var valuesToFill = new Content();

            const string serialized = "{\"Repeats\":[],\"Tables\":[],\"Lists\":[],\"Fields\":[],\"Images\":[]}";

			var deserialized = JsonConvert.DeserializeObject<Content>(serialized);

			Assert.IsTrue(valuesToFill.Equals(deserialized));
		}

		[TestMethod]
		public void EqualsTest_ObjectsAreEqual_Equals()
		{
			
			var firstValuesToFill = new Content(
				// Add field.
				new FieldContent("Report date", new DateTime(2000, 01, 01).ToShortDateString()),
				// Add table.
				new TableContent("Team Members Table")
					.AddRow(
						new FieldContent("Name", "Eric"),
						new FieldContent("Role", "Program Manager"))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer")),
				// Add nested list.	
				new ListContent("Team Members Nested List")
					.AddItem(new ListItemContent("Role", "Program Manager")
						.AddNestedItem(new FieldContent("Name", "Eric"))
						.AddNestedItem(new FieldContent("Name", "Ann")))
					.AddItem(new ListItemContent("Role", "Developer")
						.AddNestedItem(new FieldContent("Name", "Bob"))
						.AddNestedItem(new FieldContent("Name", "Richard"))),
				// Add image
				new ImageContent("photo", new byte[] { 1, 2, 3 })
				);

			var secondValuesToFill = new Content(
				// Add field.
				new FieldContent("Report date", new DateTime(2000, 01, 01).ToShortDateString()),
				// Add table.
				new TableContent("Team Members Table")
					.AddRow(
						new FieldContent("Name", "Eric"),
						new FieldContent("Role", "Program Manager"))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer")),
				// Add nested list.	
				new ListContent("Team Members Nested List")
					.AddItem(new ListItemContent("Role", "Program Manager")
						.AddNestedItem(new FieldContent("Name", "Eric"))
						.AddNestedItem(new FieldContent("Name", "Ann")))
					.AddItem(new ListItemContent("Role", "Developer")
						.AddNestedItem(new FieldContent("Name", "Bob"))
						.AddNestedItem(new FieldContent("Name", "Richard"))),
				// Add image
				new ImageContent("photo", new byte[] { 1, 2, 3 })
				);

			Assert.IsTrue(firstValuesToFill.Equals(secondValuesToFill));

		}


		[TestMethod]
		public void EqualsTest_ObjectsDifferByItemsCount_NotEquals()
		{

			var firstValuesToFill = new Content(
				// Add field.
				new FieldContent("Report date", new DateTime(2000, 01, 01).ToShortDateString()),
				// Add table.
				new TableContent("Team Members Table")
					.AddRow(
						new FieldContent("Name", "Eric"),
						new FieldContent("Role", "Program Manager"))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer"))
				);

			var secondValuesToFill = new Content(
				// Add field.
				new FieldContent("Report date", new DateTime(2000, 01, 01).ToShortDateString())
				);

			Assert.IsFalse(firstValuesToFill.Equals(secondValuesToFill));

		}

		[TestMethod]
		public void EqualsTest_CopareWithNull_NotEquals()
		{

			var firstValuesToFill = new Content(
				// Add field.
				new FieldContent("Report date", new DateTime(2000, 01, 01).ToShortDateString()),
				// Add table.
				new TableContent("Team Members Table")
					.AddRow(
						new FieldContent("Name", "Eric"),
						new FieldContent("Role", "Program Manager"))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer"))
				);

			Assert.IsFalse(firstValuesToFill.Equals(null));

		}
	}
}
