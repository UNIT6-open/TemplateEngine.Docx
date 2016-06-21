using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TemplateEngine.Docx.Tests
{
    [TestClass]
    public class TableContentTests
    {
        [TestMethod]
        public void TableContentConstrictorWithName_FillsName()
        {
            var tableContent = new TableContent("Name");

            Assert.AreEqual("Name", tableContent.Name);
        }

        [TestMethod]
        public void TableContentConstructorWithNameAndEnumerable_FillsNameAndRows()
        {
            var tableContent = new TableContent("Name", new List<TableRowContent>());

            Assert.IsNotNull(tableContent.Rows);
            Assert.AreEqual("Name", tableContent.Name);
        }

        [TestMethod]
        public void TableContentConstructorWithNameAndRows_FillsNameAndRows()
        {
            var tableContent = new TableContent("Name", new TableRowContent(), new TableRowContent());

            Assert.AreEqual(2, tableContent.Rows.Count());
            Assert.AreEqual("Name", tableContent.Name);
        }

		[TestMethod]
		public void TableContentFluentConstructorWithNameAndEnumerable_FillsNameAndRows()
		{
			var tableContent = TableContent.Create("Name", new List<TableRowContent>());

			Assert.IsNotNull(tableContent.Rows);
			Assert.AreEqual("Name", tableContent.Name);
		}

        [TestMethod]
        public void TableContentFluentConstructorWithNameAndRows_FillsNameAndRows()
        {
            var tableContent = TableContent.Create("Name", new TableRowContent(), new TableRowContent());

            Assert.AreEqual(2, tableContent.Rows.Count());
            Assert.AreEqual("Name", tableContent.Name);
        }
        [TestMethod]
        public void TableAddRowFluent_AddsRow()
        {
            var tableContent = TableContent.Create("Name")
				.AddRow(new FieldContent());

            Assert.AreEqual(1, tableContent.Rows.Count());
            Assert.AreEqual("Name", tableContent.Name);
        }


		[TestMethod]
		public void EqualsTest_ValuesAreEqual_Equals()
		{
			var firstTableContent =
				new TableContent("Team Members Table")
					.AddRow(
						new FieldContent("Name", "Eric"),
						new FieldContent("Role", "Program Manager"))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer"));

			var secondTableContent =
				new TableContent("Team Members Table")
					.AddRow(
						new FieldContent("Name", "Eric"),
						new FieldContent("Role", "Program Manager"))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer"));

			
			Assert.IsTrue(firstTableContent.Equals(secondTableContent));
		}

		[TestMethod]
		public void EqualsTest_ValuesDifferByName_NotEquals()
		{
			var firstTableContent =
				new TableContent("Team Members Table1")
					.AddRow(
						new FieldContent("Name", "Eric"),
						new FieldContent("Role", "Program Manager"))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer"));

			var secondTableContent =
				new TableContent("Team Members Table2")
					.AddRow(
						new FieldContent("Name", "Eric"),
						new FieldContent("Role", "Program Manager"))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer"));

			
			Assert.IsFalse(firstTableContent.Equals(secondTableContent));
		}
		[TestMethod]
		public void EqualsTest_ValuesDifferByFieldName_NotEquals()
		{
			var firstTableContent =
				new TableContent("Team Members Table")
					.AddRow(
						new FieldContent("Name1", "Eric"),
						new FieldContent("Role", "Program Manager"))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer"));

			var secondTableContent =
				new TableContent("Team Members Table")
					.AddRow(
						new FieldContent("Name2", "Eric"),
						new FieldContent("Role", "Program Manager"))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer"));

			
			Assert.IsFalse(firstTableContent.Equals(secondTableContent));
		}
		[TestMethod]
		public void EqualsTest_ValuesDifferByFieldCount_NotEquals()
		{
			var firstTableContent =
				new TableContent("Team Members Table")
					.AddRow(
						new FieldContent("Name", "Eric"))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer"));

			var secondTableContent =
				new TableContent("Team Members Table")
					.AddRow(
						new FieldContent("Name", "Eric"),
						new FieldContent("Role", "Program Manager"))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer"));

			
			Assert.IsFalse(firstTableContent.Equals(secondTableContent));
		}

		[TestMethod]
		public void EqualsTest_ValuesDifferByRowCount_NotEquals()
		{
			var firstTableContent =
				new TableContent("Team Members Table")
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer"));

			var secondTableContent =
				new TableContent("Team Members Table")
					.AddRow(
						new FieldContent("Name", "Eric"),
						new FieldContent("Role", "Program Manager"))
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer"));

			
			Assert.IsFalse(firstTableContent.Equals(secondTableContent));
		}

		[TestMethod]
		public void EqualsTest_CompareWithNull_NotEquals()
		{
			var firstTableContent =
				new TableContent("Team Members Table")
					.AddRow(
						new FieldContent("Name", "Bob"),
						new FieldContent("Role", "Developer"));
			
			Assert.IsFalse(firstTableContent.Equals(null));
		}

    }
}
