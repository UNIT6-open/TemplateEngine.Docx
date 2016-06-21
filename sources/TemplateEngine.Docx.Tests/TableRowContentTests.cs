using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TemplateEngine.Docx.Tests
{
    [TestClass]
    public class TableRowContentTests
    {
        [TestMethod]
        public void TableRowContentConstructorWithEnumerable_FillsFields()
        {
            var tableRowContent = new TableRowContent(new List<FieldContent>());

            Assert.IsNotNull(tableRowContent.Fields);
        }

        [TestMethod]
        public void TableRowContentConstructorWithFields_FillsFields()
        {
            var tableRowContent = new TableRowContent(new FieldContent(), new FieldContent());

            Assert.AreEqual(2, tableRowContent.Fields.Count());
        }

		[TestMethod]
		public void EqualsTest_ValuesAreEqual_Equals()
		{
			var firstTableRow = new TableRowContent(
				new FieldContent("Name", "Eric"),
				new FieldContent("Role", "Program Manager"));

			var secondTableRow = new TableRowContent(
				new FieldContent("Name", "Eric"),
				new FieldContent("Role", "Program Manager"));

			Assert.IsTrue(firstTableRow.Equals(secondTableRow));
		}

		[TestMethod]
		public void EqualsTest_ValuesDifferByFieldName_NotEquals()
		{
			var firstTableRow = new TableRowContent(
				new FieldContent("Name1", "Eric"),
				new FieldContent("Role", "Program Manager"));

			var secondTableRow = new TableRowContent(
				new FieldContent("Name2", "Eric"),
				new FieldContent("Role", "Program Manager"));

			Assert.IsFalse(firstTableRow.Equals(secondTableRow));
		}

		[TestMethod]
		public void EqualsTest_ValuesDifferByFieldValue_NotEquals()
		{
			var firstTableRow = new TableRowContent(
				new FieldContent("Name", "Eric1"),
				new FieldContent("Role", "Program Manager"));

			var secondTableRow = new TableRowContent(
				new FieldContent("Name", "Eric2"),
				new FieldContent("Role", "Program Manager"));

			Assert.IsFalse(firstTableRow.Equals(secondTableRow));
		}

		[TestMethod]
		public void EqualsTest_ValuesDifferByFieldsCount_NotEquals()
		{
			var firstTableRow = new TableRowContent(
				new FieldContent("Name", "Eric"));

			var secondTableRow = new TableRowContent(
				new FieldContent("Name", "Eric"),
				new FieldContent("Role", "Program Manager"));

			Assert.IsFalse(firstTableRow.Equals(secondTableRow));
		}

		[TestMethod]
		public void EqualsTest_CompareWithNull_NotEquals()
		{
			var firstTableRow = new TableRowContent(
				new FieldContent("Name", "Eric"));
			
			Assert.IsFalse(firstTableRow.Equals(null));
		}
    }
}
