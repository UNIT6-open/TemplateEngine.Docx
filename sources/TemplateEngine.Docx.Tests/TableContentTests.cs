using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
    }
}
