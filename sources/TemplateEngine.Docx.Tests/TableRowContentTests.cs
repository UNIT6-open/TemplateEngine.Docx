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
    }
}
