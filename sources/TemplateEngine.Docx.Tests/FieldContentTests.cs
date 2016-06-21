using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TemplateEngine.Docx.Tests
{
    [TestClass]
    public class FieldContentTests
    {
        [TestMethod]
        public void FieldContentConstructorWithArguments_FillNameAndValue()
        {
            var fieldContent = new FieldContent("Name", "Value");

            Assert.AreEqual("Name", fieldContent.Name);
            Assert.AreEqual("Value", fieldContent.Value);
        }

        [TestMethod]
        public void EqualsTest_ValuesAreEquel_Equals()
        {
            var firstFieldContent = new FieldContent("Name", "Value");
            var secondFieldContent = new FieldContent("Name", "Value");

            Assert.IsTrue(firstFieldContent.Equals(secondFieldContent));
        }

        [TestMethod]
        public void EqualsTest_ValuesAreNotEqual_NotEquals()
        {
            var firstFieldContent = new FieldContent("Name", "Value");
            var secondFieldContent = new FieldContent("Name", "Value2");

            Assert.IsFalse(firstFieldContent.Equals(secondFieldContent));
        }
        [TestMethod]
        public void EqualsTest_CompareWithNull_NotEquals()
        {
            var firstFieldContent = new FieldContent("Name", "Value");
            var secondFieldContent = new FieldContent("Name", "Value2");

            Assert.IsFalse(firstFieldContent.Equals(null));
        }
    }
}
