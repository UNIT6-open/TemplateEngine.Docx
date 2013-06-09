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
    }
}
