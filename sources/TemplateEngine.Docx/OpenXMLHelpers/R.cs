using System.Xml.Linq;

namespace TemplateEngine.Docx
{
    internal static class R
    {
        public static XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        public static XName embed = r + "embed";

        public static XName id = r + "id";
    }
}
