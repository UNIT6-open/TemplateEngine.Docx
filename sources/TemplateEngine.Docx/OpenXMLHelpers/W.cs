using System.Xml.Linq;

namespace TemplateEngine.Docx
{
    internal static class W
    {
        public static XNamespace w =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public static XName body = w + "body";
		public static XName header = w + "hdr";
		public static XName footer = w + "ftr";
        public static XName sdt = w + "sdt";
        public static XName sdtPr = w + "sdtPr";
        public static XName tag = w + "tag";
        public static XName val = w + "val";
        public static XName sdtContent = w + "sdtContent";
        public static XName tbl = w + "tbl";
        public static XName tr = w + "tr";
        public static XName tc = w + "tc";
        public static XName tcPr = w + "tcPr";
        public static XName p = w + "p";
        public static XName r = w + "r";
        public static XName t = w + "t";
        public static XName rPr = w + "rPr";
        public static XName highlight = w + "highlight";
        public static XName pPr = w + "pPr";
        public static XName color = w + "color";
        public static XName sz = w + "sz";
        public static XName szCs = w + "szCs";
        public static XName vMerge = w + "vMerge";
        public static XName numId = w + "numId";
        public static XName numPr= w + "numPr";
        public static XName ilvl= w + "ilvl";
        public static XName num= w + "num";
		public static XName abstractNumId = w + "abstractNumId";
		public static XName abstractNum = w + "abstractNum";
		public static XName nsid = w + "nsid";
		public static XName lvlOverride = w + "lvlOverride";
		public static XName startOverride = w + "startOverride";
		public static XName lvl = w + "lvl";
		public static XName start = w + "start";
		public static XName style = w + "style";
		public static XName styleId = w + "styleId";
		public static XName numStyleLink = w + "numStyleLink";
		public static XName pStyle = w + "pStyle";
		public static XName lvlRestart = w + "lvlRestart";
		public static XName numFmt = w + "numFmt";
		public static XName lvlText = w + "lvlText";
		public static XName type = w + "type";
		public static XName isLgl = w + "isLgl";
		public static XName rStyle = w + "rStyle";
		public static XName br = w + "br";
        public static XName drawing = w + "drawing";
    }
}
