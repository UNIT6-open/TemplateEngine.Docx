using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace TemplateEngine.Docx
{
    internal static class W
    {
        public static XNamespace w =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public static XName body = w + "body";
        public static XName sdt = w + "sdt";
        public static XName sdtPr = w + "sdtPr";
        public static XName tag = w + "tag";
        public static XName val = w + "val";
        public static XName sdtContent = w + "sdtContent";
        public static XName tbl = w + "tbl";
        public static XName tr = w + "tr";
        public static XName tc = w + "tc";
        public static XName p = w + "p";
        public static XName r = w + "r";
        public static XName t = w + "t";
        public static XName rPr = w + "rPr";
        public static XName highlight = w + "highlight";
        public static XName pPr = w + "pPr";
        public static XName color = w + "color";
        public static XName sz = w + "sz";
        public static XName szCs = w + "szCs";
    }
}
