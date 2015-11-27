using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public class SpreadsheetWriterStyle
    {
    }

    public class CellStyleDfn : CellStyle
    {
        public HorizontalCellAlignment? HorizontalCellAlignment;

        public CellStyleNumFmt NumFmt;
        public CellStyleBorder Border;
        public CellStyleFill Fill;
        public CellStyleFont Font;
    }

    public class CellStyle
    {
        public int? id { get; protected internal set; }
    }

    public class CellStyleNumFmt : CellStyle
    {
        public string formatCode;

        public XElement ToXElement()
        {
            var result = new XElement(S.numFmt,
                new XAttribute(SSNoNamespace.numFmtId, this.id),
                new XAttribute(SSNoNamespace.formatCode, formatCode));
            return result;
        }
    }

    public class CellStyleBorder : CellStyle
    {
        //TODO
        public string LeftStyle;
        public string RIghtStyle;
        public string TopStyle;
        public string BottomStyle;
        public string DiagonalStyle;
    }

    public class CellStyleFill : CellStyle
    {
        // Example : Color = "FFFFFF00"; // yellow fill (ARGB)
        public string Color { get; set; }

        public XElement ToXElement()
        {
            XElement patternFill = null;
            if (this.Color != null)
            {
                var fgColor = new XElement(S.fgColor,
                                        new XAttribute(SSNoNamespace.rgb, this.Color));
                var bgColor = new XElement(S.bgColor,
                                        new XAttribute(NoNamespace.indexed, 64));
                // only "solid"
                patternFill = new XElement(S.patternFill,
                                        new XAttribute(SSNoNamespace.patternType, "solid"),
                                        fgColor,
                                        bgColor);
            }
            var result = new XElement(S.fill, patternFill);
            return result;
        }
    }

    public class CellStyleFont : CellStyle
    {
        public uint? Size;
        public string Name;
        public string Color;
        public bool? Bold;
        public bool? Italic;

        public XElement ToXElement()
        {
            XElement xsize = null;
            if (this.Size != null)
            {
                xsize = new XElement(S.sz, new XAttribute(SSNoNamespace.val, this.Size));
            }
            XElement xname = null;
            XElement xfamily = null;
            if (this.Name != null)
            {
                xname = new XElement(S.name, new XAttribute(SSNoNamespace.val, this.Name));
                xfamily = new XElement(S.family, new XAttribute(SSNoNamespace.val, (int)CellStyleFontFamilyEnum.Swiss));
            }
            XElement xcolor = null;
            if (this.Color != null)
            {
                xcolor = new XElement(S.color, new XAttribute(SSNoNamespace.rgb, this.Color));
            }
            XElement xbold = null;
            if (this.Bold == true)
            {
                xbold = new XElement(S.b);
            }
            XElement xitalic = null;
            if (this.Italic == true)
            {
                xitalic = new XElement(S.i);
            }
            var result = new XElement(S.font,
                xbold,
                xitalic,
                xsize,
                xcolor,
                xname,
                xfamily);
            return result;
        }
    }

    public enum CellStyleFontFamilyEnum
    {
        Roman = 1,
        Swiss = 2,
        Modern = 3,
        Script = 4,
        Decorative = 5,
    }
}
