using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public class ColDfn
    {
        public decimal? Width;
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
        public string Color;

        public string LeftStyle;
        public string LeftColor;

        public string RightStyle;
        public string RightColor;

        public string TopStyle;
        public string TopColor;

        public string BottomStyle;
        public string BottomColor;


        public string DiagonalStyle;
        public string DiagonalColor;
        public bool DiagonalUp;
        public bool DiagonalDown;

        //value of border style
        public const string Thin = "thin";
        public const string Medium = "medium";
        public const string Dashed = "dashed";
        public const string Dotted = "dotted";
        public const string Thick = "thick";
        public const string Double = "double";
        public const string Hair = "hair";
        public const string MediumDashed = "mediumDashed";
        public const string DashDot = "dashDot";
        public const string MediumDashDot = "mediumDashDot";
        public const string DashDotDot = "dashDotDot";
        public const string MediumDashDotDot = "mediumDashDotDot";

        public static CellStyleBorder CreateBoxBorder(string style, string color = null)
        {
            var result = new CellStyleBorder();
            result.LeftStyle = style;
            result.RightStyle = style;
            result.TopStyle = style;
            result.BottomStyle = style;
            result.Color = color;
            return result;
        }

        public XElement ToXElement()
        {
            XElement color = CellStyleUtil.CreateColorXElement(this.Color);
            XElement left = null;
            XElement right = null;
            XElement top = null;
            XElement bottom = null;
            XElement diagonal = null;
            XAttribute diagonalUp = null;
            XAttribute diagonalDown = null;
            if (this.LeftStyle != null)
            {
                left = new XElement(S.left,
                    new XAttribute(SSNoNamespace.style, this.LeftStyle),
                    CellStyleUtil.CreateColorXElement(this.LeftColor) ?? color);
            }
            if (this.RightStyle != null)
            {
                right = new XElement(S.right,
                    new XAttribute(SSNoNamespace.style, this.RightStyle),
                    CellStyleUtil.CreateColorXElement(this.RightColor) ?? color);
            }
            if (this.TopStyle != null)
            {
                top = new XElement(S.top,
                    new XAttribute(SSNoNamespace.style, this.TopStyle),
                    CellStyleUtil.CreateColorXElement(this.TopColor) ?? color);
            }
            if (this.BottomStyle != null)
            {
                bottom = new XElement(S.bottom,
                    new XAttribute(SSNoNamespace.style, this.BottomStyle),
                    CellStyleUtil.CreateColorXElement(this.BottomColor) ?? color);
            }
            if (this.DiagonalStyle != null)
            {
                diagonal = new XElement(S.diagonal,
                    new XAttribute(SSNoNamespace.style, this.DiagonalStyle),
                    CellStyleUtil.CreateColorXElement(this.DiagonalColor) ?? color);
            }
            if (this.DiagonalUp)
            {
                diagonalUp = new XAttribute("diagonalUp", 1);
            }
            if (this.DiagonalDown)
            {
                diagonalDown = new XAttribute("diagonalDown", 1);
            }
            var result = new XElement(S.border,
                diagonalUp, diagonalDown,
                left, right, top, bottom, diagonal);
            return result;
        }
    }

    public class CellStyleFill : CellStyle
    {
        // Example : Color = "FFFFFF00"; // yellow fill (ARGB)
        public string Color;

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
            XElement xcolor = CellStyleUtil.CreateColorXElement(this.Color);
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

    class CellStyleUtil
    {
        public static XElement CreateColorXElement(string color)
        {
            XElement result = null;
            if (color != null)
            {
                result = new XElement(S.color,
                    new XAttribute(SSNoNamespace.rgb, color));
            }
            return result;
        }
    }
}
