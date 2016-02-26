using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public class RowList<T> : IEnumerable<RowDfn>
    {
        public List<T> List;
        public Func<T, RowDfn> ToRowDfn = (o) => { return new RowDfn(); };

        public RowList(List<T> list)
        {
            this.List = list;
        }

        public IEnumerator<RowDfn> GetEnumerator()
        {
            foreach (var o in List)
            {
                yield return ToRowDfn(o);
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new NotImplementedException();
        }
    }

    public class ColDfn
    {
        public decimal? Width;
        public ColAutoFit AutoFit;
    }

    public class ColAutoFit
    {
        public CellDfn Standard;
        public decimal? MinWidth;
        public decimal? MaxWidth;
    }

    public class CellStyleDfn : CellStyle
    {
        public CellAlignment Alignment;
        public CellStyleNumFmt NumFmt;
        public CellStyleBorder Border;
        public CellStyleFill Fill;
        public CellStyleFont Font;
    }

    public class CellAlignment
    {
        public HorizontalCellAlignment? Horizontal;
        public VerticalCellAlignment? Vertical;
        public bool? WrapText;
    }

    public class CellStyle
    {
        //index number of style-elements in styles.xml
        //warning: Do not share CellStyle-instance with another workbook
        public int? id { get; protected internal set; }
    }

    public class CellStyleNumFmt : CellStyle
    {
        public string formatCode;
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
    }

    public class CellStyleFill : CellStyle
    {
        // Example : Color = "FFFFFF00"; // yellow fill (ARGB)
        public string Color;
    }

    public class CellStyleFont : CellStyle
    {
        public uint? Size;
        public string Name;
        public string Color;
        public bool? Bold;
        public bool? Italic;
    }

    public enum CellStyleFontFamilyEnum
    {
        Roman = 1,
        Swiss = 2,
        Modern = 3,
        Script = 4,
        Decorative = 5,
    }

    public class CellCommentDfn
    {
        //public string Author;
        public string CommentText;
        public CellStyleFont Font;
        public Dictionary<string, string> ShapeStyle;
        public AnchorDfn Anchor;
        public int RowIndex; // 0 start
        public int ColIndex; // 0 start
        public string Reference => GetColumnName(ColIndex) + (RowIndex + 1).ToString();

        private string GetColumnName(int index)
        {
            string str = "";
            if (index < 0) return str;
            do
            {
                str = Convert.ToChar(index % 26 + 0x41) + str;
            } while ((index = index / 26 - 1) != -1);
            return str;
        }

        public AnchorDfn CreateAnchor()
        {
            bool isFirstRow = (RowIndex <= 0);
            int top = (isFirstRow) ? 0 : RowIndex - 1;
            int left = ColIndex + 1;
            var result = new AnchorDfn
            {
                LeftColumn = left,
                LeftOffset = 15,
                TopRow = top,
                TopOffset = (isFirstRow) ? 2 : 10,
                RightColumn = left + 2,
                RightOffset = 15,
                BottomRow = left + 3,
                BottomOffset = (isFirstRow) ? 16 : 4
            };
            return result;
        }
    }

    public class AnchorDfn
    {
        public int LeftColumn;
        public int LeftOffset;
        public int TopRow;
        public int TopOffset;
        public int RightColumn;
        public int RightOffset;
        public int BottomRow;
        public int BottomOffset;

        public override string ToString()
        {
            var result = string.Join(",",
                new string[] {
                    LeftColumn.ToString(),
                    LeftOffset.ToString(),
                    TopRow.ToString(),
                    TopOffset.ToString(),
                    RightColumn.ToString(),
                    RightOffset.ToString(),
                    BottomRow.ToString(),
                    BottomOffset.ToString()
                });
            return result;
        }
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

        public static XElement ToXElement(CellAlignment alignment)
        {
            XElement result = null;
            XAttribute ha = null;
            XAttribute va = null;
            XAttribute wt = null;
            if (alignment.Horizontal != null)
            {
                ha = new XAttribute(SSNoNamespace.horizontal, alignment.Horizontal.ToString().ToLower());
            }
            if (alignment.Vertical != null)
            {
                va = new XAttribute(NoNamespace.vertical, alignment.Vertical.ToString().ToLower());
            }
            if (alignment.WrapText != null && alignment.WrapText == true)
            {
                wt = new XAttribute(NoNamespace.wrapText, 1);
            }
            if (ha != null || va != null || wt != null)
            {
                result = new XElement(S.alignment, ha, va, wt);
            }
            return result;
        }

        public static XElement ToXElement(CellStyleNumFmt style)
        {
            var result = new XElement(S.numFmt,
                new XAttribute(SSNoNamespace.numFmtId, style.id),
                new XAttribute(SSNoNamespace.formatCode, style.formatCode));
            return result;
        }

        public static XElement ToXElement(CellStyleBorder style)
        {
            XElement color = CreateColorXElement(style.Color);
            XElement left = null;
            XElement right = null;
            XElement top = null;
            XElement bottom = null;
            XElement diagonal = null;
            XAttribute diagonalUp = null;
            XAttribute diagonalDown = null;
            if (style.LeftStyle != null)
            {
                left = new XElement(S.left,
                    new XAttribute(SSNoNamespace.style, style.LeftStyle),
                    CreateColorXElement(style.LeftColor) ?? color);
            }
            if (style.RightStyle != null)
            {
                right = new XElement(S.right,
                    new XAttribute(SSNoNamespace.style, style.RightStyle),
                    CreateColorXElement(style.RightColor) ?? color);
            }
            if (style.TopStyle != null)
            {
                top = new XElement(S.top,
                    new XAttribute(SSNoNamespace.style, style.TopStyle),
                    CreateColorXElement(style.TopColor) ?? color);
            }
            if (style.BottomStyle != null)
            {
                bottom = new XElement(S.bottom,
                    new XAttribute(SSNoNamespace.style, style.BottomStyle),
                    CreateColorXElement(style.BottomColor) ?? color);
            }
            if (style.DiagonalStyle != null)
            {
                diagonal = new XElement(S.diagonal,
                    new XAttribute(SSNoNamespace.style, style.DiagonalStyle),
                    CreateColorXElement(style.DiagonalColor) ?? color);
            }
            if (style.DiagonalUp)
            {
                diagonalUp = new XAttribute("diagonalUp", 1);
            }
            if (style.DiagonalDown)
            {
                diagonalDown = new XAttribute("diagonalDown", 1);
            }
            var result = new XElement(S.border,
                diagonalUp, diagonalDown,
                left, right, top, bottom, diagonal);
            return result;
        }

        public static XElement ToXElement(CellStyleFill style)
        {
            XElement patternFill = null;
            if (style.Color != null)
            {
                var fgColor = new XElement(S.fgColor,
                                        new XAttribute(SSNoNamespace.rgb, style.Color));
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

        public static XElement ToXElement(CellStyleFont style)
        {
            XElement xsize = null;
            if (style.Size != null)
            {
                xsize = new XElement(S.sz, new XAttribute(SSNoNamespace.val, style.Size));
            }
            XElement xname = null;
            XElement xfamily = null;
            if (style.Name != null)
            {
                xname = new XElement(S.name, new XAttribute(SSNoNamespace.val, style.Name));
                xfamily = new XElement(S.family, new XAttribute(SSNoNamespace.val, (int)CellStyleFontFamilyEnum.Swiss));
            }
            XElement xcolor = CreateColorXElement(style.Color);
            XElement xbold = null;
            if (style.Bold == true)
            {
                xbold = new XElement(S.b);
            }
            XElement xitalic = null;
            if (style.Italic == true)
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

        public static string ToStyleString(Dictionary<string, string> dic)
        {
            var result = string.Join(";", dic.Select(x => x.Key + ":" + x.Value));
            return result;
        }
    }
}
