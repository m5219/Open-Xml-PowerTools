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
        public CommentTextDfn CommentText;
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

    public class CommentTextDfn
    {
        private List<object> list;
        private string _text;

        public CommentTextDfn(string s)
        {
            _text = s;
        }

        public CommentTextDfn(object[] ar)
        {
            _append(ar);
        }

        public CommentTextDfn(IEnumerable<object> ar)
        {
            _append(ar);
        }

        public static implicit operator CommentTextDfn(string s) => new CommentTextDfn(s);

        public string Text {
            get
            {
                if (list == null) return _text;
                string result = string.Join("", list.Where(x => x is string).Select(x => x as string));
                return result;
            }
            set
            {
                _text = value;
                list = null;
            }
        }

        private CommentTextDfn _append(object o)
        {
            if (list == null) list = new List<object>();
            list.Add(o);
            return this;
        }

        private CommentTextDfn _append(IEnumerable<object> ar)
        {
            if (ar != null)
            {
                foreach (var o in ar)
                {
                    _append(o);
                }
            }
            return this;
        }

        public CommentTextDfn Append(CellStyleFont f)
        {
            return _append(f);
        }

        public CommentTextDfn Append(string s)
        {
            return _append(s);
        }

        internal IEnumerable<object> GetEnumerable()
        {
            if (list == null)
            {
                yield return _text ?? "";
            }
            else
            {
                foreach (var o in list) yield return o;
            }
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

        /// <summary>
        /// if ct.GetEnumerable() return { string1, CellStyleFont1, string2, string3, CellStyleFont2, CellStyleFont3, string4 }
        /// then CommentText inclue
        ///     Run( Text(string1) )
        ///     Run( RunProperties(CellStyleFont1), Text(string2) ))
        ///     Run( Text(string3) )
        ///     //not inclue CellStyleFont2
        ///     Run( RunProperties(CellStyleFont3), Text(string4) ))
        /// </summary>
        /// <param name="ct"></param>
        /// <returns></returns>
        public static DocumentFormat.OpenXml.Spreadsheet.CommentText ToCommentText(CommentTextDfn ct)
        {
            var commentText = new DocumentFormat.OpenXml.Spreadsheet.CommentText();
            if (ct != null)
            {
                bool isPreserve = false;
                CellStyleFont style = null;
                string text = null;
                Action<CellStyleFont, string> appendRun = (s,t) => {
                    if (t == null) return;
                    var run = CellStyleUtil.ToRun(s, t, isPreserve);
                    commentText.Append(run);
                    isPreserve = true;
                    style = null;
                    text = null;
                };
                foreach (var o in ct.GetEnumerable())
                {
                    if (o is CellStyleFont)
                    {
                        if (style != null) appendRun(style, text);
                        style = o as CellStyleFont;
                    }
                    else if (o is string)
                    {
                        if (text != null) appendRun(style, text);
                        text = o as string;
                        appendRun(style, text);
                    }
                }
                appendRun(style, text);
            }

            return commentText;
        }

        public static DocumentFormat.OpenXml.Spreadsheet.Run ToRun(CellStyleFont style, string text, bool isPreserve)
        {
            var run = new DocumentFormat.OpenXml.Spreadsheet.Run();
            if (style != null)
            {
                var runProperties = new DocumentFormat.OpenXml.Spreadsheet.RunProperties();
                if (style.Bold == true)
                {
                    runProperties.Append(new DocumentFormat.OpenXml.Spreadsheet.Bold());
                }
                if (style.Italic == true)
                {
                    runProperties.Append(new DocumentFormat.OpenXml.Spreadsheet.Italic());
                }
                if (style.Size != null)
                {
                    runProperties.Append(new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = style.Size });
                }
                if (!string.IsNullOrEmpty(style.Color))
                {
                    runProperties.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = style.Color });
                }
                if (!string.IsNullOrEmpty(style.Name))
                {
                    runProperties.Append(new DocumentFormat.OpenXml.Spreadsheet.RunFont { Val = style.Name });
                }
                runProperties.Append(new DocumentFormat.OpenXml.Spreadsheet.RunPropertyCharSet { Val = 1 });
                run.Append(runProperties);
            }

            var t = new DocumentFormat.OpenXml.Spreadsheet.Text(text ?? "");
            if (isPreserve)
            {
                t.Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
            }
            run.Append(t);

            return run;
        }
    }
}
