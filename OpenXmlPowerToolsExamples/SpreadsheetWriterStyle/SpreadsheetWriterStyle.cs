using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace SpreadsheetWriterExample
{
    class Program
    {
        static void Main(string[] args)
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            //BorderExample(tempDi);
            CellAlignmentExample(tempDi);
            FillExample(tempDi);
            FontExample(tempDi);
            NumFmtExample(tempDi);
        }

        static void CellAlignmentExample(DirectoryInfo dir)
        {
            var centerCellStyle = new CellStyleDfn { HorizontalCellAlignment = HorizontalCellAlignment.Center };
            var leftCellStyle = new CellStyleDfn { HorizontalCellAlignment = HorizontalCellAlignment.Left };
            var rightCellStyle = new CellStyleDfn { HorizontalCellAlignment = HorizontalCellAlignment.Right };
            WorkbookDfn wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    new WorksheetDfn
                    {
                        Name = "CellAlignment",
                        Rows = new RowDfn[]
                        {
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "center",
                                        Style = centerCellStyle,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "left",
                                        Style = leftCellStyle,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "right",
                                        Style = rightCellStyle,
                                    },
                                }
                            },
                        }
                    }
                }
            };
            SpreadsheetWriter.Write(Path.Combine(dir.FullName, "CellAlignmentExample.xlsx"), wb);
        }

        static void FillExample(DirectoryInfo dir)
        {
            var headerCellStyle = new CellStyleDfn { Font = new CellStyleFont { Bold = true }, HorizontalCellAlignment = HorizontalCellAlignment.Center };
            var yellowFillCellStyle = new CellStyleDfn { Fill = new CellStyleFill { Color = "FFFFFF00" } };
            WorkbookDfn wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    new WorksheetDfn
                    {
                        Name = "Fill",
                        ColumnHeadings = new CellDfn[]
                        {
                            new CellDfn
                            {
                                Value = "Color",
                                Style = headerCellStyle,
                            },
                            new CellDfn
                            {
                                Value = "Fill",
                                Style = headerCellStyle,
                            },
                        },
                        Rows = new RowDfn[]
                        {
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "FFFFFF00",
                                    },
                                    new CellDfn {
                                        Style = yellowFillCellStyle,
                                    },
                                }
                            },
                        }
                    }
                }
            };
            SpreadsheetWriter.Write(Path.Combine(dir.FullName, "FillExample.xlsx"), wb);
        }

        static void FontExample(DirectoryInfo dir)
        {
            var boldCellStyle = new CellStyleDfn { Font = new CellStyleFont { Bold = true } };
            var italicCellStyle = new CellStyleDfn { Font = new CellStyleFont { Italic = true } };
            var boldAndItalicCellStyle = new CellStyleDfn { Font = new CellStyleFont { Bold = true, Italic = true } };
            WorkbookDfn wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    new WorksheetDfn
                    {
                        Name = "Font",
                        Rows = new RowDfn[]
                        {
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "Bold",
                                        Style = boldCellStyle,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "Italic",
                                        Style = italicCellStyle,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "Bold & Italic",
                                        Style = boldAndItalicCellStyle,
                                    },
                                }
                            },
                        }
                    }
                }
            };
            SpreadsheetWriter.Write(Path.Combine(dir.FullName, "FontExample.xlsx"), wb);
        }

        static void NumFmtExample(DirectoryInfo dir)
        {
            var headerCellStyle = new CellStyleDfn { Font = new CellStyleFont { Bold = true}, HorizontalCellAlignment = HorizontalCellAlignment.Center };
            WorkbookDfn wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    new WorksheetDfn
                    {
                        Name = "numFmt",
                        ColumnHeadings = new CellDfn[]
                        {
                            new CellDfn
                            {
                                Value = "numFmt",
                                Style = headerCellStyle,
                            },
                            new CellDfn
                            {
                                Value = "Value",
                                Style = headerCellStyle,
                            },
                            new CellDfn
                            {
                                Value = "formatted",
                                Style = headerCellStyle,
                            }
                        },
                        Rows = new RowDfn[]
                        {
                            //Standard format : 1
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "0",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)1234.56,
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)1234.56,
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "0" } },
                                    },
                                }
                            },
                            //Standard format : 2
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "0.00",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)1234.56,
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)1234.56,
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "0.00" } },
                                    },
                                }
                            },
                            //Standard format : 3
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "#,##0",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)1234.56,
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)1234.56,
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "#,##0" } },
                                    },
                                }
                            },
                            //Standard format : 4
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "#,##0.00",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)1234.56,
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)1234.56,
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "#,##0.00" } },
                                    },
                                }
                            },
                            //Standard format : 9
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "0%",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)1234.56,
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)1234.56,
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "0%" } },
                                    },
                                }
                            },
                            //Standard format : 10
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "0.00%",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)1234.56,
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)1234.56,
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "0.00%" } },
                                    },
                                }
                            },
                            //Standard format : 11
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "0.00E+00",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)1234.56,
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)1234.56,
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "0.00E+00" } },
                                    },
                                }
                            },
                            //Standard format : 12
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "# ?/?",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)5.25,
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)5.25,
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "# ?/?" } },
                                    },
                                }
                            },
                            //Standard format : 13
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "# ??/??",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)5.25,
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)5.25,
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "# ??/??" } },
                                    },
                                }
                            },
                            //Standard format : 14
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "mm-dd-yy",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5).ToString(),
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5),
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "mm-dd-yy" } },
                                    },
                                }
                            },
                            //Standard format : 15
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "d-mmm-yy",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5).ToString(),
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5),
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "d-mmm-yy" } },
                                    },
                                }
                            },
                            //Standard format : 16
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "d-mmm",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5).ToString(),
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5),
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "d-mmm" } },
                                    },
                                }
                            },
                            //Standard format : 17
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "mmm-yy",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5).ToString(),
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5),
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "mmm-yy" } },
                                    },
                                }
                            },
                            //Standard format : 18
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "h:mm AM/PM",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5).ToString(),
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5),
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "h:mm AM/PM" } },
                                    },
                                }
                            },
                            //Standard format : 19
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "h:mm:ss AM/PM",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5).ToString(),
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5),
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "h:mm:ss AM/PM" } },
                                    },
                                }
                            },
                            //Standard format : 20
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "h:mm",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5).ToString(),
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5),
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "h:mm" } },
                                    },
                                }
                            },
                            //Standard format : 21
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "h:mm:ss",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5).ToString(),
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5),
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "h:mm:ss" } },
                                    },
                                }
                            },
                            //Standard format : 22
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "h/d/yy h:mm",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5).ToString(),
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5),
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "h/d/yy h:mm" } },
                                    },
                                }
                            },
                            //Standard format : 37
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "#,##0;(#,##0)",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)-1234.56,
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)-1234.56,
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "#,##0;(#,##0)" } },
                                    },
                                }
                            },
                            //Standard format : 38
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "#,##0;[Red](#,##0)",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)-1234.56,
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)-1234.56,
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "#,##0;[Red](#,##0)" } },
                                    },
                                }
                            },
                            //Standard format : 39
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "#,##0.00;(#,##0.00)",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)-1234.56,
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)-1234.56,
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "#,##0.00;(#,##0.00)" } },
                                    },
                                }
                            },
                            //Standard format : 40
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "#,##0.00;[Red](#,##0.00)",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)-1234.56,
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)-1234.56,
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "#,##0.00;[Red](#,##0.00)" } },
                                    },
                                }
                            },
                            //Standard format : 45
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "mm:ss",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5).ToString(),
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8, 13, 4, 5),
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "mm:ss" } },
                                    },
                                }
                            },
                            //Standard format : 46
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "[h]:mm:ss",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = new DateTime(2012, 1, 8, 0, 4, 5).ToString(),
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8, 0, 4, 5),
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "[h]:mm:ss" } },
                                    },
                                }
                            },
                            //Standard format : 47
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "mmss.0",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = new DateTime(2012, 1, 8, 0, 4, 5).ToString(),
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8, 0, 4, 5),
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "mmss.0" } },
                                    },
                                }
                            },
                            //Standard format : 48
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "##0.0E+0",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)1234.56,
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)1234.56,
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "##0.0E+0" } },
                                    },
                                }
                            },
                            //Standard format : 49
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "@",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "0000",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "0000",
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "@" } },
                                    },
                                }
                            },
                        }
                    }
                }
            };
            SpreadsheetWriter.Write(Path.Combine(dir.FullName, "NumFmtExample.xlsx"), wb);
        }
    }
}
