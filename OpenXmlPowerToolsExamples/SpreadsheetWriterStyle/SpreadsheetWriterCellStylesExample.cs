using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
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

            // cell-styles
            BorderExample(tempDi);
            CellAlignmentExample(tempDi);
            FillExample(tempDi);
            FontExample(tempDi);
            DefaultFontExample(tempDi);
            NumFmtExample(tempDi);

            // Formula
            FormulaExample(tempDi);
            DefinedNameAndDataValidationExample(tempDi);

            // Comment
            CommentExample(tempDi);

            // Table,Filter
            TableAndFilterExample(tempDi);

            // row / column styles
            RowHeightExample(tempDi);
            ColWidthExample(tempDi);
            ColAutoFitExample(tempDi);
            ColAutoFitExample2(tempDi);

            // Write to MemoryStream
            WriteToStreamExample(tempDi);

#if ExecMemorySavingExample
            // memory-saving
            //OutOfMemory(tempDi);
            MemorySaving(tempDi);
#endif
        }

        static void BorderExample(DirectoryInfo dir)
        {
            var headerCellStyle = new CellStyleDfn { Font = new CellStyleFont { Bold = true }, Alignment = new CellAlignment { Horizontal = HorizontalCellAlignment.Center } };
            var boxBorderCellStyle = new CellStyleDfn { Border = CellStyleBorder.CreateBoxBorder(CellStyleBorder.Thin) };
            var redBoxBorderCellStyle = new CellStyleDfn { Border = CellStyleBorder.CreateBoxBorder(CellStyleBorder.Thin, "FFFF0000") };
            var underBorderCellStyle = new CellStyleDfn { Border = new CellStyleBorder { BottomStyle = CellStyleBorder.Thin } };
            var thickUnderBorderCellStyle = new CellStyleDfn { Border = new CellStyleBorder { BottomStyle = CellStyleBorder.Thick } };
            var redUnderBorderCellStyle = new CellStyleDfn { Border = new CellStyleBorder { BottomStyle = CellStyleBorder.Thin, Color = "FFFF0000" } };
            var diagonalUpBorderCellStyle = new CellStyleDfn { Border = new CellStyleBorder { DiagonalUp = true, DiagonalStyle = CellStyleBorder.Thin } };
            var diagonalDownBorderCellStyle = new CellStyleDfn { Border = new CellStyleBorder { DiagonalDown = true, DiagonalStyle = CellStyleBorder.Thin } };
            var colorfulBorderCellStyle = new CellStyleDfn { Border = new CellStyleBorder {
                LeftStyle = CellStyleBorder.Thick,
                LeftColor = "FFFF0000",
                RightStyle = CellStyleBorder.Thick,
                RightColor = "FF00FF00",
                TopStyle = CellStyleBorder.Thick,
                TopColor = "FF0000FF",
                BottomStyle = CellStyleBorder.Thick,
                BottomColor = "FFFFFF00",
                DiagonalStyle = CellStyleBorder.Thick,
                DiagonalColor = "FF00FFFF",
                DiagonalUp = true, DiagonalDown = true
            }};
            var emptyRow = new RowDfn { Cells = new CellDfn[] { new CellDfn() } };
            WorkbookDfn wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    new WorksheetDfn
                    {
                        Name = "Border",
                        ColumnHeadings = new CellDfn[]
                        {
                            new CellDfn
                            {
                                Value = "Caption",
                                Style = headerCellStyle,
                            },
                            new CellDfn
                            {
                                Value = "Border",
                                Style = headerCellStyle,
                            },
                        },
                        Rows = new RowDfn[]
                        {
                            emptyRow,
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "box",
                                    },
                                    new CellDfn {
                                        Style = boxBorderCellStyle,
                                    },
                                },
                            },
                            emptyRow,
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "box (red)",
                                    },
                                    new CellDfn {
                                        Style = redBoxBorderCellStyle,
                                    },
                                },
                            },
                            emptyRow,
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "under",
                                    },
                                    new CellDfn {
                                        Style = underBorderCellStyle,
                                    },
                                },
                            },
                            emptyRow,
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "thick under",
                                    },
                                    new CellDfn {
                                        Style = thickUnderBorderCellStyle,
                                    },
                                },
                            },
                            emptyRow,
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "under (red)",
                                    },
                                    new CellDfn {
                                        Style = redUnderBorderCellStyle,
                                    },
                                },
                            },
                            emptyRow,
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "diagonal up",
                                    },
                                    new CellDfn {
                                        Style = diagonalUpBorderCellStyle,
                                    },
                                },
                            },
                            emptyRow,
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "diagonal down",
                                    },
                                    new CellDfn {
                                        Style = diagonalDownBorderCellStyle,
                                    },
                                },
                            },
                            emptyRow,
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "colorful",
                                    },
                                    new CellDfn {
                                        Style = colorfulBorderCellStyle,
                                    },
                                },
                            },
                        }
                    }
                }
            };
            SpreadsheetWriter.Write(Path.Combine(dir.FullName, "BorderExample.xlsx"), wb);
        }

        static void CellAlignmentExample(DirectoryInfo dir)
        {
            var hcenterCellStyle = new CellStyleDfn { Alignment = new CellAlignment { Horizontal = HorizontalCellAlignment.Center } };
            var leftCellStyle = new CellStyleDfn { Alignment = new CellAlignment { Horizontal = HorizontalCellAlignment.Left } };
            var rightCellStyle = new CellStyleDfn { Alignment = new CellAlignment { Horizontal = HorizontalCellAlignment.Right } };
            var topCellStyle = new CellStyleDfn { Alignment = new CellAlignment { Vertical = VerticalCellAlignment.Top } };
            var vcenterCellStyle = new CellStyleDfn { Alignment = new CellAlignment { Vertical = VerticalCellAlignment.Center } };
            var bottomCellStyle = new CellStyleDfn { Alignment = new CellAlignment { Vertical = VerticalCellAlignment.Bottom } };
            var centerCellStyle = new CellStyleDfn { Alignment = new CellAlignment
            {
                Horizontal = HorizontalCellAlignment.Center,
                Vertical = VerticalCellAlignment.Center
            }};
            var wrapTextCellStyle = new CellStyleDfn { Alignment = new CellAlignment { WrapText = true } };
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
                                        Value = "horizontal center",
                                        Style = hcenterCellStyle,
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
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "top",
                                        Style = topCellStyle,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "vertical center",
                                        Style = vcenterCellStyle,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "bottom",
                                        Style = bottomCellStyle,
                                    },
                                }
                            },
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
                                        Value = "wrap text\nnew line",
                                        Style = wrapTextCellStyle,
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
            var headerCellStyle = new CellStyleDfn {
                Font = new CellStyleFont { Bold = true },
                Alignment = new CellAlignment { Horizontal = HorizontalCellAlignment.Center },
            };
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
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "size 24",
                                        Style = new CellStyleDfn { Font = new CellStyleFont { Size = 24} },
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "red",
                                        Style = new CellStyleDfn { Font = new CellStyleFont { Color = "FFFF0000"} },
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "font: Arial",
                                        Style = new CellStyleDfn { Font = new CellStyleFont { Name = "Arial"} },
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "font: Times New Roman",
                                        Style = new CellStyleDfn { Font = new CellStyleFont { Name = "Times New Roman"} },
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "font: Courier New",
                                        Style = new CellStyleDfn { Font = new CellStyleFont { Name = "Courier New"} },
                                    },
                                }
                            },
                        }
                    }
                }
            };
            SpreadsheetWriter.Write(Path.Combine(dir.FullName, "FontExample.xlsx"), wb);
        }

        static void DefaultFontExample(DirectoryInfo dir)
        {
            WorkbookDfn wb = new WorkbookDfn
            {
                DefaultFont = new CellStyleFont { Name = "Arial", Size = 14},
                Worksheets = new WorksheetDfn[]
                {
                    new WorksheetDfn
                    {
                        Name = "DefaultFont",
                        Rows = new RowDfn[]
                        {
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "default font",
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "font: Times New Roman",
                                        Style = new CellStyleDfn { Font = new CellStyleFont { Name = "Times New Roman"} },
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "font: Courier New",
                                        Style = new CellStyleDfn { Font = new CellStyleFont { Name = "Courier New"} },
                                    },
                                }
                            },
                        }
                    }
                }
            };
            SpreadsheetWriter.Write(Path.Combine(dir.FullName, "DefaultFontExample.xlsx"), wb);
        }

        static void NumFmtExample(DirectoryInfo dir)
        {
            var headerCellStyle = new CellStyleDfn
            {
                Font = new CellStyleFont { Bold = true },
                Alignment = new CellAlignment { Horizontal = HorizontalCellAlignment.Center },
            };
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

        static void FormulaExample(DirectoryInfo dir)
        {
            WorkbookDfn wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    new WorksheetDfn
                    {
                        Name = "Formula",
                        Rows = new RowDfn[]
                        {
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = 1,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = 2,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = 3,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Formula = "SUM(A1:A3)",
                                    },
                                }
                            },
                        }
                    }
                }
            };
            SpreadsheetWriter.Write(Path.Combine(dir.FullName, "FormulaExample.xlsx"), wb);
        }

        static void DefinedNameAndDataValidationExample(DirectoryInfo dir)
        {
            WorkbookDfn wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    new WorksheetDfn
                    {
                        Name = "DataValidation",
                        Rows = new RowDfn[]
                        {
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "whole",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "decimal",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "list",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "date",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "time",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "textLength",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "custom",
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "",
                                    },
                                }
                            },
                        },
                        DataValidations = new DataValidationDfn[] {
                            new DataValidationDfn
                            {
                                ReferenceSequence = "A2",
                                Type = DataValidationType.whole,
                                AllowBlank = true,
                                Operator = DataValidationOperator.between,
                                Formula1 = "1",
                                Formula2 = "100",
                                PromptTitle = "Type: whole",
                                PromptMessage = "Please enter a number between 1 and 100",
                            },
                            new DataValidationDfn
                            {
                                ReferenceSequence = "B2",
                                Type = DataValidationType.@decimal,
                                AllowBlank = true,
                                Operator = DataValidationOperator.greaterThanOrEqual,
                                Formula1 = "0",
                                PromptTitle = "Type: decimal",
                                PromptMessage = "Please enter a number of 0 or more",
                            },
                            new DataValidationDfn
                            {
                                ReferenceSequence = "C2",
                                Type = DataValidationType.list,
                                AllowBlank = true,
                                Formula1 = "NameOfRange1",
                            },
                            new DataValidationDfn
                            {
                                ReferenceSequence = "D2",
                                Type = DataValidationType.date,
                                AllowBlank = true,
                                Operator = DataValidationOperator.greaterThan,
                                Formula1 = "1",
                                PromptTitle = "Type: date",
                                PromptMessage = "Please enter a date",
                            },
                            new DataValidationDfn
                            {
                                ReferenceSequence = "E2",
                                Type = DataValidationType.time,
                                AllowBlank = true,
                                Operator = DataValidationOperator.greaterThan,
                                Formula1 = "0",
                                PromptTitle = "Type: time",
                                PromptMessage = "Please enter a time",
                            },
                            new DataValidationDfn
                            {
                                ReferenceSequence = "F2",
                                Type = DataValidationType.textLength,
                                AllowBlank = true,
                                Operator = DataValidationOperator.lessThan,
                                Formula1 = "26",
                                PromptTitle = "Type: textLength",
                                PromptMessage = "Please enter your name (max 25 characters)",
                            },
                            new DataValidationDfn
                            {
                                ReferenceSequence = "G2",
                                Type = DataValidationType.custom,
                                AllowBlank = true,
                                Formula1 = "=LEN(G2)=LENB(G2)",
                                PromptTitle = "Type: custom",
                                PromptMessage = "Please enter your name (one-byte character only)",
                            },
                        },
                    },
                    new WorksheetDfn
                    {
                        Name = "DefinedNameSheet",
                        Rows = new RowDfn[]
                        {
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "Name of range 1",
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "AAAAA",
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "BBBBB",
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "CCCCC",
                                    },
                                }
                            },
                        }
                    },
                },
                DefinedNames = new DefinedNameDfn[] {
                    new DefinedNameDfn { Name = "NameOfRange1", Text = "DefinedNameSheet!$A$2:$A$4" },
                },
            };
            SpreadsheetWriter.Write(Path.Combine(dir.FullName, "DefinedNameAndDataValidationExample.xlsx"), wb);
        }

        static void CommentExample(DirectoryInfo dir)
        {
            WorkbookDfn wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    new WorksheetDfn
                    {
                        Name = "Comment",
                        Rows = new RowDfn[]
                        {
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "comment ->",
                                    },
                                },
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "custom comment-text style ->",
                                    },
                                },
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "Excel style ->",
                                    },
                                },
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "custom shape style ->",
                                    },
                                },
                            },
                        },
                        Comments = new CellCommentDfn[]
                        {
                            new CellCommentDfn
                            {
                                CommentText = "Hello!\nThanks!",
                                RowIndex = 0,
                                ColIndex = 1
                            },
                            new CellCommentDfn
                            {
                                CommentText = new CommentTextDfn(new object[] {
                                    //Set alternately: CellStyleFont, comment-string, CellStyleFont, comment-string,...
                                    new CellStyleFont
                                    {
                                        Bold = true,
                                        Italic = true,
                                        Size = 16,
                                        Color = "FFFF0000", //red
                                        Name = "Tahoma",
                                    },
                                    "Custom comment-text style",
                                }),
                                RowIndex = 1,
                                ColIndex = 1
                            },
                            new CellCommentDfn
                            {
                                CommentText = new CommentTextDfn(new object[] {
                                    new CellStyleFont
                                    {
                                        Bold = true,
                                    },
                                    "Author:",
                                    new CellStyleFont(),
                                    "\ncomment...",
                                }),
                                RowIndex = 2,
                                ColIndex = 1
                            },
                            new CellCommentDfn
                            {
                                CommentText = "Custom shape style:\nThe quick brown fox jumps over the lazy dog",
                                ShapeStyle = new Dictionary<string,string>() {
                                    { "position", "absolute" },
                                    { "margin-left", "151.5pt" },
                                    { "margin-top", "21.5pt" },
                                    { "width", "211pt" },
                                    { "height", "31pt" },
                                    { "z-index", "4" },
                                    { "visibility", "visible" },
                                },
                                Anchor = new AnchorDfn {
                                    LeftColumn = 3,
                                    LeftOffset = 15,
                                    TopRow = 1,
                                    TopOffset = 14,
                                    RightColumn = 7,
                                    RightOffset = 53,
                                    BottomRow = 3,
                                    BottomOffset = 18
                                },
                                RowIndex = 3,
                                ColIndex = 1
                            },
                        },
                    }
                }
            };
            SpreadsheetWriter.Write(Path.Combine(dir.FullName, "CommentExample.xlsx"), wb);
        }

        static void TableAndFilterExample(DirectoryInfo dir)
        {
            WorkbookDfn wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    new WorksheetDfn
                    {
                        Name = "TableAndFilter",
                        TableName = "Presidents",
                        ColumnHeadings = new CellDfn[]
                        {
                            new CellDfn
                            {
                                Value = "Name",
                            },
                            new CellDfn
                            {
                                Value = "Number",
                            },
                        },
                        FilterColumns = new FilterColumnDfn[]
                        {
                            new FilterColumnDfn
                            {
                                ColId = 0,
                                //NOTICE:You must do this filter processing
                                Filters = new[] { "Bush", "Obama" },
                            },
                        },
                        Rows = new RowDfn[]
                        {
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        Value = "Bush",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = "51",
                                    },
                                },
                            },
                            new RowDfn
                            {
                                //hide by filter results
                                Hidden = true,
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        Value = "Clinton",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = "52",
                                    },
                                },
                            },
                            new RowDfn
                            {
                                //hide by filter results
                                Hidden = true,
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        Value = "Clinton",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = "53",
                                    },
                                },
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        Value = "Bush",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = "54",
                                    },
                                },
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        Value = "Bush",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = "55",
                                    },
                                },
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        Value = "Obama",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = "56",
                                    },
                                },
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        Value = "Obama",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = "57",
                                    },
                                },
                            },
                        },
                    }
                }
            };
            SpreadsheetWriter.Write(Path.Combine(dir.FullName, "TableAndFilterExample.xlsx"), wb);
        }

        static void RowHeightExample(DirectoryInfo dir)
        {
            WorkbookDfn wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    new WorksheetDfn
                    {
                        Name = "RowHeight",
                        Rows = new RowDfn[]
                        {
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "Default height",
                                    },
                                },
                            },
                            new RowDfn
                            {
                                Height = (decimal)33.3,
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "height: 33.3",
                                    },
                                },
                            },
                        }
                    }
                }
            };
            SpreadsheetWriter.Write(Path.Combine(dir.FullName, "RowHeightExample.xlsx"), wb);
        }

        static void ColWidthExample(DirectoryInfo dir)
        {
            WorkbookDfn wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    new WorksheetDfn
                    {
                        Name = "ColWidth",
                        Cols = new ColDfn[]
                        {
                            null,
                            new ColDfn { Width = (decimal)24.68 }
                        },
                        Rows = new RowDfn[]
                        {
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "Default width",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "width: 24.68",
                                    },
                                },
                            },
                        }
                    }
                }
            };
            SpreadsheetWriter.Write(Path.Combine(dir.FullName, "ColWidthExample.xlsx"), wb);
        }

        static void ColAutoFitExample(DirectoryInfo dir)
        {
            var wrapTextCellStyle = new CellStyleDfn { Alignment = new CellAlignment { WrapText =true } };
            WorkbookDfn wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    new WorksheetDfn
                    {
                        Name = "ColAutoFit",
                        Cols = new ColDfn[]
                        {
                            new ColDfn { AutoFit = new ColAutoFit() },
                            new ColDfn { AutoFit = new ColAutoFit() },
                            new ColDfn { AutoFit = new ColAutoFit { MinWidth = 15 } },
                            new ColDfn { AutoFit = new ColAutoFit { MaxWidth = 12 } },
                        },
                        Rows = new RowDfn[]
                        {
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "ABC",
                                        Style = new CellStyleDfn { Font = new CellStyleFont { Size = 11 } }
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "123,456,789",
                                        Style = new CellStyleDfn { Font = new CellStyleFont { Size = 24 } }
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "MinWidth",
                                        Style = wrapTextCellStyle,
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "MaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaxWidth",
                                        Style = wrapTextCellStyle,
                                    },
                                },
                            },
                        }
                    }
                }
            };
            SpreadsheetWriter.Write(Path.Combine(dir.FullName, "ColAutoFitExample.xlsx"), wb);
        }

        static void ColAutoFitExample2(DirectoryInfo dir)
        {
            var wrapTextCellStyle = new CellStyleDfn { Alignment = new CellAlignment { WrapText = true } };
            WorkbookDfn wb = new WorkbookDfn
            {
                DefaultFont = new CellStyleFont { Name = "ＭＳ ゴシック", Size = 11 },
                MeasureFallbackFontName = "VL ゴシック",
                MeasureBaseSize = 7,//DefaultFont's character-width
                Worksheets = new WorksheetDfn[]
                {
                    new WorksheetDfn
                    {
                        Name = "ColAutoFit",
                        Cols = new ColDfn[]
                        {
                            new ColDfn { AutoFit = new ColAutoFit() },
                            new ColDfn { AutoFit = new ColAutoFit() },
                            //If ColumnHeadings is less than Cols
                            new ColDfn { AutoFit = new ColAutoFit() },
                            //fit width to Standard.Value
                            new ColDfn { AutoFit = new ColAutoFit { Standard = new CellDfn { Value = "0,000,000" } } },
                            //fit width to cell.Value
                            new ColDfn { AutoFit = new ColAutoFit() },
                        },
                        ColumnHeadings = new CellDfn[]
                        {
                            new CellDfn
                            {
                                Value = "いろはにほへと",
                                Style = new CellStyleDfn { Font = new CellStyleFont { Name = "ＭＳ ゴシック", Size = 13 } }
                            },
                            new CellDfn
                            {
                                Value = "two",
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
                                        Value = "one",
                                    },
                                },
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "one",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "two",
                                    },
                                },
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "one",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "two",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "threeeeeeeeeeeeeeeeeeeeee",
                                    },
                                },
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        //if CellDataType is CellDataType.Date
                                        //   and Value as DateTime
                                        //then AutoFit-width size same as "mm-dd-yyyy" (using ToShortDateString())
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2015, 12, 5),
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "mm-dd-yyyy" } },
                                    },
                                },
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {},
                                    new CellDfn {},
                                    new CellDfn {},
                                    //fit width to Standard.Value of ColDfn
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = 1234567,
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "#,##0" } },
                                    },
                                    //fit width to cell.Value
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = 1234567,
                                        Style = new CellStyleDfn { NumFmt = new CellStyleNumFmt { formatCode = "#,##0" } },
                                    },
                                },
                            },
                        }
                    }
                }
            };
            SpreadsheetWriter.Write(Path.Combine(dir.FullName, "ColAutoFitExample2.xlsx"), wb);
        }

        static void WriteToStreamExample(DirectoryInfo dir)
        {
            WorkbookDfn wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    new WorksheetDfn
                    {
                        Name = "WriteToStream",
                        Rows = new RowDfn[]
                        {
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "Hello",
                                    },
                                },
                            },
                        }
                    }
                }
            };
            MemoryStream stream;
            SpreadsheetWriter.Write(out stream, wb);
            File.WriteAllBytes(Path.Combine(dir.FullName, "WriteToStreamExample.xlsx"), stream.ToArray());
        }

        static void OutOfMemory(DirectoryInfo dir)
        {
            var rows = new List<RowDfn>();
            for (var r = 0; r < 30000; r++)
            {
                var row = new RowDfn();
                var cells = new List<CellDfn>();
                for (var c = 0; c < 1024; c++)
                {
                    var cell = new CellDfn {
                        CellDataType = CellDataType.String,
                        Value = string.Format("{0}-{1}", r + 1, c + 1),
                    };
                    cells.Add(cell);
                }
                row.Cells = cells.ToArray();
                rows.Add(row);
            }
            WorkbookDfn wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    new WorksheetDfn
                    {
                        Name = "OutOfMemory",
                        Rows = rows.ToArray(),
                    }
                }
            };
            SpreadsheetWriter.Write(Path.Combine(dir.FullName, "OutOfMemory.xlsx"), wb);
        }

        static void MemorySaving(DirectoryInfo dir)
        {
            var list = new List<int>();
            for (var r = 0; r < 30000; r++)
            {
                list.Add(r + 1);
            }
            var rows = new RowList<int>(list);
            rows.ToRowDfn = (o) => {
                var row = new RowDfn();
                var cells = new List<CellDfn>();
                for (var c = 0; c < 1024; c++)
                {
                    var cell = new CellDfn
                    {
                        CellDataType = CellDataType.String,
                        Value = string.Format("{0}-{1}", Convert.ToInt32(o), c + 1),
                    };
                    cells.Add(cell);
                }
                row.Cells = cells.ToArray();
                return row;
            };
            WorkbookDfn wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    new WorksheetDfn
                    {
                        Name = "MemorySaving",
                        Rows = rows,
                    }
                }
            };
            SpreadsheetWriter.Write(Path.Combine(dir.FullName, "MemorySaving.xlsx"), wb);
        }
    }
}
