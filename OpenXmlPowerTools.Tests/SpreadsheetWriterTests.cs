/***************************************************************************

Copyright (c) Microsoft Corporation 2012-2015.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

Developer: Eric White
Blog: http://www.ericwhite.com
Twitter: @EricWhiteDev
Email: eric@ericwhite.com

***************************************************************************/

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Sw = OpenXmlPowerTools;
using Xunit;

namespace OxPt
{
    public class SwTests
    {
        [Fact]
        public void SW001_Simple()
        {
            var boldFont = new Sw.CellStyleFont { Bold = true };
            var boldCellStyle = new Sw.CellStyleDfn { Font = boldFont };
            var boldAndLeftCellStyle = new Sw.CellStyleDfn { Font = boldFont, Alignment = new Sw.CellAlignment { Horizontal = Sw.HorizontalCellAlignment.Left } };
            var format0_00 = new Sw.CellStyleNumFmt { formatCode = "0.00" };
            var numCellStyle = new Sw.CellStyleDfn { NumFmt = format0_00 };
            Sw.WorkbookDfn wb = new Sw.WorkbookDfn
            {
                Worksheets = new Sw.WorksheetDfn[]
                {
                    new Sw.WorksheetDfn
                    {
                        Name = "MyFirstSheet",
                        TableName = "NamesAndRates",
                        ColumnHeadings = new Sw.CellDfn[]
                        {
                            new Sw.CellDfn
                            {
                                Value = "Name",
                                Style = boldCellStyle,
                            },
                            new Sw.CellDfn
                            {
                                Value = "Age",
                                Style = boldAndLeftCellStyle,
                            },
                            new Sw.CellDfn
                            {
                                Value = "Rate",
                                Style = boldAndLeftCellStyle,
                            }
                        },
                        Rows = new Sw.RowDfn[]
                        {
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "Eric",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = 50,
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = (decimal)45.00,
                                        Style = numCellStyle,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "Bob",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = 42,
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = (decimal)78.00,
                                        Style = numCellStyle,
                                    },
                                }
                            },
                        }
                    }
                }
            };
            var outXlsx = new FileInfo(Path.Combine(Sw.TestUtil.TempDir.FullName, "SW001-Simple.xlsx"));
            Sw.SpreadsheetWriter.Write(outXlsx.FullName, wb);
            Validate(outXlsx);
        }

        [Fact]
        public void SW002_AllDataTypes()
        {
            var boldFont = new Sw.CellStyleFont { Bold = true };
            var boldCellStyle = new Sw.CellStyleDfn { Font = boldFont };
            var boldAndRightCellStyle = new Sw.CellStyleDfn { Font = boldFont, Alignment = new Sw.CellAlignment { Horizontal = Sw.HorizontalCellAlignment.Right } };
            var rightCellStyle = new Sw.CellStyleDfn { Alignment = new Sw.CellAlignment { Horizontal = Sw.HorizontalCellAlignment.Right } };
            var formatMMDDYY = new Sw.CellStyleNumFmt { formatCode = "mm-dd-yy" };
            var dateCellStyle = new Sw.CellStyleDfn { NumFmt = formatMMDDYY };
            var dateAndBoldCellStyle = new Sw.CellStyleDfn { Font = boldFont, NumFmt = formatMMDDYY, Alignment = new Sw.CellAlignment { Horizontal = Sw.HorizontalCellAlignment.Center } };
            Sw.WorkbookDfn wb = new Sw.WorkbookDfn
            {
                Worksheets = new Sw.WorksheetDfn[]
                {
                    new Sw.WorksheetDfn
                    {
                        Name = "MyFirstSheet",
                        ColumnHeadings = new Sw.CellDfn[]
                        {
                            new Sw.CellDfn
                            {
                                Value = "DataType",
                                Style = boldCellStyle,
                            },
                            new Sw.CellDfn
                            {
                                Value = "Value",
                                Style = boldAndRightCellStyle,
                            },
                        },
                        Rows = new Sw.RowDfn[]
                        {
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "Boolean",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Boolean,
                                        Value = true,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "Boolean",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Boolean,
                                        Value = false,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "String",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "A String",
                                        Style = rightCellStyle,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "int",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = (int)100,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "int?",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = (int?)100,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "int? (is null)",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = null,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "uint",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = (uint)101,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "long",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = Int64.MaxValue,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "float",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = (float)123.45,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "double",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = (double)123.45,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "decimal",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = (decimal)123.45,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8),
                                        Style = dateCellStyle,
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Date,
                                        Value = new DateTime(2012, 1, 9),
                                        Style = dateAndBoldCellStyle,
                                    },
                                }
                            },
                        }
                    }
                }
            };
            var outXlsx = new FileInfo(Path.Combine(Sw.TestUtil.TempDir.FullName, "SW002-DataTypes.xlsx"));
            Sw.SpreadsheetWriter.Write(outXlsx.FullName, wb);
            Validate(outXlsx);
        }

        private void Validate(FileInfo fi)
        {
            using (SpreadsheetDocument sDoc = SpreadsheetDocument.Open(fi.FullName, true))
            {
                OpenXmlValidator v = new OpenXmlValidator();
                var errors = v.Validate(sDoc).Where(ve => !s_ExpectedErrors.Contains(ve.Description));

#if false
                // if a test fails validation post-processing, then can use this code to determine the SDK
                // validation error(s).

                if (errors.Count() != 0)
                {
                    StringBuilder sb = new StringBuilder();
                    foreach (var item in errors)
                    {
                        sb.Append(item.Description).Append(Environment.NewLine);
                    }
                    var s = sb.ToString();
                    Console.WriteLine(s);
                }
#endif

                Assert.Equal(0, errors.Count());
            }
        }

        private static List<string> s_ExpectedErrors = new List<string>()
        {
            "The attribute 't' has invalid value 'd'. The Enumeration constraint failed.",
        };
    }
}
