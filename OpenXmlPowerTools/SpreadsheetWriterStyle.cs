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
        //TODO
    }

    public class CellStyleFont : CellStyle
    {
        public uint? Size;
        public string Name;
        public bool? Bold;
        public bool? Italic;

        public XElement ToXElement()
        {
            //TODO Size
            //TODO Name
            var result = new XElement(S.font,
                this.Bold == true ? new XElement(S.b) : null,
                            this.Italic == true ? new XElement(S.i) : null);
            return result;
        }
    }

}
