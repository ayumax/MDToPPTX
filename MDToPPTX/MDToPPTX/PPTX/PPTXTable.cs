using System;
using System.Collections.Generic;
using System.Text;

namespace MDToPPTX.PPTX
{
    public enum PPTXTableColumnAlign
    {
        Left = 0,
        Center = 1,
        Right = 2
    }

    public class PPTXTableColumn
    {
        public float Width { get; set; } = 0;
        public PPTXTableColumnAlign Alignment { get; set; } = PPTXTableColumnAlign.Center;
    }


    public class PPTXTableCell
    {
        public PPTXText Texts { get; set; } = new PPTXText();
    }

    public class PPTXTable
    {
        public PPTXTransform Transform { get; set; } = new PPTXTransform();

        public List<PPTXTableColumn> Columns { get; set; } = new List<PPTXTableColumn>();
        public Dictionary<ValueTuple<int, int>, PPTXTableCell> Cells { get; set; } = new Dictionary<(int, int), PPTXTableCell>();
    }
}
