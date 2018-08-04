using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using MDToPPTX.PPTX;

namespace MDToPPTX.Markdown
{
    class SlideTableManager
    {
        private PPTXTable CurrentTable;
        private PPTXTableCell CurrentTableCell;

        private SlideManager SlideManager;

        public void Init(SlideManager SlideManager)
        {
            this.SlideManager = SlideManager;

            CurrentTable = null;
            CurrentTableCell = null;
        }

        public bool IsReadyCell => CurrentTableCell != null;

        public void Write(PPTXTextRun Text)
        {
            if (SlideManager.FontStack.Count > 0)
            {
                Text.Font = SlideManager.FontStack.Peek();
            }

            if (SlideManager.LinkStack.Count > 0)
            {
                Text.Link = SlideManager.LinkStack.Peek();
            }

            CurrentTableCell.Texts.Texts.Add(Text);
        }

        public void AddTable(PPTXTable Table)
        {
            CurrentTable = Table;

            Table.Transform = SlideManager.NewTransform();  

            foreach(var col in Table.Columns)
            {
                col.Width = Table.Transform.SizeX / (float)Table.Columns.Count;
            }
        }

        public void AddTableEnd()
        {
            if (CurrentTable == null) return;

            var lastTextAreaSize = 0.0f;

            lastTextAreaSize = CurrentTable.Rows.Sum(_row => _row.Height);

            CurrentTable.Transform.SizeY = lastTextAreaSize;
            CurrentTableCell = null;
            CurrentTable = null;

            SlideManager.SetContentTransform(CurrentTable.Transform);
        }

        public void AddTableRow()
        {
            if (CurrentTable == null) return;

            CurrentTable.Rows.Add(new PPTXTableRow());
        }

        public void NextTableCell()
        {
            if (CurrentTable == null) return;

            CurrentTableCell = new PPTXTableCell();
            CurrentTable.Rows.Last().Cells.Add(CurrentTableCell);
        }

        public void EndTableRow()
        {
            if (CurrentTable == null) return;

            float maxFontHeight = 0;

            var lastRow = CurrentTable.Rows.Last();

            foreach (var _cell in lastRow.Cells)
            {
                foreach (var _textRun in _cell.Texts.Texts)
                {
                    maxFontHeight = Math.Max(maxFontHeight, SlideManager.FontHeght(_textRun.Font) * 1.2f);
                }
            }

            lastRow.Height = maxFontHeight;
        }
    }
}
