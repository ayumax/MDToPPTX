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

            Table.Transform = SlideManager.NewTransform;  
        }

        public void AddTableEnd()
        {
            if (CurrentTable == null) return;

            var lastTextAreaSize = 0.0f;

            foreach (var _rowGroupList in CurrentTable.Cells.GroupBy(_cell => _cell.Key.Item1))
            {
                float maxFontHeight = 0;

                foreach (var _rowGroup in _rowGroupList)
                {
                    foreach (var _textRun in _rowGroup.Value.Texts.Texts)
                    {
                        maxFontHeight = Math.Max(maxFontHeight, SlideManager.FontHeght(_textRun.Font) * 1.2f);
                    }
                }

                lastTextAreaSize += maxFontHeight;
            }

            CurrentTable.Transform.SizeY = lastTextAreaSize;
            CurrentTableCell = null;
            CurrentTable = null;
        }

        public void SetTableCell(int RowIndex, int ColIndex)
        {
            if (CurrentTable == null) return;

            CurrentTableCell = new PPTXTableCell();
            CurrentTable.Cells.Add((RowIndex, ColIndex), CurrentTableCell);
        }
    }
}
