using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using P15 = DocumentFormat.OpenXml.Office2013.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;

namespace MDToPPTX.PPTX.OpenXML
{
    class TableSlideHelper : SlideHelperBase
    {
        public void AddContent(ShapeTree shapeTree1, uint ObjectID, PPTXTable Content, Dictionary<string, string> HyperLinkIDMap)
        {
            GraphicFrame graphicFrame1 = new GraphicFrame();

            AddTableCommonProperty(graphicFrame1, ObjectID);

            Transform transform1 = CreateTransform(Content.Transform);

            A.Graphic graphic1 = new A.Graphic();

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" };


            A.Table table1 = new A.Table();

            A.TableProperties tableProperties1 = new A.TableProperties() { FirstRow = true, BandRow = true };
            A.TableStyleId tableStyleId1 = new A.TableStyleId();
            tableStyleId1.Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";

            tableProperties1.Append(tableStyleId1);
            table1.Append(tableProperties1);

            A.TableGrid tableGrid1 = new A.TableGrid();

            foreach(var tableColumn in Content.Columns)
            {
                tableGrid1.Append(CreateColumn(tableColumn.Width));
            }

            table1.Append(tableGrid1);

            foreach (var _tableRow in Content.Rows)
            {
                table1.Append(CreateRow(Content.Columns, _tableRow, HyperLinkIDMap));
            }


            graphicData1.Append(table1);

            graphic1.Append(graphicData1);
       
            graphicFrame1.Append(transform1);
            graphicFrame1.Append(graphic1);

            shapeTree1.Append(graphicFrame1);
        }

        private void AddTableCommonProperty(GraphicFrame graphicFrame1, uint ObjectID)
        {
            NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = new NonVisualGraphicFrameProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties2 = new NonVisualDrawingProperties() { Id = ObjectID, Name = "表 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{7AB8EDC7-F9EF-4752-9A46-413B9437344B}\" />");

            nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

            nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

            nonVisualDrawingProperties2.Append(nonVisualDrawingPropertiesExtensionList1);

            NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoGrouping = true };

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new ApplicationNonVisualDrawingProperties();

            ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList1 = new ApplicationNonVisualDrawingPropertiesExtensionList();

            ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 = new ApplicationNonVisualDrawingPropertiesExtension() { Uri = "{D42A27DB-BD31-4B8C-83A1-F6EECF244321}" };

            P14.ModificationId modificationId1 = new P14.ModificationId() { Val = (UInt32Value)833561296U };
            modificationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            applicationNonVisualDrawingPropertiesExtension1.Append(modificationId1);

            applicationNonVisualDrawingPropertiesExtensionList1.Append(applicationNonVisualDrawingPropertiesExtension1);

            applicationNonVisualDrawingProperties2.Append(applicationNonVisualDrawingPropertiesExtensionList1);

            nonVisualGraphicFrameProperties1.Append(nonVisualDrawingProperties2);
            nonVisualGraphicFrameProperties1.Append(nonVisualGraphicFrameDrawingProperties1);
            nonVisualGraphicFrameProperties1.Append(applicationNonVisualDrawingProperties2);

            graphicFrame1.Append(nonVisualGraphicFrameProperties1);
        }

        private A.GridColumn CreateColumn(float Width)
        {
            A.GridColumn gridColumn1 = new A.GridColumn() { Width = Utils.GetCmToShapeScale(Width) };

            A.ExtensionList extensionList1 = new A.ExtensionList();

            A.Extension extension1 = new A.Extension() { Uri = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}" };

            OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:colId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"3243622648\" />");

            extension1.Append(openXmlUnknownElement2);

            extensionList1.Append(extension1);

            gridColumn1.Append(extensionList1);

            return gridColumn1;
        }

        private A.TableRow CreateRow(List<PPTXTableColumn> Cols, PPTXTableRow Row, Dictionary<string, string> HyperLinkIDMap)
        {
            A.TableRow tableRow1 = new A.TableRow() { Height = (Int64)(Row.Height * 100) };

            foreach(var Cell in Cols.Select((Col, ColIndex)=> new { Col = Col, ColIndex = ColIndex }))
            {
                A.TableCell tableCell1 = new A.TableCell();

                A.TextBody textBody1 = new A.TextBody();
                A.BodyProperties bodyProperties1 = new A.BodyProperties();
                A.ListStyle listStyle1 = new A.ListStyle();

                textBody1.Append(bodyProperties1);
                textBody1.Append(listStyle1);

                foreach (var _textLine in Row.Cells[Cell.ColIndex].Texts.Texts)
                {
                    var paragraph = new A.Paragraph();

                    var cellAlign = A.TextAlignmentTypeValues.Center;
                    switch (Cell.Col.Alignment)
                    {
                        case PPTXTableColumnAlign.Left:
                            cellAlign = A.TextAlignmentTypeValues.Left;
                            break;
                        case PPTXTableColumnAlign.Right:
                            cellAlign = A.TextAlignmentTypeValues.Right;
                            break;
                    }

                    paragraph.Append(new A.ParagraphProperties() { Alignment = cellAlign });

                    paragraph.Append(new A.Run()
                    {
                        RunProperties = CreateRunProperties(_textLine, HyperLinkIDMap),
                        Text = new A.Text(_textLine.Text)
                    });

                    textBody1.Append(paragraph);
                }

                A.TableCellProperties tableCellProperties1 = new A.TableCellProperties();

                tableCell1.Append(textBody1);
                tableCell1.Append(tableCellProperties1);

                tableRow1.Append(tableCell1);
            }
           

            A.ExtensionList extensionList4 = new A.ExtensionList();

            A.Extension extension4 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            OpenXmlUnknownElement openXmlUnknownElement5 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"2927081048\" />");

            extension4.Append(openXmlUnknownElement5);

            extensionList4.Append(extension4);

            tableRow1.Append(extensionList4);

            return tableRow1;
        }
    }
}
