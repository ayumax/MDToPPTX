using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using D = DocumentFormat.OpenXml.Drawing;

namespace MDToPPTX
{
    static class AddShape
    {
        private const int cm2shapescale = 360000;
        private const int degree2shapescale = 60000;

        static private ShapeStyle makeShapeStyle()
        {
            ShapeStyle shapeStyle1 = new ShapeStyle();

            D.LineReference lineReference1 = new D.LineReference() { Index = (UInt32Value)2U };

            D.SchemeColor schemeColor2 = new D.SchemeColor() { Val = D.SchemeColorValues.Accent1 };
            D.Shade shade1 = new D.Shade() { Val = 50000 };

            schemeColor2.Append(shade1);

            lineReference1.Append(schemeColor2);

            D.FillReference fillReference1 = new D.FillReference() { Index = (UInt32Value)1U };
            D.SchemeColor schemeColor3 = new D.SchemeColor() { Val = D.SchemeColorValues.Accent1 };

            fillReference1.Append(schemeColor3);

            D.EffectReference effectReference1 = new D.EffectReference() { Index = (UInt32Value)0U };
            D.SchemeColor schemeColor4 = new D.SchemeColor() { Val = D.SchemeColorValues.Accent1 };

            effectReference1.Append(schemeColor4);

            D.FontReference fontReference1 = new D.FontReference() { Index = D.FontCollectionIndexValues.Minor };
            D.SchemeColor schemeColor5 = new D.SchemeColor() { Val = D.SchemeColorValues.Light1 };

            fontReference1.Append(schemeColor5);

            shapeStyle1.Append(lineReference1);
            shapeStyle1.Append(fillReference1);
            shapeStyle1.Append(effectReference1);
            shapeStyle1.Append(fontReference1);
            return shapeStyle1;

        }



        static private TextBody makeTextBody()
        {
            TextBody textBody1 = new TextBody();
            D.BodyProperties bodyProperties1 = new D.BodyProperties() { RightToLeftColumns = false, Anchor = D.TextAnchoringTypeValues.Center };
            D.ListStyle listStyle1 = new D.ListStyle();

            D.Paragraph paragraph1 = new D.Paragraph();
            D.ParagraphProperties paragraphProperties1 = new D.ParagraphProperties() { Alignment = D.TextAlignmentTypeValues.Center };
            D.EndParagraphRunProperties endParagraphRunProperties1 = new D.EndParagraphRunProperties() { Language = "es-ES" };

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(endParagraphRunProperties1);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph1);
            return textBody1;
        }


        static private NonVisualShapeProperties makeNonVisualShapeProperties()
        {
            NonVisualShapeProperties nonVisualShapeProperties1 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties1 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "1 Shape Name" };
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties1 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties1.Append(nonVisualDrawingProperties1);
            nonVisualShapeProperties1.Append(nonVisualShapeDrawingProperties1);
            nonVisualShapeProperties1.Append(applicationNonVisualDrawingProperties1);
            return nonVisualShapeProperties1;
        }


        static private ShapeProperties makeShapeProperties(
            D.ShapeTypeValues shapeType = D.ShapeTypeValues.Rectangle,// Any of the built-in shapes (ellipse, rectangle, etc)
            string fill_rgbColorHex = "EEECE1", // Hexadecimal RGB color code to fill the shape.
            bool isnooutline = false,  //no outline
            D.SchemeColorValues outlineSchemeColor = D.SchemeColorValues.Text1,
            long x = 360000, // Represents the shape x position in 1/36000 cm.
            long y = 720000, // Represents the shape y position in 1/36000 cm.
            long width = 720000, // Shapw width in in 1/36000 cm.
            long height = 720000, // Shapw height in in 1/36000 cm.
            bool horizontalFlip = false,
            bool verticalFlip = false,
            int angle = 0,  //2700000 * 4 = 180 degree
            bool isdashline = false,
            bool isTailEndArrow = false,
            bool isHeadEndArrow = false
            )
        {
            ShapeProperties shapeProperties1 = new ShapeProperties();

            D.Transform2D transform2D1 = new D.Transform2D();
            D.Offset offset1 = new D.Offset() { X = x, Y = y };
            D.Extents extents1 = new D.Extents() { Cx = width, Cy = height };

            transform2D1.HorizontalFlip = horizontalFlip;
            transform2D1.VerticalFlip = verticalFlip;
            transform2D1.Rotation = angle;

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            D.PresetGeometry presetGeometry1 = new D.PresetGeometry() { Preset = shapeType };
            D.AdjustValueList adjustValueList1 = new D.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            D.SolidFill solidFill1 = new D.SolidFill();
            D.RgbColorModelHex rgbColorModelHex1 = new D.RgbColorModelHex() { Val = fill_rgbColorHex };

            solidFill1.Append(rgbColorModelHex1);

            D.Outline outline1 = new D.Outline() { Width = 12700 };
            if (isnooutline)
            {
                D.NoFill nofill = new D.NoFill();
                outline1.Append(nofill);
            }
            else
            {
                D.SolidFill solidFill2 = new D.SolidFill();
                D.SchemeColor schemeColor1 = new D.SchemeColor() { Val = outlineSchemeColor };
                solidFill2.Append(schemeColor1);
                outline1.Append(solidFill2);

                //dash
                if (isdashline)
                {
                    D.PresetDash presetDash = new D.PresetDash() { Val = D.PresetLineDashValues.Dash };
                    outline1.Append(presetDash);
                }

                //arrow
                if (isTailEndArrow)
                {
                    D.TailEnd tailend = new D.TailEnd() { Type = D.LineEndValues.Arrow };
                    outline1.Append(tailend);
                }
                if (isHeadEndArrow)
                {
                    D.HeadEnd headend = new D.HeadEnd() { Type = D.LineEndValues.Arrow };
                    outline1.Append(headend);
                }
            }

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(solidFill1);
            shapeProperties1.Append(outline1);
            return shapeProperties1;
        }
        public static void AddShape_(
                ShapeTree ppshapeTree,
                D.ShapeTypeValues shapeType = D.ShapeTypeValues.Rectangle,// Any of the built-in shapes (ellipse, rectangle, etc)
                string fill_rgbColorHex = "EEECE1", // Hexadecimal RGB color code to fill the shape.
                bool isnooutline = false,
                D.SchemeColorValues outlineSchemeColor = D.SchemeColorValues.Text1,
                float x = 1, // Represents the shape x position in 1 cm.
                float y = 2, // Represents the shape y position in 2 cm.
                float width = 2, // Shapw width in in 2 cm.
                float height = 2, // Shapw height in in 2 cm.
                bool horizontalFlip = false,
                bool verticalFlip = false,
                int angle = 0,   //degree
                bool isdashline = false,
                bool isTailEndArrow = false,
                bool isHeadEndArrow = false
            )
        {
            Shape shape1 = new Shape();
            shape1.Append(makeNonVisualShapeProperties());
            shape1.Append(makeShapeProperties(
                shapeType: shapeType,
                fill_rgbColorHex: fill_rgbColorHex,
                isnooutline: isnooutline,
                outlineSchemeColor: outlineSchemeColor,
                x: getcm2shapescale(x),
                y: getcm2shapescale(y),
                width: getcm2shapescale(width),
                height: getcm2shapescale(height),
                horizontalFlip: horizontalFlip, verticalFlip: verticalFlip,
                angle: getDegree2shapescale(angle),
                isdashline: isdashline,
                isTailEndArrow: isTailEndArrow,
                isHeadEndArrow: isHeadEndArrow)
                );
            shape1.Append(makeShapeStyle());
            shape1.Append(makeTextBody());
            ppshapeTree.Append(shape1);

        }

        static long getcm2shapescale(float cm_val)
        {
            return (long)(cm_val * cm2shapescale);
        }

        static int getDegree2shapescale(int angle)
        {
            return angle * degree2shapescale;
        }

  
    }
    class Program
    {
        static void generateShapes(ShapeTree ppshapeTree)
        {
            AddShape.AddShape_(ppshapeTree);
    //ppshapeTree,
    //D.ShapeTypeValues.Rectangle,
    //x: (0),
    //y: (100),
    //width: (360000),
    //height: (720000)
    //);
        }

        static void Main(string[] args)
        {
            ShapeTree ppshapeTree;
            string fileName = @"C:\Users\ayuma\Desktop\sample.pptx";
            PresentationDocument ppt = null;
            using (ppt = PresentationDocument.Open(fileName, true))
            {
                Console.WriteLine("\"" + fileName + "\" has opened.");

                // Get the relationship ID of the first slide.
                PresentationPart part = ppt.PresentationPart;
                OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
                string relId = (slideIds[0] as SlideId).RelationshipId;

                // Get the slide part from the relationship ID.
                SlidePart ppslide = (SlidePart)part.GetPartById(relId);
                if (ppslide != null)
                {
                    ppshapeTree = ppslide.Slide.CommonSlideData.ShapeTree;
                    generateShapes(ppshapeTree);

                    ppslide.Slide.Save();
                    Console.WriteLine("\"" + fileName + "\" has been changed.");
                }
            }
        }
    }
}
