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
    class SlideWriterHelper
    {
        public static void SetSlideID(PresentationPart presentationPart, SlidePart slidePart1)
        {
            // Insert the new slide into the slide list after the previous slide.
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            // Find the highest slide ID in the current list.
            uint maxSlideId = 1;
            SlideId prevSlideId = null;

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                    prevSlideId = slideId;
                }
            }

            maxSlideId++;

            SlideId newSlideId = slideIdList.AppendChild(new SlideId());
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart1);
        }

        public static A.Transform2D CreateTransform2D(PPTXTransform transform)
        {
            A.Transform2D transform2D25 = null;

            if (transform.AutoLayout == false)
            {
                transform2D25 = new A.Transform2D()
                {
                    Offset = new A.Offset()
                    {
                        X = Utils.GetCmToShapeScale(transform.PositionX),
                        Y = Utils.GetCmToShapeScale(transform.PositionY)
                    },
                    Extents = new A.Extents()
                    {
                        Cx = Utils.GetCmToShapeScale(transform.SizeX),
                        Cy = Utils.GetCmToShapeScale(transform.SizeY)
                    }
                };
            }


            return transform2D25;
        }

        public static Transform CreateTransform(PPTXTransform transform)
        {
            Transform retTransform = null;

            if (transform.AutoLayout == false)
            {
                retTransform = new Transform()
                {
                    Offset = new A.Offset()
                    {
                        X = Utils.GetCmToShapeScale(transform.PositionX),
                        Y = Utils.GetCmToShapeScale(transform.PositionY)
                    },
                    Extents = new A.Extents()
                    {
                        Cx = Utils.GetCmToShapeScale(transform.SizeX),
                        Cy = Utils.GetCmToShapeScale(transform.SizeY)
                    }
                };
            }


            return retTransform;
        }

        public static int CreateHyperLinkMap(PPTXSlide SlideContent, SlidePart slidePart1, Dictionary<string, string> HyperLinkIDMap)
        {
            int lastIndex = 2;

            var textRuns = SlideContent.TextAreas
                .SelectMany(_textArea => _textArea.Texts.SelectMany(_text => _text.Texts))
                .Where(_textRun => _textRun.Link.IsEnable)
                .Select((_textRun, _Index) => new { Link = _textRun.Link, Index = _Index });

            foreach (var linkItem in textRuns)
            {
                if (HyperLinkIDMap.ContainsKey(linkItem.Link.LinkKey))
                {
                    continue;
                }

                var linkId = $"rId{linkItem.Index + 2}";
                lastIndex = linkItem.Index + 2;

                slidePart1.AddHyperlinkRelationship(new System.Uri(linkItem.Link.LinkURL, System.UriKind.Absolute), true, linkId);

                HyperLinkIDMap.Add(linkItem.Link.LinkKey, linkId);
            }

            return lastIndex;
        }

        public static A.ParagraphProperties CrateParagraphProperties(PPTXText Content)
        {
            var paragraphPorp = new A.ParagraphProperties();

            var firstTextRun = Content.Texts.FirstOrDefault();
            if (firstTextRun == null) return paragraphPorp;

            switch (firstTextRun.Font.HAlign)
            {
                case EPPTXHAlign.Left:
                    paragraphPorp.Alignment = A.TextAlignmentTypeValues.Left;
                    break;
                case EPPTXHAlign.Center:
                    paragraphPorp.Alignment = A.TextAlignmentTypeValues.Center;
                    break;
                case EPPTXHAlign.Right:
                    paragraphPorp.Alignment = A.TextAlignmentTypeValues.Right;
                    break;
            }

            switch (Content.Bullet)
            {
                case PPTXBullet.None:
                    paragraphPorp.Append(new A.NoBullet());
                    break;

                case PPTXBullet.Circle:
                    paragraphPorp.Append(new A.BulletFont() { Typeface = "Wingdings", Panose = "05000000000000000000", PitchFamily = 2, CharacterSet = 2 });
                    paragraphPorp.Append(new A.CharacterBullet() { Char = "l" });
                    break;

                case PPTXBullet.Rectangle:
                    paragraphPorp.Append(new A.BulletFont() { Typeface = "Wingdings", Panose = "05000000000000000000", PitchFamily = 2, CharacterSet = 2 });
                    paragraphPorp.Append(new A.CharacterBullet() { Char = "n" });
                    break;

                case PPTXBullet.Diamond:
                    paragraphPorp.Append(new A.BulletFont() { Typeface = "Wingdings", Panose = "05000000000000000000", PitchFamily = 2, CharacterSet = 2 });
                    paragraphPorp.Append(new A.CharacterBullet() { Char = "u" });
                    break;

                case PPTXBullet.RectangleBorder:
                    paragraphPorp.Append(new A.BulletFont() { Typeface = "Wingdings", Panose = "05000000000000000000", PitchFamily = 2, CharacterSet = 2 });
                    paragraphPorp.Append(new A.CharacterBullet() { Char = "p" });
                    break;

                case PPTXBullet.Check:
                    paragraphPorp.Append(new A.BulletFont() { Typeface = "Wingdings", Panose = "05000000000000000000", PitchFamily = 2, CharacterSet = 2 });
                    paragraphPorp.Append(new A.CharacterBullet() { Char = "ü" });
                    break;

                case PPTXBullet.Arrow:
                    paragraphPorp.Append(new A.BulletFont() { Typeface = "Wingdings", Panose = "05000000000000000000", PitchFamily = 2, CharacterSet = 2 });
                    paragraphPorp.Append(new A.CharacterBullet() { Char = "Ø" });
                    break;

                case PPTXBullet.MiniCircle:
                    break;

                case PPTXBullet.Number:
                    paragraphPorp.Append(new A.BulletFont() { Typeface = "+mj-lt" });
                    paragraphPorp.Append(new A.AutoNumberedBullet() { Type = A.TextAutoNumberSchemeValues.ArabicPeriod });
                    break;

                case PPTXBullet.CircleNumber:
                    paragraphPorp.Append(new A.BulletFont() { Typeface = "+mj-ea" });
                    paragraphPorp.Append(new A.AutoNumberedBullet() { Type = A.TextAutoNumberSchemeValues.CircleNumberDoubleBytePlain });
                    break;
            }

            return paragraphPorp;
        }

        public static A.RunProperties CreateRunProperties(PPTXTextRun Text, Dictionary<string, string> HyperLinkIDMap)
        {
            A.RunProperties runProperties3 = new A.RunProperties() { Kumimoji = true, Language = "ja-JP", AlternativeLanguage = "en-US", FontSize = (int)(Text.Font.FontSize * 100), Dirty = false };

            runProperties3.Bold = Text.Font.Bold;
            runProperties3.Italic = Text.Font.Italic;

            if (Text.Font.UnderLine)
            {
                runProperties3.Underline = A.TextUnderlineValues.Single;
            }

            if (Text.Font.Strike)
            {
                runProperties3.Strike = A.TextStrikeValues.SingleStrike;
            }

            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = Text.Font.FontFamily, Panose = "020B0604030504040204", PitchFamily = 50, CharacterSet = -128 };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = Text.Font.FontFamily, Panose = "020B0604030504040204", PitchFamily = 50, CharacterSet = -128 };

            if (Text.Font.ForegroundColor.IsTransparent == false)
            {
                A.SolidFill solidFill1 = new A.SolidFill();
                solidFill1.Append(CreateRGBColorModeHex(Text.Font.ForegroundColor));
                runProperties3.Append(solidFill1);
            }

            runProperties3.Append(latinFont1);
            runProperties3.Append(eastAsianFont1);

            if (HyperLinkIDMap.ContainsKey(Text.Link.LinkKey))
            {
                A.HyperlinkOnClick hyperlinkOnClick1 = new A.HyperlinkOnClick() { Id = HyperLinkIDMap[Text.Link.LinkKey] };

                runProperties3.Append(hyperlinkOnClick1);
            }


            return runProperties3;
        }

        public static A.RgbColorModelHex CreateRGBColorModeHex(PPTXColor Color) => new A.RgbColorModelHex() { Val = $"{Color.Color.R:X02}{Color.Color.G:X02}{Color.Color.B:X02}" };
    }
}
