﻿using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Office2010.Drawing;
using NonVisualDrawingProperties = DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties;

namespace OpenXmlPptxCreator
{
    /// <summary>
    /// Create a presentation at a specified file path. The presentation document type is pptx, by default.
    /// </summary>
    public class PpPresentation: IDisposable
    {
        private readonly PresentationDocument presentationDoc = null;
        private readonly PresentationPart presentationPart = null;
        private readonly Slide firstSlide = null;

        private readonly int presentationWidth;
        private readonly int presentationHeight;

        private int countOfSlidesWithContent = 0;

        public PpPresentation(string filepath, int width, int height)
        {
            // Create a presentation at a specified file path. The presentation document type is pptx, by default.
            presentationDoc = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation);
            presentationPart = presentationDoc.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            presentationWidth = width;
            presentationHeight = height;

            firstSlide = CreatePresentationParts(width, height); //9144000, 6858000
        }

        public void AddSlideWithImage(string imagePath)
        {
            if (countOfSlidesWithContent == 0)
            {
                InsertImageInLastSlide(firstSlide, imagePath, "image/png", presentationWidth, presentationHeight);
                firstSlide.Save();
            }
            else
            {
                Slide slide = InsertSlide( string.Empty);
                InsertImageInLastSlide(slide, imagePath, "image/png", presentationWidth, presentationHeight);
                presentationPart.Presentation.Save();
            }
            countOfSlidesWithContent++;
        }

        #region Create Base Presentation Structure

        private Slide CreatePresentationParts(Int32Value width, Int32Value height)
        {
            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
            SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });
            SlideSize slideSize1 = new SlideSize() { Cx = width, Cy = height, Type = SlideSizeValues.Custom };
            NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

            presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1,
                defaultTextStyle1);

            SlidePart slidePart1;
            SlideLayoutPart slideLayoutPart1;
            SlideMasterPart slideMasterPart1;
            ThemePart themePart1;

            slidePart1 = CreateSlidePart(presentationPart, "rId2");
            slideLayoutPart1 = CreateSlideLayoutPart(slidePart1, "rId1");
            slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1, "rId1");
            themePart1 = CreateTheme(slideMasterPart1, "rId5");

            slideMasterPart1.AddPart(slideLayoutPart1, "rId1");
            presentationPart.AddPart(slideMasterPart1, "rId1");
            presentationPart.AddPart(themePart1, "rId5");
            return slidePart1.Slide;
        }

        private SlidePart CreateSlidePart(PresentationPart presentationPart, string slideId)
        {
            SlidePart slidePart1 = presentationPart.AddNewPart<SlidePart>(slideId);
            slidePart1.Slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new TransformGroup()),
                        new P.Shape(
                            new P.NonVisualShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" },
                                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                            new P.ShapeProperties()
                            , new P.TextBody(
                                new BodyProperties(),
                                new ListStyle(),
                                new Paragraph(new EndParagraphRunProperties() { Language = "en-US" }))
                            )
                        )
                    ),
                new ColorMapOverride(new MasterColorMapping()));
            return slidePart1;
        }

        private SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1, string id)
        {
            SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>(id);
            SlideLayout slideLayout = new SlideLayout(
                new CommonSlideData(new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new TransformGroup()),
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
                            new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                        new P.ShapeProperties()
                        //,new P.TextBody(
                        //    new BodyProperties(),
                        //    new ListStyle(),
                        //    new Paragraph(new EndParagraphRunProperties()))
                        )
                    )
                ),
                new ColorMapOverride(new MasterColorMapping()));
            slideLayoutPart1.SlideLayout = slideLayout;
            return slideLayoutPart1;
        }

        private SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1, string id)
        {
            SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>(id);
            SlideMaster slideMaster = new SlideMaster(
                new CommonSlideData(
                    new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new TransformGroup()),
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties()
                            { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
                            new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape()
                            { Type = PlaceholderValues.Title })),
                        new P.ShapeProperties()
                        //,new P.TextBody(
                        //    new BodyProperties(),
                        //    new ListStyle(),
                        //    new Paragraph())
                        )
                    )
                ),
                new P.ColorMap()
                {
                    Background1 = D.ColorSchemeIndexValues.Light1,
                    Text1 = D.ColorSchemeIndexValues.Dark1,
                    Background2 = D.ColorSchemeIndexValues.Light2,
                    Text2 = D.ColorSchemeIndexValues.Dark2,
                    Accent1 = D.ColorSchemeIndexValues.Accent1,
                    Accent2 = D.ColorSchemeIndexValues.Accent2,
                    Accent3 = D.ColorSchemeIndexValues.Accent3,
                    Accent4 = D.ColorSchemeIndexValues.Accent4,
                    Accent5 = D.ColorSchemeIndexValues.Accent5,
                    Accent6 = D.ColorSchemeIndexValues.Accent6,
                    Hyperlink = D.ColorSchemeIndexValues.Hyperlink,
                    FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink
                },
                new SlideLayoutIdList(new SlideLayoutId()
                { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
                new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
            slideMasterPart1.SlideMaster = slideMaster;

            return slideMasterPart1;
        }

        private ThemePart CreateTheme(SlideMasterPart slideMasterPart1, string id)
        {
            ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>(id);
            D.Theme theme1 = new D.Theme() { Name = "Office Theme" };

            D.ThemeElements themeElements1 = new D.ThemeElements(
                new D.ColorScheme(
                        new D.Dark1Color(new D.SystemColor()
                        { Val = D.SystemColorValues.WindowText, LastColor = "000000" }),
                        new D.Light1Color(new D.SystemColor()
                        { Val = D.SystemColorValues.Window, LastColor = "FFFFFF" }),
                        new D.Dark2Color(new D.RgbColorModelHex() { Val = "1F497D" }),
                        new D.Light2Color(new D.RgbColorModelHex() { Val = "EEECE1" }),
                        new D.Accent1Color(new D.RgbColorModelHex() { Val = "4F81BD" }),
                        new D.Accent2Color(new D.RgbColorModelHex() { Val = "C0504D" }),
                        new D.Accent3Color(new D.RgbColorModelHex() { Val = "9BBB59" }),
                        new D.Accent4Color(new D.RgbColorModelHex() { Val = "8064A2" }),
                        new D.Accent5Color(new D.RgbColorModelHex() { Val = "4BACC6" }),
                        new D.Accent6Color(new D.RgbColorModelHex() { Val = "F79646" }),
                        new D.Hyperlink(new D.RgbColorModelHex() { Val = "0000FF" }),
                        new D.FollowedHyperlinkColor(new D.RgbColorModelHex() { Val = "800080" }))
                { Name = "Office" },
                new D.FontScheme(
                        new D.MajorFont(
                            new D.LatinFont() { Typeface = "Calibri" },
                            new D.EastAsianFont() { Typeface = "" },
                            new D.ComplexScriptFont() { Typeface = "" }),
                        new D.MinorFont(
                            new D.LatinFont() { Typeface = "Calibri" },
                            new D.EastAsianFont() { Typeface = "" },
                            new D.ComplexScriptFont() { Typeface = "" }))
                { Name = "Office" },
                new D.FormatScheme(
                        new D.FillStyleList(
                            new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
                            new D.GradientFill(
                                new D.GradientStopList(
                                    new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 50000 },
                                                new D.SaturationModulation() { Val = 300000 })
                                    { Val = D.SchemeColorValues.PhColor })
                                    { Position = 0 },
                                    new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 37000 },
                                                new D.SaturationModulation() { Val = 300000 })
                                    { Val = D.SchemeColorValues.PhColor })
                                    { Position = 35000 },
                                    new D.GradientStop(new D.SchemeColor(new D.Tint() { Val = 15000 },
                                                new D.SaturationModulation() { Val = 350000 })
                                    { Val = D.SchemeColorValues.PhColor })
                                    { Position = 100000 }
                                ),
                                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
                            new D.NoFill(),
                            new D.PatternFill(),
                            new D.GroupFill()),
                        new D.LineStyleList(
                            new D.Outline(
                                new D.SolidFill(
                                    new D.SchemeColor(
                                            new D.Shade() { Val = 95000 },
                                            new D.SaturationModulation() { Val = 105000 })
                                    { Val = D.SchemeColorValues.PhColor }),
                                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
                            {
                                Width = 9525,
                                CapType = D.LineCapValues.Flat,
                                CompoundLineType = D.CompoundLineValues.Single,
                                Alignment = D.PenAlignmentValues.Center
                            },
                            new D.Outline(
                                new D.SolidFill(
                                    new D.SchemeColor(
                                            new D.Shade() { Val = 95000 },
                                            new D.SaturationModulation() { Val = 105000 })
                                    { Val = D.SchemeColorValues.PhColor }),
                                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
                            {
                                Width = 9525,
                                CapType = D.LineCapValues.Flat,
                                CompoundLineType = D.CompoundLineValues.Single,
                                Alignment = D.PenAlignmentValues.Center
                            },
                            new D.Outline(
                                new D.SolidFill(
                                    new D.SchemeColor(
                                            new D.Shade() { Val = 95000 },
                                            new D.SaturationModulation() { Val = 105000 })
                                    { Val = D.SchemeColorValues.PhColor }),
                                new D.PresetDash() { Val = D.PresetLineDashValues.Solid })
                            {
                                Width = 9525,
                                CapType = D.LineCapValues.Flat,
                                CompoundLineType = D.CompoundLineValues.Single,
                                Alignment = D.PenAlignmentValues.Center
                            }),
                        new D.EffectStyleList(
                            new D.EffectStyle(
                                new D.EffectList(
                                    new D.OuterShadow(
                                        new D.RgbColorModelHex(
                                                new D.Alpha() { Val = 38000 })
                                        { Val = "000000" })
                                    {
                                        BlurRadius = 40000L,
                                        Distance = 20000L,
                                        Direction = 5400000,
                                        RotateWithShape = false
                                    })),
                            new D.EffectStyle(
                                new D.EffectList(
                                    new D.OuterShadow(
                                        new D.RgbColorModelHex(
                                                new D.Alpha() { Val = 38000 })
                                        { Val = "000000" })
                                    {
                                        BlurRadius = 40000L,
                                        Distance = 20000L,
                                        Direction = 5400000,
                                        RotateWithShape = false
                                    })),
                            new D.EffectStyle(
                                new D.EffectList(
                                    new D.OuterShadow(
                                        new D.RgbColorModelHex(
                                                new D.Alpha() { Val = 38000 })
                                        { Val = "000000" })
                                    {
                                        BlurRadius = 40000L,
                                        Distance = 20000L,
                                        Direction = 5400000,
                                        RotateWithShape = false
                                    }))),
                        new D.BackgroundFillStyleList(
                            new D.SolidFill(new D.SchemeColor() { Val = D.SchemeColorValues.PhColor }),
                            new D.GradientFill(
                                new D.GradientStopList(
                                    new D.GradientStop(
                                            new D.SchemeColor(new D.Tint() { Val = 50000 },
                                                    new D.SaturationModulation() { Val = 300000 })
                                            { Val = D.SchemeColorValues.PhColor })
                                    { Position = 0 },
                                    new D.GradientStop(
                                            new D.SchemeColor(new D.Tint() { Val = 50000 },
                                                    new D.SaturationModulation() { Val = 300000 })
                                            { Val = D.SchemeColorValues.PhColor })
                                    { Position = 0 },
                                    new D.GradientStop(
                                            new D.SchemeColor(new D.Tint() { Val = 50000 },
                                                    new D.SaturationModulation() { Val = 300000 })
                                            { Val = D.SchemeColorValues.PhColor })
                                    { Position = 0 }),
                                new D.LinearGradientFill() { Angle = 16200000, Scaled = true }),
                            new D.GradientFill(
                                new D.GradientStopList(
                                    new D.GradientStop(
                                            new D.SchemeColor(new D.Tint() { Val = 50000 },
                                                    new D.SaturationModulation() { Val = 300000 })
                                            { Val = D.SchemeColorValues.PhColor })
                                    { Position = 0 },
                                    new D.GradientStop(
                                            new D.SchemeColor(new D.Tint() { Val = 50000 },
                                                    new D.SaturationModulation() { Val = 300000 })
                                            { Val = D.SchemeColorValues.PhColor })
                                    { Position = 0 }),
                                new D.LinearGradientFill() { Angle = 16200000, Scaled = true })))
                { Name = "Office" });

            theme1.Append(themeElements1);
            theme1.Append(new D.ObjectDefaults());
            theme1.Append(new D.ExtraColorSchemeList());

            themePart1.Theme = theme1;
            return themePart1;

        }

        #endregion

        #region Add Slide with Image

        /// <summary>
        /// Insert a new Slide into PowerPoint
        /// </summary>
        /// <param name="presentationPart">Presentation Part</param>
        /// <param name="layoutName">Layout of the new Slide</param>
        /// <returns>Slide Instance</returns>
        private Slide InsertSlide(string layoutName)
        {
            UInt32 slideId = 256U;
            // Get the Slide Id collection of the presentation document
            var slideIdList = presentationPart.Presentation.SlideIdList;
            if (slideIdList == null)
            {
                throw new NullReferenceException("The number of slide is empty, please select a ppt with a slide at least again");
            }
            slideId += Convert.ToUInt32(slideIdList.Count());

            // Creates an Slide instance and adds its children.
            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
            slide.Save(slidePart);

            // Get SlideMasterPart and SlideLayoutPart from the existing Presentation Part
            SlideMasterPart slideMasterPart = presentationPart.SlideMasterParts.First();
            SlideLayoutPart slideLayoutPart = slideMasterPart.SlideLayoutParts.FirstOrDefault();
            if (slideLayoutPart == null)
            {
                throw new Exception("The slide layout " + layoutName + " is not found");
            }

            slidePart.AddPart<SlideLayoutPart>(slideLayoutPart);
            slidePart.Slide.CommonSlideData = (CommonSlideData)slideMasterPart.SlideLayoutParts.FirstOrDefault()?.SlideLayout.CommonSlideData.Clone();

            // Create SlideId instance and Set property
            SlideId newSlideId = presentationPart.Presentation.SlideIdList.AppendChild<SlideId>(new SlideId());
            newSlideId.Id = slideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);
            return GetSlideByRelationShipId(presentationPart, newSlideId.RelationshipId);
        }

        /// <summary>
        /// Get Slide By RelationShip ID
        /// </summary>
        /// <param name="presentationPart">Presentation Part</param>
        /// <param name="relationshipId">Relationship ID</param>
        /// <returns>Slide Object</returns>
        private Slide GetSlideByRelationShipId(PresentationPart presentationPart, StringValue relationshipId)
        {
            // Get Slide object by Relationship ID
            SlidePart slidePart = presentationPart.GetPartById(relationshipId) as SlidePart;
            if (slidePart != null)
            {
                return slidePart.Slide;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Insert Image into Slide
        /// </summary>
        /// <param name="filePath">PowerPoint Path</param>
        /// <param name="imagePath">Image Path</param>
        /// <param name="imageExt">Image Extension</param>
        private void InsertImageInLastSlide(Slide slide, string imagePath, string imageExt, Int64Value width, Int64Value height)
        {
            // Creates an Picture instance and adds its children.
            P.Picture picture = new P.Picture();
            string embedId = string.Empty;
            embedId = "rId" + (slide.Elements().Count() + 915).ToString();
            P.NonVisualPictureProperties nonVisualPictureProperties = new P.NonVisualPictureProperties(new P.NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Picture 5" }, new P.NonVisualPictureDrawingProperties(new D.PictureLocks() { NoChangeAspect = true }), new ApplicationNonVisualDrawingProperties());
            P.BlipFill blipFill = new P.BlipFill();
            Blip blip = new Blip() { Embed = embedId };

            // Creates an BlipExtensionList instance and adds its children
            BlipExtensionList blipExtensionList = new BlipExtensionList();
            BlipExtension blipExtension = new BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };
            UseLocalDpi useLocalDpi = new UseLocalDpi() { Val = false };
            useLocalDpi.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
            blipExtension.Append(useLocalDpi);
            blipExtensionList.Append(blipExtension);
            blip.Append(blipExtensionList);
            Stretch stretch = new Stretch();
            FillRectangle fillRectangle = new FillRectangle();
            stretch.Append(fillRectangle);
            blipFill.Append(blip);
            blipFill.Append(stretch);

            // Creates an ShapeProperties instance and adds its children.
            P.ShapeProperties shapeProperties = new P.ShapeProperties();
            D.Transform2D transform2D = new D.Transform2D();
            //D.Offset offset = new D.Offset() { X = 457200L, Y = 1524000L };
            D.Offset offset = new D.Offset() { X = 0L, Y = 0L };
            //D.Extents extents = new D.Extents() { Cx = 8229600L, Cy = 5029200L };
            D.Extents extents = new D.Extents() { Cx = width, Cy = height };
            transform2D.Append(offset);
            transform2D.Append(extents);
            D.PresetGeometry presetGeometry = new D.PresetGeometry() { Preset = D.ShapeTypeValues.Rectangle };
            D.AdjustValueList adjustValueList = new D.AdjustValueList();
            presetGeometry.Append(adjustValueList);
            shapeProperties.Append(transform2D);
            shapeProperties.Append(presetGeometry);
            picture.Append(nonVisualPictureProperties);
            picture.Append(blipFill);
            picture.Append(shapeProperties);
            slide.CommonSlideData.ShapeTree.AppendChild(picture);

            // Generates content of imagePart.
            ImagePart imagePart = slide.SlidePart.AddNewPart<ImagePart>(imageExt, embedId);
            FileStream fileStream = new FileStream(imagePath, FileMode.Open);
            imagePart.FeedData(fileStream);
            fileStream.Close();
        }

        #endregion

        public void Dispose()
        {
            presentationDoc.Close();
            presentationDoc?.Dispose();
        }
    }
}
