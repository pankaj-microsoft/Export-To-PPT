using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P15 = DocumentFormat.OpenXml.Office2013.PowerPoint;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;

namespace CoreGoals
{
    public class CreatePresentationServiceV2
    {

        // This is temporary data for the presentation. 
        public List<Goal.GoalDetail> GetGoalsArray()
        {
            List<Goal.GoalDetail> goalsArray = new List<Goal.GoalDetail>();

            for (int i = 1; i <= 3601; i++)
            {
                Random rnd = new Random();

                goalsArray.Add(
                    new Goal.GoalDetail(i % 2 != 0, "Goal Name - " + i.ToString(), "Goal Owner - " + i.ToString(), (i + 1 * 10) + "K")
                );
            }

            return goalsArray;
        }

        // Creates a PresentationDocument.
        public void CreatePackage(string filePath, List<Goal.GoalDetail> goalsArray)
        {
            using (PresentationDocument package = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation))
            {
                CreateParts(package, goalsArray);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(PresentationDocument document, List<Goal.GoalDetail> goalsArray)
        {

            // ****************************** Custom Code Changes Start *************************************
            int totalSlides = (int)Math.Ceiling(goalsArray.Count / 6.0);
            Console.WriteLine("Total Slides: " + totalSlides);
            // ****************************** Custom Code Changes End *************************************

            // This is a general presentation level setting and should be set once per presentation.
            // ****************************** Presentation level settings Start *************************************
            ThumbnailPart thumbnailPart1 = document.AddNewPart<ThumbnailPart>("image/jpeg", "rId2");
            GenerateThumbnailPart1Content(thumbnailPart1);

            PresentationPart presentationPart1 = document.AddPresentationPart();
            GeneratePresentationPart1Content(presentationPart1, totalSlides);

            SlideMasterPart slideMasterPart1 = presentationPart1.AddNewPart<SlideMasterPart>("rId1");
            GenerateSlideMasterPart1Content(slideMasterPart1);

            SlideLayoutPart slideLayoutPart1 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId1");
            GenerateSlideLayoutPart1Content(slideLayoutPart1);

            slideLayoutPart1.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart2 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId8");
            GenerateSlideLayoutPart2Content(slideLayoutPart2);

            slideLayoutPart2.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart3 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId3");
            GenerateSlideLayoutPart3Content(slideLayoutPart3);

            slideLayoutPart3.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart4 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId7");
            GenerateSlideLayoutPart4Content(slideLayoutPart4);

            slideLayoutPart4.AddPart(slideMasterPart1, "rId1");

            ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId12");
            GenerateThemePart1Content(themePart1);

            presentationPart1.AddPart(themePart1, "rId6");

            SlideLayoutPart slideLayoutPart5 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId2");
            GenerateSlideLayoutPart5Content(slideLayoutPart5);

            slideLayoutPart5.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart6 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId6");
            GenerateSlideLayoutPart6Content(slideLayoutPart6);

            slideLayoutPart6.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart7 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId11");
            GenerateSlideLayoutPart7Content(slideLayoutPart7);

            slideLayoutPart7.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart8 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId5");
            GenerateSlideLayoutPart8Content(slideLayoutPart8);

            slideLayoutPart8.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart9 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId10");
            GenerateSlideLayoutPart9Content(slideLayoutPart9);

            slideLayoutPart9.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart10 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId4");
            GenerateSlideLayoutPart10Content(slideLayoutPart10);

            slideLayoutPart10.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart11 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId9");
            GenerateSlideLayoutPart11Content(slideLayoutPart11);

            slideLayoutPart11.AddPart(slideMasterPart1, "rId1");

            TableStylesPart tableStylesPart1 = presentationPart1.AddNewPart<TableStylesPart>("rId7");
            GenerateTableStylesPart1Content(tableStylesPart1);

            // ****************************** Presentation level settings End *************************************

            // ****************************** Slide Creation Start *************************************


            for (int i = 0; i < totalSlides; i++)
            {
                List<Goal.GoalDetail> goalsArrayForSlide = goalsArray.Skip(i * 6).Take(6).ToList();

                SlidePart slidePart1 = presentationPart1.AddNewPart<SlidePart>("rId" + (i + 20).ToString());
                GenerateSlidePartContent(slidePart1, goalsArrayForSlide);

                slidePart1.AddPart(slideLayoutPart1, "rId1");

                ImagePart imagePart1 = slidePart1.AddNewPart<ImagePart>("image/png", "rId3");
                GenerateImagePart1Content(imagePart1);

                ImagePart imagePart2 = slidePart1.AddNewPart<ImagePart>("image/svg+xml", "rId7");
                GenerateImagePart2Content(imagePart2);

                // ImagePart imagePart3 = slidePart1.AddNewPart<ImagePart>("image/png", "rId2");
                // GenerateImagePart3Content(imagePart3);

                ImagePart imagePart4 = slidePart1.AddNewPart<ImagePart>("image/png", "rId6");
                GenerateImagePart4Content(imagePart4);

                ImagePart imagePart5 = slidePart1.AddNewPart<ImagePart>("image/svg+xml", "rId5");
                GenerateImagePart5Content(imagePart5);

                ImagePart imagePart6 = slidePart1.AddNewPart<ImagePart>("image/png", "rId4");
                GenerateImagePart6Content(imagePart6);
            }

            // ****************************** Slide Creation End *************************************

            ViewPropertiesPart viewPropertiesPart1 = presentationPart1.AddNewPart<ViewPropertiesPart>("rId5");
            GenerateViewPropertiesPart1Content(viewPropertiesPart1);

            PresentationPropertiesPart presentationPropertiesPart1 = presentationPart1.AddNewPart<PresentationPropertiesPart>("rId4");
            GeneratePresentationPropertiesPart1Content(presentationPropertiesPart1);

            ExtendedPart extendedPart1 = document.AddExtendedPart("http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels", "application/vnd.ms-office.classificationlabels+xml", "xml", "rId5");
            GenerateExtendedPart1Content(extendedPart1);

            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId4");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            SetPackageProperties(document);
        }

        // Generates content of thumbnailPart1.
        private void GenerateThumbnailPart1Content(ThumbnailPart thumbnailPart1)
        {
            System.IO.Stream data = GetBinaryDataStream(thumbnailPart1Data);
            thumbnailPart1.FeedData(data);
            data.Close();
        }

        // Generates content of presentationPart1.
        private void GeneratePresentationPart1Content(PresentationPart presentationPart1, int totalSlides)
        {
            Presentation presentation1 = new Presentation() { SaveSubsetFonts = true, AutoCompressPictures = false };
            presentation1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            presentation1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            presentation1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList();
            SlideMasterId slideMasterId1 = new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" };

            slideMasterIdList1.Append(slideMasterId1);

            SlideIdList slideIdList1 = new SlideIdList();
            // ****************************** Custom Code Start *************************************
            for (int i = 0; i < totalSlides; i++)
            {
                SlideId slideId = new SlideId() { Id = (UInt32Value)(i + 256U), RelationshipId = "rId" + (i + 20).ToString() };
                slideIdList1.Append(slideId);
            }
            // ****************************** Custom Code End *************************************
            SlideSize slideSize1 = new SlideSize() { Cx = 12192000, Cy = 6858000 };
            NotesSize notesSize1 = new NotesSize() { Cx = 6858000L, Cy = 9144000L };

            DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

            A.DefaultParagraphProperties defaultParagraphProperties1 = new A.DefaultParagraphProperties();
            A.DefaultRunProperties defaultRunProperties1 = new A.DefaultRunProperties() { Language = "en-US" };

            defaultParagraphProperties1.Append(defaultRunProperties1);

            A.Level1ParagraphProperties level1ParagraphProperties1 = new A.Level1ParagraphProperties() { LeftMargin = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties2 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill1.Append(schemeColor1);
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties2.Append(solidFill1);
            defaultRunProperties2.Append(latinFont1);
            defaultRunProperties2.Append(eastAsianFont1);
            defaultRunProperties2.Append(complexScriptFont1);

            level1ParagraphProperties1.Append(defaultRunProperties2);

            A.Level2ParagraphProperties level2ParagraphProperties1 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties3 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill2.Append(schemeColor2);
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties3.Append(solidFill2);
            defaultRunProperties3.Append(latinFont2);
            defaultRunProperties3.Append(eastAsianFont2);
            defaultRunProperties3.Append(complexScriptFont2);

            level2ParagraphProperties1.Append(defaultRunProperties3);

            A.Level3ParagraphProperties level3ParagraphProperties1 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties4 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill3.Append(schemeColor3);
            A.LatinFont latinFont3 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont3 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont3 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties4.Append(solidFill3);
            defaultRunProperties4.Append(latinFont3);
            defaultRunProperties4.Append(eastAsianFont3);
            defaultRunProperties4.Append(complexScriptFont3);

            level3ParagraphProperties1.Append(defaultRunProperties4);

            A.Level4ParagraphProperties level4ParagraphProperties1 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties5 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill4.Append(schemeColor4);
            A.LatinFont latinFont4 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont4 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont4 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties5.Append(solidFill4);
            defaultRunProperties5.Append(latinFont4);
            defaultRunProperties5.Append(eastAsianFont4);
            defaultRunProperties5.Append(complexScriptFont4);

            level4ParagraphProperties1.Append(defaultRunProperties5);

            A.Level5ParagraphProperties level5ParagraphProperties1 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties6 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill5.Append(schemeColor5);
            A.LatinFont latinFont5 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont5 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont5 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties6.Append(solidFill5);
            defaultRunProperties6.Append(latinFont5);
            defaultRunProperties6.Append(eastAsianFont5);
            defaultRunProperties6.Append(complexScriptFont5);

            level5ParagraphProperties1.Append(defaultRunProperties6);

            A.Level6ParagraphProperties level6ParagraphProperties1 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties7 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill6 = new A.SolidFill();
            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill6.Append(schemeColor6);
            A.LatinFont latinFont6 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont6 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont6 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties7.Append(solidFill6);
            defaultRunProperties7.Append(latinFont6);
            defaultRunProperties7.Append(eastAsianFont6);
            defaultRunProperties7.Append(complexScriptFont6);

            level6ParagraphProperties1.Append(defaultRunProperties7);

            A.Level7ParagraphProperties level7ParagraphProperties1 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties8 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill7 = new A.SolidFill();
            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill7.Append(schemeColor7);
            A.LatinFont latinFont7 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont7 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont7 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties8.Append(solidFill7);
            defaultRunProperties8.Append(latinFont7);
            defaultRunProperties8.Append(eastAsianFont7);
            defaultRunProperties8.Append(complexScriptFont7);

            level7ParagraphProperties1.Append(defaultRunProperties8);

            A.Level8ParagraphProperties level8ParagraphProperties1 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties9 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill8 = new A.SolidFill();
            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill8.Append(schemeColor8);
            A.LatinFont latinFont8 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont8 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont8 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties9.Append(solidFill8);
            defaultRunProperties9.Append(latinFont8);
            defaultRunProperties9.Append(eastAsianFont8);
            defaultRunProperties9.Append(complexScriptFont8);

            level8ParagraphProperties1.Append(defaultRunProperties9);

            A.Level9ParagraphProperties level9ParagraphProperties1 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties10 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill9 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill9.Append(schemeColor9);
            A.LatinFont latinFont9 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont9 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont9 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties10.Append(solidFill9);
            defaultRunProperties10.Append(latinFont9);
            defaultRunProperties10.Append(eastAsianFont9);
            defaultRunProperties10.Append(complexScriptFont9);

            level9ParagraphProperties1.Append(defaultRunProperties10);

            defaultTextStyle1.Append(defaultParagraphProperties1);
            defaultTextStyle1.Append(level1ParagraphProperties1);
            defaultTextStyle1.Append(level2ParagraphProperties1);
            defaultTextStyle1.Append(level3ParagraphProperties1);
            defaultTextStyle1.Append(level4ParagraphProperties1);
            defaultTextStyle1.Append(level5ParagraphProperties1);
            defaultTextStyle1.Append(level6ParagraphProperties1);
            defaultTextStyle1.Append(level7ParagraphProperties1);
            defaultTextStyle1.Append(level8ParagraphProperties1);
            defaultTextStyle1.Append(level9ParagraphProperties1);

            PresentationExtensionList presentationExtensionList1 = new PresentationExtensionList();

            PresentationExtension presentationExtension1 = new PresentationExtension() { Uri = "{EFAFB233-063F-42B5-8137-9DF3F51BA10A}" };

            P15.SlideGuideList slideGuideList1 = new P15.SlideGuideList();
            slideGuideList1.AddNamespaceDeclaration("p15", "http://schemas.microsoft.com/office/powerpoint/2012/main");

            presentationExtension1.Append(slideGuideList1);

            presentationExtensionList1.Append(presentationExtension1);

            presentation1.Append(slideMasterIdList1);
            presentation1.Append(slideIdList1);
            presentation1.Append(slideSize1);
            presentation1.Append(notesSize1);
            presentation1.Append(defaultTextStyle1);
            presentation1.Append(presentationExtensionList1);

            presentationPart1.Presentation = presentation1;
        }

        // Generates content of slidePart1.
        private void GenerateSlidePartContent(SlidePart slidePart1, List<Goal.GoalDetail> goalsArrayForSlide)
        {
            Slide slide1 = new Slide();
            slide1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slide1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slide1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData1 = new CommonSlideData();

            ShapeTree shapeTree1 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties1 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties1 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties1 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties1 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties1.Append(nonVisualDrawingProperties1);
            nonVisualGroupShapeProperties1.Append(nonVisualGroupShapeDrawingProperties1);
            nonVisualGroupShapeProperties1.Append(applicationNonVisualDrawingProperties1);

            GroupShapeProperties groupShapeProperties1 = new GroupShapeProperties();

            A.TransformGroup transformGroup1 = new A.TransformGroup();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset1 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents1 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup1.Append(offset1);
            transformGroup1.Append(extents1);
            transformGroup1.Append(childOffset1);
            transformGroup1.Append(childExtents1);

            groupShapeProperties1.Append(transformGroup1);

            Shape shape1 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties1 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties2 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{45F6242F-AC5D-C0BD-BDE9-01DEB85343A0}\" />");

            // nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

            nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

            nonVisualDrawingProperties2.Append(nonVisualDrawingPropertiesExtensionList1);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties1 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties1.Append(shapeLocks1);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape1 = new PlaceholderShape() { Type = PlaceholderValues.CenteredTitle };

            applicationNonVisualDrawingProperties2.Append(placeholderShape1);

            nonVisualShapeProperties1.Append(nonVisualDrawingProperties2);
            nonVisualShapeProperties1.Append(nonVisualShapeDrawingProperties1);
            nonVisualShapeProperties1.Append(applicationNonVisualDrawingProperties2);
            ShapeProperties shapeProperties1 = new ShapeProperties();

            TextBody textBody1 = new TextBody();
            A.BodyProperties bodyProperties1 = new A.BodyProperties();
            A.ListStyle listStyle1 = new A.ListStyle();

            A.Paragraph paragraph1 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties1 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph1.Append(endParagraphRunProperties1);

            textBody1.Append(bodyProperties1);
            textBody1.Append(listStyle1);
            textBody1.Append(paragraph1);

            shape1.Append(nonVisualShapeProperties1);
            shape1.Append(shapeProperties1);
            shape1.Append(textBody1);

            Shape shape2 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties2 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties3 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Subtitle 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList2 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension2 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{7673E653-A5EE-3E20-32AE-02B33E22DEA9}\" />");

            // nonVisualDrawingPropertiesExtension2.Append(openXmlUnknownElement2);

            nonVisualDrawingPropertiesExtensionList2.Append(nonVisualDrawingPropertiesExtension2);

            nonVisualDrawingProperties3.Append(nonVisualDrawingPropertiesExtensionList2);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties2 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks2 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties2.Append(shapeLocks2);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties3 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape2 = new PlaceholderShape() { Type = PlaceholderValues.SubTitle, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties3.Append(placeholderShape2);

            nonVisualShapeProperties2.Append(nonVisualDrawingProperties3);
            nonVisualShapeProperties2.Append(nonVisualShapeDrawingProperties2);
            nonVisualShapeProperties2.Append(applicationNonVisualDrawingProperties3);
            ShapeProperties shapeProperties2 = new ShapeProperties();

            TextBody textBody2 = new TextBody();
            A.BodyProperties bodyProperties2 = new A.BodyProperties();
            A.ListStyle listStyle2 = new A.ListStyle();

            A.Paragraph paragraph2 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph2.Append(endParagraphRunProperties2);

            textBody2.Append(bodyProperties2);
            textBody2.Append(listStyle2);
            textBody2.Append(paragraph2);

            shape2.Append(nonVisualShapeProperties2);
            shape2.Append(shapeProperties2);
            shape2.Append(textBody2);

            Picture picture1 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties1 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties4 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Picture 3", Description = "A blurry image of a dark green and white background\n\nDescription automatically generated" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList3 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension3 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement3 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{300F7853-C0C5-4B5E-AA3D-9BA084EFB62B}\" />");

            // nonVisualDrawingPropertiesExtension3.Append(openXmlUnknownElement3);

            nonVisualDrawingPropertiesExtensionList3.Append(nonVisualDrawingPropertiesExtension3);

            nonVisualDrawingProperties4.Append(nonVisualDrawingPropertiesExtensionList3);
            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new NonVisualPictureDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties4 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties4);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);
            nonVisualPictureProperties1.Append(applicationNonVisualDrawingProperties4);

            BlipFill blipFill1 = new BlipFill();
            A.Blip blip1 = new A.Blip() { Embed = "rId2" };

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            ShapeProperties shapeProperties3 = new ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset2 = new A.Offset() { X = 12241L, Y = -14682L };
            A.Extents extents2 = new A.Extents() { Cx = 12188190L, Cy = 6860144L };

            transform2D1.Append(offset2);
            transform2D1.Append(extents2);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties3.Append(transform2D1);
            shapeProperties3.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties3);

            Shape shape3 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties3 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties5 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Shape 125" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList4 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension4 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement4 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{79A49323-3704-7849-89E5-774ED6647856}\" />");

            // nonVisualDrawingPropertiesExtension4.Append(openXmlUnknownElement4);

            nonVisualDrawingPropertiesExtensionList4.Append(nonVisualDrawingPropertiesExtension4);

            nonVisualDrawingProperties5.Append(nonVisualDrawingPropertiesExtensionList4);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties3 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties5 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties3.Append(nonVisualDrawingProperties5);
            nonVisualShapeProperties3.Append(nonVisualShapeDrawingProperties3);
            nonVisualShapeProperties3.Append(applicationNonVisualDrawingProperties5);

            ShapeProperties shapeProperties4 = new ShapeProperties();

            A.Transform2D transform2D2 = new A.Transform2D();
            A.Offset offset3 = new A.Offset() { X = 686195L, Y = 697634L };
            A.Extents extents3 = new A.Extents() { Cx = 10852432L, Cy = 365713L };

            transform2D2.Append(offset3);
            transform2D2.Append(extents3);

            A.PresetGeometry presetGeometry2 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList2 = new A.AdjustValueList();

            presetGeometry2.Append(adjustValueList2);
            A.NoFill noFill1 = new A.NoFill();
            A.Outline outline1 = new A.Outline();

            shapeProperties4.Append(transform2D2);
            shapeProperties4.Append(presetGeometry2);
            shapeProperties4.Append(noFill1);
            shapeProperties4.Append(outline1);

            TextBody textBody3 = new TextBody();
            A.BodyProperties bodyProperties3 = new A.BodyProperties();
            A.ListStyle listStyle3 = new A.ListStyle();

            A.Paragraph paragraph3 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1620 };

            paragraph3.Append(endParagraphRunProperties3);

            textBody3.Append(bodyProperties3);
            textBody3.Append(listStyle3);
            textBody3.Append(paragraph3);

            shape3.Append(nonVisualShapeProperties3);
            shape3.Append(shapeProperties4);
            shape3.Append(textBody3);

            Shape shape4 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties4 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties6 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Text 126" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList5 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension5 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement5 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{93CFC7DA-1334-BDE3-62FB-264262A7994F}\" />");

            // nonVisualDrawingPropertiesExtension5.Append(openXmlUnknownElement5);

            nonVisualDrawingPropertiesExtensionList5.Append(nonVisualDrawingPropertiesExtension5);

            nonVisualDrawingProperties6.Append(nonVisualDrawingPropertiesExtensionList5);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties4 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties6 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties4.Append(nonVisualDrawingProperties6);
            nonVisualShapeProperties4.Append(nonVisualShapeDrawingProperties4);
            nonVisualShapeProperties4.Append(applicationNonVisualDrawingProperties6);

            ShapeProperties shapeProperties5 = new ShapeProperties();

            A.Transform2D transform2D3 = new A.Transform2D();
            A.Offset offset4 = new A.Offset() { X = 686194L, Y = 495574L };
            A.Extents extents4 = new A.Extents() { Cx = 8639872L, Cy = 457142L };

            transform2D3.Append(offset4);
            transform2D3.Append(extents4);

            A.PresetGeometry presetGeometry3 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList3 = new A.AdjustValueList();

            presetGeometry3.Append(adjustValueList3);
            A.NoFill noFill2 = new A.NoFill();
            A.Outline outline2 = new A.Outline();

            shapeProperties5.Append(transform2D3);
            shapeProperties5.Append(presetGeometry3);
            shapeProperties5.Append(noFill2);
            shapeProperties5.Append(outline2);

            TextBody textBody4 = new TextBody();
            A.BodyProperties bodyProperties4 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle4 = new A.ListStyle();

            A.Paragraph paragraph4 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing1 = new A.LineSpacing();
            A.SpacingPercent spacingPercent1 = new A.SpacingPercent() { Val = 85744 };

            lineSpacing1.Append(spacingPercent1);

            paragraphProperties1.Append(lineSpacing1);

            A.Run run1 = new A.Run();

            A.RunProperties runProperties1 = new A.RunProperties() { Language = "en-US", FontSize = 2000 };

            A.SolidFill solidFill10 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "242424" };
            A.Alpha alpha1 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex1.Append(alpha1);

            solidFill10.Append(rgbColorModelHex1);
            A.LatinFont latinFont10 = new A.LatinFont() { Typeface = "Segoe UI Semibold", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont10 = new A.EastAsianFont() { Typeface = "Segoe UI Semibold", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont10 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", PitchFamily = 34, CharacterSet = -120 };

            runProperties1.Append(solidFill10);
            runProperties1.Append(latinFont10);
            runProperties1.Append(eastAsianFont10);
            runProperties1.Append(complexScriptFont10);
            A.Text text1 = new A.Text();
            text1.Text = "Monthly Overview";

            run1.Append(runProperties1);
            run1.Append(text1);
            A.EndParagraphRunProperties endParagraphRunProperties4 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 2000 };

            paragraph4.Append(paragraphProperties1);
            paragraph4.Append(run1);
            paragraph4.Append(endParagraphRunProperties4);

            textBody4.Append(bodyProperties4);
            textBody4.Append(listStyle4);
            textBody4.Append(paragraph4);

            shape4.Append(nonVisualShapeProperties4);
            shape4.Append(shapeProperties5);
            shape4.Append(textBody4);

            Shape shape5 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties5 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties7 = new NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Text 52" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList6 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension6 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement6 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{A81CF8E3-0752-ABE6-88F6-6E9DCCB62A60}\" />");

            // nonVisualDrawingPropertiesExtension6.Append(openXmlUnknownElement6);

            nonVisualDrawingPropertiesExtensionList6.Append(nonVisualDrawingPropertiesExtension6);

            nonVisualDrawingProperties7.Append(nonVisualDrawingPropertiesExtensionList6);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties5 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties7 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties5.Append(nonVisualDrawingProperties7);
            nonVisualShapeProperties5.Append(nonVisualShapeDrawingProperties5);
            nonVisualShapeProperties5.Append(applicationNonVisualDrawingProperties7);

            ShapeProperties shapeProperties6 = new ShapeProperties();

            A.Transform2D transform2D4 = new A.Transform2D();
            A.Offset offset5 = new A.Offset() { X = 663320L, Y = 6469053L };
            A.Extents extents5 = new A.Extents() { Cx = 3888430L, Cy = 228571L };

            transform2D4.Append(offset5);
            transform2D4.Append(extents5);

            A.PresetGeometry presetGeometry4 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList4 = new A.AdjustValueList();

            presetGeometry4.Append(adjustValueList4);
            A.NoFill noFill3 = new A.NoFill();
            A.Outline outline3 = new A.Outline();

            shapeProperties6.Append(transform2D4);
            shapeProperties6.Append(presetGeometry4);
            shapeProperties6.Append(noFill3);
            shapeProperties6.Append(outline3);

            TextBody textBody5 = new TextBody();
            A.BodyProperties bodyProperties5 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle5 = new A.ListStyle();

            A.Paragraph paragraph5 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing2 = new A.LineSpacing();
            A.SpacingPercent spacingPercent2 = new A.SpacingPercent() { Val = 85714 };

            lineSpacing2.Append(spacingPercent2);

            paragraphProperties2.Append(lineSpacing2);

            A.Run run2 = new A.Run();

            A.RunProperties runProperties2 = new A.RunProperties() { Language = "en-US", FontSize = 1000 };

            A.SolidFill solidFill11 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "616161" };
            A.Alpha alpha2 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex2.Append(alpha2);

            solidFill11.Append(rgbColorModelHex2);
            A.LatinFont latinFont11 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont11 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont11 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties2.Append(solidFill11);
            runProperties2.Append(latinFont11);
            runProperties2.Append(eastAsianFont11);
            runProperties2.Append(complexScriptFont11);
            A.Text text2 = new A.Text();

            // ****************************** Custom Code Changes Start *************************************
            DateTime today = DateTime.Today;
            string formattedDate = today.ToString("MMMM d, yyyy");
            text2.Text = "Exported from Viva Goals on " + formattedDate + ".";
            // ****************************** Custom Code Changes End *************************************

            run2.Append(runProperties2);
            run2.Append(text2);

            A.Run run3 = new A.Run();

            A.RunProperties runProperties3 = new A.RunProperties() { Language = "en-US", FontSize = 1000 };

            A.SolidFill solidFill12 = new A.SolidFill();

            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset1 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor10.Append(luminanceModulation1);
            schemeColor10.Append(luminanceOffset1);

            solidFill12.Append(schemeColor10);
            A.LatinFont latinFont12 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont12 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont12 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties3.Append(solidFill12);
            runProperties3.Append(latinFont12);
            runProperties3.Append(eastAsianFont12);
            runProperties3.Append(complexScriptFont12);
            A.Text text3 = new A.Text();
            text3.Text = ". ";

            run3.Append(runProperties3);
            run3.Append(text3);

            A.Run run4 = new A.Run();

            A.RunProperties runProperties4 = new A.RunProperties() { Language = "en-US", FontSize = 1000, Underline = A.TextUnderlineValues.Single };

            A.SolidFill solidFill13 = new A.SolidFill();

            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset2 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor11.Append(luminanceModulation2);
            schemeColor11.Append(luminanceOffset2);

            solidFill13.Append(schemeColor11);
            A.LatinFont latinFont13 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont13 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont13 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties4.Append(solidFill13);
            runProperties4.Append(latinFont13);
            runProperties4.Append(eastAsianFont13);
            runProperties4.Append(complexScriptFont13);
            A.Text text4 = new A.Text();
            text4.Text = "Open in Viva Goals";

            run4.Append(runProperties4);
            run4.Append(text4);

            A.EndParagraphRunProperties endParagraphRunProperties5 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1000, Underline = A.TextUnderlineValues.Single };

            A.SolidFill solidFill14 = new A.SolidFill();

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset3 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor12.Append(luminanceModulation3);
            schemeColor12.Append(luminanceOffset3);

            solidFill14.Append(schemeColor12);

            endParagraphRunProperties5.Append(solidFill14);

            paragraph5.Append(paragraphProperties2);
            paragraph5.Append(run2);
            paragraph5.Append(run3);
            paragraph5.Append(run4);
            paragraph5.Append(endParagraphRunProperties5);

            textBody5.Append(bodyProperties5);
            textBody5.Append(listStyle5);
            textBody5.Append(paragraph5);

            shape5.Append(nonVisualShapeProperties5);
            shape5.Append(shapeProperties6);
            shape5.Append(textBody5);

            Picture picture2 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties2 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties8 = new NonVisualDrawingProperties() { Id = (UInt32Value)8U, Name = "Picture 7" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList7 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension7 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement7 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{C6C12BA7-8B8F-A0F0-D233-674EAFD19E93}\" />");

            // nonVisualDrawingPropertiesExtension7.Append(openXmlUnknownElement7);

            nonVisualDrawingPropertiesExtensionList7.Append(nonVisualDrawingPropertiesExtension7);

            nonVisualDrawingProperties8.Append(nonVisualDrawingPropertiesExtensionList7);
            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties2 = new NonVisualPictureDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties8 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties2.Append(nonVisualDrawingProperties8);
            nonVisualPictureProperties2.Append(nonVisualPictureDrawingProperties2);
            nonVisualPictureProperties2.Append(applicationNonVisualDrawingProperties8);

            BlipFill blipFill2 = new BlipFill();
            A.Blip blip2 = new A.Blip() { Embed = "rId3" };

            A.Stretch stretch2 = new A.Stretch();
            A.FillRectangle fillRectangle2 = new A.FillRectangle();

            stretch2.Append(fillRectangle2);

            blipFill2.Append(blip2);
            blipFill2.Append(stretch2);

            ShapeProperties shapeProperties7 = new ShapeProperties();

            A.Transform2D transform2D5 = new A.Transform2D();
            A.Offset offset6 = new A.Offset() { X = 223167L, Y = 223408L };
            A.Extents extents6 = new A.Extents() { Cx = 203200L, Cy = 203200L };

            transform2D5.Append(offset6);
            transform2D5.Append(extents6);

            A.PresetGeometry presetGeometry5 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList5 = new A.AdjustValueList();

            presetGeometry5.Append(adjustValueList5);

            shapeProperties7.Append(transform2D5);
            shapeProperties7.Append(presetGeometry5);

            picture2.Append(nonVisualPictureProperties2);
            picture2.Append(blipFill2);
            picture2.Append(shapeProperties7);

            GraphicFrame graphicFrame1 = new GraphicFrame();

            NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = new NonVisualGraphicFrameProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties9 = new NonVisualDrawingProperties() { Id = (UInt32Value)9U, Name = "Table 8" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList8 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension8 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement8 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{344B74EF-5D43-C816-2976-B1653E735D33}\" />");

            // nonVisualDrawingPropertiesExtension8.Append(openXmlUnknownElement8);

            nonVisualDrawingPropertiesExtensionList8.Append(nonVisualDrawingPropertiesExtension8);

            nonVisualDrawingProperties9.Append(nonVisualDrawingPropertiesExtensionList8);

            NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoGrouping = true };

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties9 = new ApplicationNonVisualDrawingProperties();

            nonVisualGraphicFrameProperties1.Append(nonVisualDrawingProperties9);
            nonVisualGraphicFrameProperties1.Append(nonVisualGraphicFrameDrawingProperties1);
            nonVisualGraphicFrameProperties1.Append(applicationNonVisualDrawingProperties9);

            Transform transform1 = new Transform();
            A.Offset offset7 = new A.Offset() { X = 663320L, Y = 1055367L };
            A.Extents extents7 = new A.Extents() { Cx = 10729953L, Cy = 5187778L };

            transform1.Append(offset7);
            transform1.Append(extents7);

            A.Graphic graphic1 = new A.Graphic();

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" };

            A.Table table1 = new A.Table();

            A.TableProperties tableProperties1 = new A.TableProperties() { FirstRow = true, BandRow = true };

            A.EffectList effectList1 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 38100L, HorizontalRatio = 101000, VerticalRatio = 101000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.PresetColor presetColor1 = new A.PresetColor() { Val = A.PresetColorValues.Black };
            A.Alpha alpha3 = new A.Alpha() { Val = 3000 };

            presetColor1.Append(alpha3);

            outerShadow1.Append(presetColor1);

            effectList1.Append(outerShadow1);
            A.TableStyleId tableStyleId1 = new A.TableStyleId();
            tableStyleId1.Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";

            tableProperties1.Append(effectList1);
            tableProperties1.Append(tableStyleId1);

            A.TableGrid tableGrid1 = new A.TableGrid();

            A.GridColumn gridColumn1 = new A.GridColumn() { Width = 7646290L };

            A.ExtensionList extensionList1 = new A.ExtensionList();

            A.Extension extension1 = new A.Extension() { Uri = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement9 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:colId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"3835035813\" />");

            // extension1.Append(openXmlUnknownElement9);

            extensionList1.Append(extension1);

            gridColumn1.Append(extensionList1);

            A.GridColumn gridColumn2 = new A.GridColumn() { Width = 1828800L };

            A.ExtensionList extensionList2 = new A.ExtensionList();

            A.Extension extension2 = new A.Extension() { Uri = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement10 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:colId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"622552574\" />");

            // extension2.Append(openXmlUnknownElement10);

            extensionList2.Append(extension2);

            gridColumn2.Append(extensionList2);

            A.GridColumn gridColumn3 = new A.GridColumn() { Width = 1254863L };

            A.ExtensionList extensionList3 = new A.ExtensionList();

            A.Extension extension3 = new A.Extension() { Uri = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement11 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:colId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"3617387509\" />");

            // extension3.Append(openXmlUnknownElement11);

            extensionList3.Append(extension3);

            gridColumn3.Append(extensionList3);

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);

            A.TableRow tableRow1 = new A.TableRow() { Height = 475571L };

            A.TableCell tableCell1 = new A.TableCell();

            A.TextBody textBody6 = new A.TextBody();
            A.BodyProperties bodyProperties6 = new A.BodyProperties();
            A.ListStyle listStyle6 = new A.ListStyle();

            A.Paragraph paragraph6 = new A.Paragraph();

            A.Run run5 = new A.Run();

            A.RunProperties runProperties5 = new A.RunProperties() { Language = "en-US", FontSize = 1000, Bold = false, Italic = false, Dirty = false };

            A.SolidFill solidFill15 = new A.SolidFill();
            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill15.Append(schemeColor13);
            A.LatinFont latinFont14 = new A.LatinFont() { Typeface = "Segoe UI" };
            A.ComplexScriptFont complexScriptFont14 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            runProperties5.Append(solidFill15);
            runProperties5.Append(latinFont14);
            runProperties5.Append(complexScriptFont14);
            A.Text text5 = new A.Text();
            text5.Text = "Goals";

            run5.Append(runProperties5);
            run5.Append(text5);

            paragraph6.Append(run5);

            textBody6.Append(bodyProperties6);
            textBody6.Append(listStyle6);
            textBody6.Append(paragraph6);

            A.TableCellProperties tableCellProperties1 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties1 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill4 = new A.NoFill();

            leftBorderLineProperties1.Append(noFill4);

            A.RightBorderLineProperties rightBorderLineProperties1 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill5 = new A.NoFill();

            rightBorderLineProperties1.Append(noFill5);

            A.TopBorderLineProperties topBorderLineProperties1 = new A.TopBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill6 = new A.NoFill();

            topBorderLineProperties1.Append(noFill6);

            A.BottomBorderLineProperties bottomBorderLineProperties1 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill7 = new A.NoFill();
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round1 = new A.Round();
            A.HeadEnd headEnd1 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd1 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties1.Append(noFill7);
            bottomBorderLineProperties1.Append(presetDash1);
            bottomBorderLineProperties1.Append(round1);
            bottomBorderLineProperties1.Append(headEnd1);
            bottomBorderLineProperties1.Append(tailEnd1);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties1 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill8 = new A.NoFill();
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties1.Append(noFill8);
            topLeftToBottomRightBorderLineProperties1.Append(presetDash2);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties1 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill9 = new A.NoFill();
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties1.Append(noFill9);
            bottomLeftToTopRightBorderLineProperties1.Append(presetDash3);

            A.SolidFill solidFill16 = new A.SolidFill();
            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill16.Append(schemeColor14);

            tableCellProperties1.Append(leftBorderLineProperties1);
            tableCellProperties1.Append(rightBorderLineProperties1);
            tableCellProperties1.Append(topBorderLineProperties1);
            tableCellProperties1.Append(bottomBorderLineProperties1);
            tableCellProperties1.Append(topLeftToBottomRightBorderLineProperties1);
            tableCellProperties1.Append(bottomLeftToTopRightBorderLineProperties1);
            tableCellProperties1.Append(solidFill16);

            tableCell1.Append(textBody6);
            tableCell1.Append(tableCellProperties1);

            A.TableCell tableCell2 = new A.TableCell();

            A.TextBody textBody7 = new A.TextBody();
            A.BodyProperties bodyProperties7 = new A.BodyProperties();
            A.ListStyle listStyle7 = new A.ListStyle();

            A.Paragraph paragraph7 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing3 = new A.LineSpacing();
            A.SpacingPercent spacingPercent3 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing3.Append(spacingPercent3);

            A.SpaceBefore spaceBefore1 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints1 = new A.SpacingPoints() { Val = 0 };

            spaceBefore1.Append(spacingPoints1);

            A.SpaceAfter spaceAfter1 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints2 = new A.SpacingPoints() { Val = 0 };

            spaceAfter1.Append(spacingPoints2);
            A.BulletColorText bulletColorText1 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText1 = new A.BulletSizeText();
            A.BulletFontText bulletFontText1 = new A.BulletFontText();
            A.NoBullet noBullet1 = new A.NoBullet();
            A.TabStopList tabStopList1 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties11 = new A.DefaultRunProperties();

            paragraphProperties3.Append(lineSpacing3);
            paragraphProperties3.Append(spaceBefore1);
            paragraphProperties3.Append(spaceAfter1);
            paragraphProperties3.Append(bulletColorText1);
            paragraphProperties3.Append(bulletSizeText1);
            paragraphProperties3.Append(bulletFontText1);
            paragraphProperties3.Append(noBullet1);
            paragraphProperties3.Append(tabStopList1);
            paragraphProperties3.Append(defaultRunProperties11);

            A.Run run6 = new A.Run();

            A.RunProperties runProperties6 = new A.RunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1000, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline4 = new A.Outline();
            A.NoFill noFill10 = new A.NoFill();

            outline4.Append(noFill10);

            A.SolidFill solidFill17 = new A.SolidFill();
            A.PresetColor presetColor2 = new A.PresetColor() { Val = A.PresetColorValues.Black };

            solidFill17.Append(presetColor2);
            A.EffectList effectList2 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText1 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText1 = new A.UnderlineFillText();
            A.LatinFont latinFont15 = new A.LatinFont() { Typeface = "Segoe UI" };
            A.EastAsianFont eastAsianFont14 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont15 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            runProperties6.Append(outline4);
            runProperties6.Append(solidFill17);
            runProperties6.Append(effectList2);
            runProperties6.Append(underlineFollowsText1);
            runProperties6.Append(underlineFillText1);
            runProperties6.Append(latinFont15);
            runProperties6.Append(eastAsianFont14);
            runProperties6.Append(complexScriptFont15);
            A.Text text6 = new A.Text();
            text6.Text = "Owner";

            run6.Append(runProperties6);
            run6.Append(text6);

            paragraph7.Append(paragraphProperties3);
            paragraph7.Append(run6);

            textBody7.Append(bodyProperties7);
            textBody7.Append(listStyle7);
            textBody7.Append(paragraph7);

            A.TableCellProperties tableCellProperties2 = new A.TableCellProperties() { LeftMargin = 82296, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties2 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill11 = new A.NoFill();

            leftBorderLineProperties2.Append(noFill11);

            A.RightBorderLineProperties rightBorderLineProperties2 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill12 = new A.NoFill();

            rightBorderLineProperties2.Append(noFill12);

            A.TopBorderLineProperties topBorderLineProperties2 = new A.TopBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill13 = new A.NoFill();

            topBorderLineProperties2.Append(noFill13);

            A.BottomBorderLineProperties bottomBorderLineProperties2 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill14 = new A.NoFill();
            A.PresetDash presetDash4 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round2 = new A.Round();
            A.HeadEnd headEnd2 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd2 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties2.Append(noFill14);
            bottomBorderLineProperties2.Append(presetDash4);
            bottomBorderLineProperties2.Append(round2);
            bottomBorderLineProperties2.Append(headEnd2);
            bottomBorderLineProperties2.Append(tailEnd2);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties2 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill15 = new A.NoFill();
            A.PresetDash presetDash5 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties2.Append(noFill15);
            topLeftToBottomRightBorderLineProperties2.Append(presetDash5);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties2 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill16 = new A.NoFill();
            A.PresetDash presetDash6 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties2.Append(noFill16);
            bottomLeftToTopRightBorderLineProperties2.Append(presetDash6);

            A.SolidFill solidFill18 = new A.SolidFill();
            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill18.Append(schemeColor15);

            tableCellProperties2.Append(leftBorderLineProperties2);
            tableCellProperties2.Append(rightBorderLineProperties2);
            tableCellProperties2.Append(topBorderLineProperties2);
            tableCellProperties2.Append(bottomBorderLineProperties2);
            tableCellProperties2.Append(topLeftToBottomRightBorderLineProperties2);
            tableCellProperties2.Append(bottomLeftToTopRightBorderLineProperties2);
            tableCellProperties2.Append(solidFill18);

            tableCell2.Append(textBody7);
            tableCell2.Append(tableCellProperties2);

            A.TableCell tableCell3 = new A.TableCell();

            A.TextBody textBody8 = new A.TextBody();
            A.BodyProperties bodyProperties8 = new A.BodyProperties();
            A.ListStyle listStyle8 = new A.ListStyle();

            A.Paragraph paragraph8 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing4 = new A.LineSpacing();
            A.SpacingPercent spacingPercent4 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing4.Append(spacingPercent4);

            A.SpaceBefore spaceBefore2 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints3 = new A.SpacingPoints() { Val = 0 };

            spaceBefore2.Append(spacingPoints3);

            A.SpaceAfter spaceAfter2 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints4 = new A.SpacingPoints() { Val = 0 };

            spaceAfter2.Append(spacingPoints4);
            A.BulletColorText bulletColorText2 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText2 = new A.BulletSizeText();
            A.BulletFontText bulletFontText2 = new A.BulletFontText();
            A.NoBullet noBullet2 = new A.NoBullet();
            A.TabStopList tabStopList2 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties12 = new A.DefaultRunProperties();

            paragraphProperties4.Append(lineSpacing4);
            paragraphProperties4.Append(spaceBefore2);
            paragraphProperties4.Append(spaceAfter2);
            paragraphProperties4.Append(bulletColorText2);
            paragraphProperties4.Append(bulletSizeText2);
            paragraphProperties4.Append(bulletFontText2);
            paragraphProperties4.Append(noBullet2);
            paragraphProperties4.Append(tabStopList2);
            paragraphProperties4.Append(defaultRunProperties12);

            A.Run run7 = new A.Run();

            A.RunProperties runProperties7 = new A.RunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1000, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline5 = new A.Outline();
            A.NoFill noFill17 = new A.NoFill();

            outline5.Append(noFill17);

            A.SolidFill solidFill19 = new A.SolidFill();
            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill19.Append(schemeColor16);
            A.EffectList effectList3 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText2 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText2 = new A.UnderlineFillText();
            A.LatinFont latinFont16 = new A.LatinFont() { Typeface = "Segoe UI" };
            A.EastAsianFont eastAsianFont15 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont16 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            runProperties7.Append(outline5);
            runProperties7.Append(solidFill19);
            runProperties7.Append(effectList3);
            runProperties7.Append(underlineFollowsText2);
            runProperties7.Append(underlineFillText2);
            runProperties7.Append(latinFont16);
            runProperties7.Append(eastAsianFont15);
            runProperties7.Append(complexScriptFont16);
            A.Text text7 = new A.Text();
            text7.Text = "Target";

            run7.Append(runProperties7);
            run7.Append(text7);

            paragraph8.Append(paragraphProperties4);
            paragraph8.Append(run7);

            textBody8.Append(bodyProperties8);
            textBody8.Append(listStyle8);
            textBody8.Append(paragraph8);

            A.TableCellProperties tableCellProperties3 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties3 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill18 = new A.NoFill();

            leftBorderLineProperties3.Append(noFill18);

            A.RightBorderLineProperties rightBorderLineProperties3 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill19 = new A.NoFill();

            rightBorderLineProperties3.Append(noFill19);

            A.TopBorderLineProperties topBorderLineProperties3 = new A.TopBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill20 = new A.NoFill();

            topBorderLineProperties3.Append(noFill20);

            A.BottomBorderLineProperties bottomBorderLineProperties3 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill21 = new A.NoFill();
            A.PresetDash presetDash7 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round3 = new A.Round();
            A.HeadEnd headEnd3 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd3 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties3.Append(noFill21);
            bottomBorderLineProperties3.Append(presetDash7);
            bottomBorderLineProperties3.Append(round3);
            bottomBorderLineProperties3.Append(headEnd3);
            bottomBorderLineProperties3.Append(tailEnd3);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties3 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill22 = new A.NoFill();
            A.PresetDash presetDash8 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties3.Append(noFill22);
            topLeftToBottomRightBorderLineProperties3.Append(presetDash8);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties3 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill23 = new A.NoFill();
            A.PresetDash presetDash9 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties3.Append(noFill23);
            bottomLeftToTopRightBorderLineProperties3.Append(presetDash9);

            A.SolidFill solidFill20 = new A.SolidFill();
            A.SchemeColor schemeColor17 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill20.Append(schemeColor17);

            tableCellProperties3.Append(leftBorderLineProperties3);
            tableCellProperties3.Append(rightBorderLineProperties3);
            tableCellProperties3.Append(topBorderLineProperties3);
            tableCellProperties3.Append(bottomBorderLineProperties3);
            tableCellProperties3.Append(topLeftToBottomRightBorderLineProperties3);
            tableCellProperties3.Append(bottomLeftToTopRightBorderLineProperties3);
            tableCellProperties3.Append(solidFill20);

            tableCell3.Append(textBody8);
            tableCell3.Append(tableCellProperties3);

            A.ExtensionList extensionList4 = new A.ExtensionList();

            A.Extension extension4 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement12 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"2429191634\" />");

            // extension4.Append(openXmlUnknownElement12);

            extensionList4.Append(extension4);

            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);
            tableRow1.Append(extensionList4);

            A.TableRow tableRow2 = new A.TableRow() { Height = 781011L };

            A.TableCell tableCell4 = new A.TableCell();

            A.TextBody textBody1001 = new A.TextBody();
            A.BodyProperties bodyProperties1001 = new A.BodyProperties();
            A.ListStyle listStyle1001 = new A.ListStyle();

            A.Paragraph paragraph1001 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties1001 = new A.ParagraphProperties() { LeftMargin = 171450, RightMargin = 0, Indent = -171450, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing1001 = new A.LineSpacing();
            A.SpacingPercent spacingPercent1001 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing1001.Append(spacingPercent1001);

            A.SpaceBefore spaceBefore1001 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints1001 = new A.SpacingPoints() { Val = 0 };

            spaceBefore1001.Append(spacingPoints1001);

            A.SpaceAfter spaceAfter1001 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints1002 = new A.SpacingPoints() { Val = 0 };

            spaceAfter1001.Append(spacingPoints1002);
            A.BulletColorText bulletColorText1001 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText1001 = new A.BulletSizeText();
            A.BulletFontText bulletFontText1001 = new A.BulletFontText();

            A.PictureBullet pictureBullet1001 = new A.PictureBullet();

            A.Blip blip1001 = new A.Blip() { Embed = "rId4" };

            A.BlipExtensionList blipExtensionList1001 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1001 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            blipExtensionList1001.Append(blipExtension1001);

            blip1001.Append(blipExtensionList1001);

            pictureBullet1001.Append(blip1001);
            A.TabStopList tabStopList1001 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties1001 = new A.DefaultRunProperties();

            paragraphProperties1001.Append(lineSpacing1001);
            paragraphProperties1001.Append(spaceBefore1001);
            paragraphProperties1001.Append(spaceAfter1001);
            paragraphProperties1001.Append(bulletColorText1001);
            paragraphProperties1001.Append(bulletSizeText1001);
            paragraphProperties1001.Append(bulletFontText1001);
            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count > 0 && !goalsArrayForSlide.ElementAt(0).isMainGoal)
                paragraphProperties1001.Append(pictureBullet1001);
            // ********************************* Custom Code Changes End *********************************
            paragraphProperties1001.Append(tabStopList1001);
            paragraphProperties1001.Append(defaultRunProperties1001);

            A.Run run1001 = new A.Run();

            // ********************************* Custom Code Changes Start *********************************
            int fontSize1001 = goalsArrayForSlide.Count > 0 && goalsArrayForSlide.ElementAt(0).isMainGoal ? 1300 : 1200;
            bool isBold1001 = goalsArrayForSlide.Count > 0 && goalsArrayForSlide.ElementAt(0).isMainGoal;
            string fontName1001 = goalsArrayForSlide.Count > 0 && goalsArrayForSlide.ElementAt(0).isMainGoal ? "Segoe UI Semibold" : "Segoe UI";
            // ********************************* Custom Code Changes End *********************************

            A.RunProperties runProperties1001 = new A.RunProperties() { Language = "en-US", FontSize = fontSize1001, Bold = isBold1001, Italic = false };
            A.EffectList effectList1001 = new A.EffectList();
            A.LatinFont latinFont1001 = new A.LatinFont() { Typeface = fontName1001 };
            A.ComplexScriptFont complexScriptFont1001 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            runProperties1001.Append(effectList1001);
            runProperties1001.Append(latinFont1001);
            runProperties1001.Append(complexScriptFont1001);
            A.Text text1001 = new A.Text();

            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count() > 0)
                text1001.Text = goalsArrayForSlide.ElementAt(0).goalName ?? "";
            else
                text1001.Text = "NO DATA";
            // ********************************* Custom Code Changes End *********************************

            run1001.Append(runProperties1001);
            run1001.Append(text1001);

            A.EndParagraphRunProperties endParagraphRunProperties1001 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false };
            A.LatinFont latinFont1002 = new A.LatinFont() { Typeface = "Segoe UI" };
            A.ComplexScriptFont complexScriptFont1002 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            endParagraphRunProperties1001.Append(latinFont1002);
            endParagraphRunProperties1001.Append(complexScriptFont1002);

            paragraph1001.Append(paragraphProperties1001);
            paragraph1001.Append(run1001);
            paragraph1001.Append(endParagraphRunProperties1001);

            textBody1001.Append(bodyProperties1001);
            textBody1001.Append(listStyle1001);
            textBody1001.Append(paragraph1001);

            A.TableCellProperties tableCellProperties1001 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties1001 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill1001 = new A.NoFill();

            leftBorderLineProperties1001.Append(noFill1001);

            A.RightBorderLineProperties rightBorderLineProperties1001 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill1002 = new A.NoFill();

            rightBorderLineProperties1001.Append(noFill1002);

            A.TopBorderLineProperties topBorderLineProperties1001 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill1003 = new A.NoFill();
            A.PresetDash presetDash1001 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round1001 = new A.Round();
            A.HeadEnd headEnd1001 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd1001 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties1001.Append(noFill1003);
            topBorderLineProperties1001.Append(presetDash1001);
            topBorderLineProperties1001.Append(round1001);
            topBorderLineProperties1001.Append(headEnd1001);
            topBorderLineProperties1001.Append(tailEnd1001);

            A.BottomBorderLineProperties bottomBorderLineProperties1001 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill1004 = new A.NoFill();
            A.PresetDash presetDash1002 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round1002 = new A.Round();
            A.HeadEnd headEnd1002 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd1002 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties1001.Append(noFill1004);
            bottomBorderLineProperties1001.Append(presetDash1002);
            bottomBorderLineProperties1001.Append(round1002);
            bottomBorderLineProperties1001.Append(headEnd1002);
            bottomBorderLineProperties1001.Append(tailEnd1002);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties1001 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill1005 = new A.NoFill();
            A.PresetDash presetDash1003 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties1001.Append(noFill1005);
            topLeftToBottomRightBorderLineProperties1001.Append(presetDash1003);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties1001 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill1006 = new A.NoFill();
            A.PresetDash presetDash1004 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties1001.Append(noFill1006);
            bottomLeftToTopRightBorderLineProperties1001.Append(presetDash1004);

            A.SolidFill solidFill1001 = new A.SolidFill();

            // ********************************* Custom Code Changes Start *********************************
            // Adding the background color according to the goal2 type
            string goalColor1001 = goalsArrayForSlide.Count > 0 && goalsArrayForSlide.ElementAt(0).isMainGoal ? "F5F5F5" : "FFFFFF";
            A.RgbColorModelHex rgbColorModelHex1001 = new A.RgbColorModelHex() { Val = goalColor1001 };
            // ********************************* Custom Code Changes End *********************************

            solidFill1001.Append(rgbColorModelHex1001);

            tableCellProperties1001.Append(leftBorderLineProperties1001);
            tableCellProperties1001.Append(rightBorderLineProperties1001);
            tableCellProperties1001.Append(topBorderLineProperties1001);
            tableCellProperties1001.Append(bottomBorderLineProperties1001);
            tableCellProperties1001.Append(topLeftToBottomRightBorderLineProperties1001);
            tableCellProperties1001.Append(bottomLeftToTopRightBorderLineProperties1001);
            tableCellProperties1001.Append(solidFill1001);

            tableCell4.Append(textBody1001);
            tableCell4.Append(tableCellProperties1001);


            A.TableCell tableCell5 = new A.TableCell();

            A.TextBody textBody10 = new A.TextBody();
            A.BodyProperties bodyProperties10 = new A.BodyProperties();
            A.ListStyle listStyle10 = new A.ListStyle();

            A.Paragraph paragraph10 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing5 = new A.LineSpacing();
            A.SpacingPercent spacingPercent5 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing5.Append(spacingPercent5);

            A.SpaceBefore spaceBefore3 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints5 = new A.SpacingPoints() { Val = 0 };

            spaceBefore3.Append(spacingPoints5);

            A.SpaceAfter spaceAfter3 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints6 = new A.SpacingPoints() { Val = 0 };

            spaceAfter3.Append(spacingPoints6);
            A.BulletColorText bulletColorText3 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText3 = new A.BulletSizeText();
            A.BulletFontText bulletFontText3 = new A.BulletFontText();
            A.NoBullet noBullet3 = new A.NoBullet();
            A.TabStopList tabStopList3 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties13 = new A.DefaultRunProperties();

            paragraphProperties5.Append(lineSpacing5);
            paragraphProperties5.Append(spaceBefore3);
            paragraphProperties5.Append(spaceAfter3);
            paragraphProperties5.Append(bulletColorText3);
            paragraphProperties5.Append(bulletSizeText3);
            paragraphProperties5.Append(bulletFontText3);
            paragraphProperties5.Append(noBullet3);
            paragraphProperties5.Append(tabStopList3);
            paragraphProperties5.Append(defaultRunProperties13);

            A.EndParagraphRunProperties endParagraphRunProperties7 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline6 = new A.Outline();
            A.NoFill noFill30 = new A.NoFill();

            outline6.Append(noFill30);

            A.SolidFill solidFill25 = new A.SolidFill();
            A.PresetColor presetColor3 = new A.PresetColor() { Val = A.PresetColorValues.Black };

            solidFill25.Append(presetColor3);
            A.EffectList effectList7 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText3 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText3 = new A.UnderlineFillText();
            A.LatinFont latinFont20 = new A.LatinFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont18 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont20 = new A.ComplexScriptFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties7.Append(outline6);
            endParagraphRunProperties7.Append(solidFill25);
            endParagraphRunProperties7.Append(effectList7);
            endParagraphRunProperties7.Append(underlineFollowsText3);
            endParagraphRunProperties7.Append(underlineFillText3);
            endParagraphRunProperties7.Append(latinFont20);
            endParagraphRunProperties7.Append(eastAsianFont18);
            endParagraphRunProperties7.Append(complexScriptFont20);

            paragraph10.Append(paragraphProperties5);
            paragraph10.Append(endParagraphRunProperties7);

            textBody10.Append(bodyProperties10);
            textBody10.Append(listStyle10);
            textBody10.Append(paragraph10);

            A.TableCellProperties tableCellProperties5 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties5 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill31 = new A.NoFill();

            leftBorderLineProperties5.Append(noFill31);

            A.RightBorderLineProperties rightBorderLineProperties5 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill32 = new A.NoFill();

            rightBorderLineProperties5.Append(noFill32);

            A.TopBorderLineProperties topBorderLineProperties5 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill33 = new A.NoFill();
            A.PresetDash presetDash14 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round6 = new A.Round();
            A.HeadEnd headEnd6 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd6 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties5.Append(noFill33);
            topBorderLineProperties5.Append(presetDash14);
            topBorderLineProperties5.Append(round6);
            topBorderLineProperties5.Append(headEnd6);
            topBorderLineProperties5.Append(tailEnd6);

            A.BottomBorderLineProperties bottomBorderLineProperties5 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill34 = new A.NoFill();
            A.PresetDash presetDash15 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round7 = new A.Round();
            A.HeadEnd headEnd7 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd7 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties5.Append(noFill34);
            bottomBorderLineProperties5.Append(presetDash15);
            bottomBorderLineProperties5.Append(round7);
            bottomBorderLineProperties5.Append(headEnd7);
            bottomBorderLineProperties5.Append(tailEnd7);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties5 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill35 = new A.NoFill();
            A.PresetDash presetDash16 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties5.Append(noFill35);
            topLeftToBottomRightBorderLineProperties5.Append(presetDash16);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties5 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill36 = new A.NoFill();
            A.PresetDash presetDash17 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties5.Append(noFill36);
            bottomLeftToTopRightBorderLineProperties5.Append(presetDash17);

            A.SolidFill solidFill1101 = new A.SolidFill();
            // ********************************* Custom Code Changes Start *********************************
            string colorCode1101 = goalsArrayForSlide.Count > 0 && goalsArrayForSlide.ElementAt(0).isMainGoal ? "F5F5F5" : "FFFFFF";
            // ********************************* Custom Code Changes End *********************************

            A.RgbColorModelHex rgbColorModelHex1101 = new A.RgbColorModelHex() { Val = colorCode1101 };

            solidFill1101.Append(rgbColorModelHex1101);

            tableCellProperties5.Append(leftBorderLineProperties5);
            tableCellProperties5.Append(rightBorderLineProperties5);
            tableCellProperties5.Append(topBorderLineProperties5);
            tableCellProperties5.Append(bottomBorderLineProperties5);
            tableCellProperties5.Append(topLeftToBottomRightBorderLineProperties5);
            tableCellProperties5.Append(bottomLeftToTopRightBorderLineProperties5);
            tableCellProperties5.Append(solidFill1101);

            tableCell5.Append(textBody10);
            tableCell5.Append(tableCellProperties5);

            A.TableCell tableCell6 = new A.TableCell();

            A.TextBody textBody11 = new A.TextBody();
            A.BodyProperties bodyProperties11 = new A.BodyProperties();
            A.ListStyle listStyle11 = new A.ListStyle();

            A.Paragraph paragraph11 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties6 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing6 = new A.LineSpacing();
            A.SpacingPercent spacingPercent6 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing6.Append(spacingPercent6);

            A.SpaceBefore spaceBefore4 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints7 = new A.SpacingPoints() { Val = 0 };

            spaceBefore4.Append(spacingPoints7);

            A.SpaceAfter spaceAfter4 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints8 = new A.SpacingPoints() { Val = 0 };

            spaceAfter4.Append(spacingPoints8);
            A.BulletColorText bulletColorText4 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText4 = new A.BulletSizeText();
            A.BulletFontText bulletFontText4 = new A.BulletFontText();
            A.NoBullet noBullet4 = new A.NoBullet();
            A.TabStopList tabStopList4 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties14 = new A.DefaultRunProperties();

            paragraphProperties6.Append(lineSpacing6);
            paragraphProperties6.Append(spaceBefore4);
            paragraphProperties6.Append(spaceAfter4);
            paragraphProperties6.Append(bulletColorText4);
            paragraphProperties6.Append(bulletSizeText4);
            paragraphProperties6.Append(bulletFontText4);
            paragraphProperties6.Append(noBullet4);
            paragraphProperties6.Append(tabStopList4);
            paragraphProperties6.Append(defaultRunProperties14);

            A.EndParagraphRunProperties endParagraphRunProperties8 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline7 = new A.Outline();
            A.NoFill noFill37 = new A.NoFill();

            outline7.Append(noFill37);

            A.SolidFill solidFill27 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "C50F1F" };

            solidFill27.Append(rgbColorModelHex5);
            A.EffectList effectList8 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText4 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText4 = new A.UnderlineFillText();
            A.LatinFont latinFont21 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont19 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont21 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties8.Append(outline7);
            endParagraphRunProperties8.Append(solidFill27);
            endParagraphRunProperties8.Append(effectList8);
            endParagraphRunProperties8.Append(underlineFollowsText4);
            endParagraphRunProperties8.Append(underlineFillText4);
            endParagraphRunProperties8.Append(latinFont21);
            endParagraphRunProperties8.Append(eastAsianFont19);
            endParagraphRunProperties8.Append(complexScriptFont21);

            paragraph11.Append(paragraphProperties6);
            paragraph11.Append(endParagraphRunProperties8);

            textBody11.Append(bodyProperties11);
            textBody11.Append(listStyle11);
            textBody11.Append(paragraph11);

            A.TableCellProperties tableCellProperties6 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties6 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill38 = new A.NoFill();

            leftBorderLineProperties6.Append(noFill38);

            A.RightBorderLineProperties rightBorderLineProperties6 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill39 = new A.NoFill();

            rightBorderLineProperties6.Append(noFill39);

            A.TopBorderLineProperties topBorderLineProperties6 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill40 = new A.NoFill();
            A.PresetDash presetDash18 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round8 = new A.Round();
            A.HeadEnd headEnd8 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd8 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties6.Append(noFill40);
            topBorderLineProperties6.Append(presetDash18);
            topBorderLineProperties6.Append(round8);
            topBorderLineProperties6.Append(headEnd8);
            topBorderLineProperties6.Append(tailEnd8);

            A.BottomBorderLineProperties bottomBorderLineProperties6 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill41 = new A.NoFill();
            A.PresetDash presetDash19 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round9 = new A.Round();
            A.HeadEnd headEnd9 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd9 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties6.Append(noFill41);
            bottomBorderLineProperties6.Append(presetDash19);
            bottomBorderLineProperties6.Append(round9);
            bottomBorderLineProperties6.Append(headEnd9);
            bottomBorderLineProperties6.Append(tailEnd9);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties6 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill42 = new A.NoFill();
            A.PresetDash presetDash20 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties6.Append(noFill42);
            topLeftToBottomRightBorderLineProperties6.Append(presetDash20);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties6 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill43 = new A.NoFill();
            A.PresetDash presetDash21 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties6.Append(noFill43);
            bottomLeftToTopRightBorderLineProperties6.Append(presetDash21);

            A.SolidFill solidFill1201 = new A.SolidFill();
            // ********************************* Custom Code Changes Start *********************************
            string colorCode1201 = goalsArrayForSlide.Count > 0 && goalsArrayForSlide.ElementAt(0).isMainGoal ? "F5F5F5" : "FFFFFF";
            // ********************************* Custom Code Changes End *********************************

            A.RgbColorModelHex rgbColorModelHex1201 = new A.RgbColorModelHex() { Val = colorCode1201 };

            solidFill1201.Append(rgbColorModelHex1201);

            tableCellProperties6.Append(leftBorderLineProperties6);
            tableCellProperties6.Append(rightBorderLineProperties6);
            tableCellProperties6.Append(topBorderLineProperties6);
            tableCellProperties6.Append(bottomBorderLineProperties6);
            tableCellProperties6.Append(topLeftToBottomRightBorderLineProperties6);
            tableCellProperties6.Append(bottomLeftToTopRightBorderLineProperties6);
            tableCellProperties6.Append(solidFill1201);

            tableCell6.Append(textBody11);
            tableCell6.Append(tableCellProperties6);

            A.ExtensionList extensionList5 = new A.ExtensionList();

            A.Extension extension5 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement13 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"587416212\" />");

            // extension5.Append(openXmlUnknownElement13);

            extensionList5.Append(extension5);

            tableRow2.Append(tableCell4);
            tableRow2.Append(tableCell5);
            tableRow2.Append(tableCell6);
            tableRow2.Append(extensionList5);

            A.TableRow tableRow3 = new A.TableRow() { Height = 748571L };

            A.TableCell tableCell7 = new A.TableCell();

            A.TextBody textBody2001 = new A.TextBody();
            A.BodyProperties bodyProperties2001 = new A.BodyProperties();
            A.ListStyle listStyle2001 = new A.ListStyle();

            A.Paragraph paragraph2001 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties2001 = new A.ParagraphProperties() { LeftMargin = 171450, RightMargin = 0, Indent = -171450, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing2001 = new A.LineSpacing();
            A.SpacingPercent spacingPercent2001 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing2001.Append(spacingPercent2001);

            A.SpaceBefore spaceBefore2001 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints2001 = new A.SpacingPoints() { Val = 0 };

            spaceBefore2001.Append(spacingPoints2001);

            A.SpaceAfter spaceAfter2001 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints2002 = new A.SpacingPoints() { Val = 0 };

            spaceAfter2001.Append(spacingPoints2002);
            A.BulletColorText bulletColorText2001 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText2001 = new A.BulletSizeText();
            A.BulletFontText bulletFontText2001 = new A.BulletFontText();

            A.PictureBullet pictureBullet2001 = new A.PictureBullet();

            A.Blip blip2001 = new A.Blip() { Embed = "rId4" };

            A.BlipExtensionList blipExtensionList2001 = new A.BlipExtensionList();

            A.BlipExtension blipExtension2001 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement2001 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId5\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension2001.Append(openXmlUnknownElement2001);

            blipExtensionList2001.Append(blipExtension2001);

            blip2001.Append(blipExtensionList2001);

            pictureBullet2001.Append(blip2001);
            A.TabStopList tabStopList2001 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties2001 = new A.DefaultRunProperties();

            paragraphProperties2001.Append(lineSpacing2001);
            paragraphProperties2001.Append(spaceBefore2001);
            paragraphProperties2001.Append(spaceAfter2001);
            paragraphProperties2001.Append(bulletColorText2001);
            paragraphProperties2001.Append(bulletSizeText2001);
            paragraphProperties2001.Append(bulletFontText2001);
            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count > 1 && !goalsArrayForSlide.ElementAt(1).isMainGoal)
                paragraphProperties2001.Append(pictureBullet2001);
            // ********************************* Custom Code Changes End *********************************
            paragraphProperties2001.Append(tabStopList2001);
            paragraphProperties2001.Append(defaultRunProperties2001);

            A.Run run2001 = new A.Run();

            // ********************************* Custom Code Changes Start *********************************
            int fontSize2002 = goalsArrayForSlide.Count > 1 && goalsArrayForSlide.ElementAt(1).isMainGoal ? 1300 : 1200;
            bool isBold2002 = goalsArrayForSlide.Count > 1 && goalsArrayForSlide.ElementAt(1).isMainGoal;
            string fontName2002 = goalsArrayForSlide.Count > 1 && goalsArrayForSlide.ElementAt(1).isMainGoal ? "Segoe UI Semibold" : "Segoe UI";
            // ********************************* Custom Code Changes End *********************************

            A.RunProperties runProperties2001 = new A.RunProperties() { Language = "en-US", FontSize = fontSize2002, Bold = isBold2002, Italic = false };
            A.EffectList effectList2001 = new A.EffectList();
            A.LatinFont latinFont2001 = new A.LatinFont() { Typeface = fontName2002 };
            A.ComplexScriptFont complexScriptFont2001 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            runProperties2001.Append(effectList2001);
            runProperties2001.Append(latinFont2001);
            runProperties2001.Append(complexScriptFont2001);
            A.Text text2001 = new A.Text();

            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count() > 1)
                text2001.Text = goalsArrayForSlide.ElementAt(1).goalName ?? "";
            else
                text2001.Text = "NO DATA";
            // ********************************* Custom Code Changes End *********************************

            run2001.Append(runProperties2001);
            run2001.Append(text2001);

            A.EndParagraphRunProperties endParagraphRunProperties2001 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false };
            A.LatinFont latinFont2002 = new A.LatinFont() { Typeface = "Segoe UI" };
            A.ComplexScriptFont complexScriptFont2002 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            endParagraphRunProperties2001.Append(latinFont2002);
            endParagraphRunProperties2001.Append(complexScriptFont2002);

            paragraph2001.Append(paragraphProperties2001);
            paragraph2001.Append(run2001);
            paragraph2001.Append(endParagraphRunProperties2001);

            textBody2001.Append(bodyProperties2001);
            textBody2001.Append(listStyle2001);
            textBody2001.Append(paragraph2001);

            A.TableCellProperties tableCellProperties2001 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties2001 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill2001 = new A.NoFill();

            leftBorderLineProperties2001.Append(noFill2001);

            A.RightBorderLineProperties rightBorderLineProperties2001 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill2002 = new A.NoFill();

            rightBorderLineProperties2001.Append(noFill2002);

            A.TopBorderLineProperties topBorderLineProperties2001 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill2003 = new A.NoFill();
            A.PresetDash presetDash2001 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round2001 = new A.Round();
            A.HeadEnd headEnd2001 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd2001 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties2001.Append(noFill2003);
            topBorderLineProperties2001.Append(presetDash2001);
            topBorderLineProperties2001.Append(round2001);
            topBorderLineProperties2001.Append(headEnd2001);
            topBorderLineProperties2001.Append(tailEnd2001);

            A.BottomBorderLineProperties bottomBorderLineProperties2001 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill2004 = new A.NoFill();
            A.PresetDash presetDash2002 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round2002 = new A.Round();
            A.HeadEnd headEnd2002 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd2002 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties2001.Append(noFill2004);
            bottomBorderLineProperties2001.Append(presetDash2002);
            bottomBorderLineProperties2001.Append(round2002);
            bottomBorderLineProperties2001.Append(headEnd2002);
            bottomBorderLineProperties2001.Append(tailEnd2002);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties2001 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill2005 = new A.NoFill();
            A.PresetDash presetDash2003 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties2001.Append(noFill2005);
            topLeftToBottomRightBorderLineProperties2001.Append(presetDash2003);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties2001 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill2006 = new A.NoFill();
            A.PresetDash presetDash2004 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties2001.Append(noFill2006);
            bottomLeftToTopRightBorderLineProperties2001.Append(presetDash2004);

            A.SolidFill solidFill2002 = new A.SolidFill();

            // ********************************* Custom Code Changes Start *********************************
            // Adding the background color according to the goal2 type
            string goalColor2002 = goalsArrayForSlide.Count > 1 && goalsArrayForSlide.ElementAt(1).isMainGoal ? "F5F5F5" : "FFFFFF";
            A.RgbColorModelHex rgbColorModelHex2002 = new A.RgbColorModelHex() { Val = goalColor2002 };
            // ********************************* Custom Code Changes End *********************************

            solidFill2002.Append(rgbColorModelHex2002);

            tableCellProperties2001.Append(leftBorderLineProperties2001);
            tableCellProperties2001.Append(rightBorderLineProperties2001);
            tableCellProperties2001.Append(topBorderLineProperties2001);
            tableCellProperties2001.Append(bottomBorderLineProperties2001);
            tableCellProperties2001.Append(topLeftToBottomRightBorderLineProperties2001);
            tableCellProperties2001.Append(bottomLeftToTopRightBorderLineProperties2001);
            tableCellProperties2001.Append(solidFill2002);

            tableCell7.Append(textBody2001);
            tableCell7.Append(tableCellProperties2001);



            A.TableCell tableCell8 = new A.TableCell();

            A.TextBody textBody13 = new A.TextBody();
            A.BodyProperties bodyProperties13 = new A.BodyProperties();
            A.ListStyle listStyle13 = new A.ListStyle();

            A.Paragraph paragraph13 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties8 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing8 = new A.LineSpacing();
            A.SpacingPercent spacingPercent8 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing8.Append(spacingPercent8);

            A.SpaceBefore spaceBefore6 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints11 = new A.SpacingPoints() { Val = 0 };

            spaceBefore6.Append(spacingPoints11);

            A.SpaceAfter spaceAfter6 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints12 = new A.SpacingPoints() { Val = 0 };

            spaceAfter6.Append(spacingPoints12);
            A.BulletColorText bulletColorText6 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText6 = new A.BulletSizeText();
            A.BulletFontText bulletFontText6 = new A.BulletFontText();
            A.NoBullet noBullet5 = new A.NoBullet();
            A.TabStopList tabStopList6 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties16 = new A.DefaultRunProperties();

            paragraphProperties8.Append(lineSpacing8);
            paragraphProperties8.Append(spaceBefore6);
            paragraphProperties8.Append(spaceAfter6);
            paragraphProperties8.Append(bulletColorText6);
            paragraphProperties8.Append(bulletSizeText6);
            paragraphProperties8.Append(bulletFontText6);
            paragraphProperties8.Append(noBullet5);
            paragraphProperties8.Append(tabStopList6);
            paragraphProperties8.Append(defaultRunProperties16);

            A.EndParagraphRunProperties endParagraphRunProperties10 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline8 = new A.Outline();
            A.NoFill noFill50 = new A.NoFill();

            outline8.Append(noFill50);

            A.SolidFill solidFill30 = new A.SolidFill();
            A.PresetColor presetColor4 = new A.PresetColor() { Val = A.PresetColorValues.Black };

            solidFill30.Append(presetColor4);
            A.EffectList effectList10 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText5 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText5 = new A.UnderlineFillText();
            A.LatinFont latinFont24 = new A.LatinFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont20 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont24 = new A.ComplexScriptFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties10.Append(outline8);
            endParagraphRunProperties10.Append(solidFill30);
            endParagraphRunProperties10.Append(effectList10);
            endParagraphRunProperties10.Append(underlineFollowsText5);
            endParagraphRunProperties10.Append(underlineFillText5);
            endParagraphRunProperties10.Append(latinFont24);
            endParagraphRunProperties10.Append(eastAsianFont20);
            endParagraphRunProperties10.Append(complexScriptFont24);

            paragraph13.Append(paragraphProperties8);
            paragraph13.Append(endParagraphRunProperties10);

            textBody13.Append(bodyProperties13);
            textBody13.Append(listStyle13);
            textBody13.Append(paragraph13);

            A.TableCellProperties tableCellProperties8 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties8 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill51 = new A.NoFill();

            leftBorderLineProperties8.Append(noFill51);

            A.RightBorderLineProperties rightBorderLineProperties8 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill52 = new A.NoFill();

            rightBorderLineProperties8.Append(noFill52);

            A.TopBorderLineProperties topBorderLineProperties8 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill53 = new A.NoFill();
            A.PresetDash presetDash26 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round12 = new A.Round();
            A.HeadEnd headEnd12 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd12 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties8.Append(noFill53);
            topBorderLineProperties8.Append(presetDash26);
            topBorderLineProperties8.Append(round12);
            topBorderLineProperties8.Append(headEnd12);
            topBorderLineProperties8.Append(tailEnd12);

            A.BottomBorderLineProperties bottomBorderLineProperties8 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill54 = new A.NoFill();
            A.PresetDash presetDash27 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round13 = new A.Round();
            A.HeadEnd headEnd13 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd13 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties8.Append(noFill54);
            bottomBorderLineProperties8.Append(presetDash27);
            bottomBorderLineProperties8.Append(round13);
            bottomBorderLineProperties8.Append(headEnd13);
            bottomBorderLineProperties8.Append(tailEnd13);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties8 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill55 = new A.NoFill();
            A.PresetDash presetDash28 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties8.Append(noFill55);
            topLeftToBottomRightBorderLineProperties8.Append(presetDash28);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties8 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill56 = new A.NoFill();
            A.PresetDash presetDash29 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties8.Append(noFill56);
            bottomLeftToTopRightBorderLineProperties8.Append(presetDash29);

            A.SolidFill solidFill2101 = new A.SolidFill();
            // ********************************* Custom Code Changes Start *********************************
            string colorCode2101 = goalsArrayForSlide.Count > 1 && goalsArrayForSlide.ElementAt(1).isMainGoal ? "F5F5F5" : "FFFFFF";
            // ********************************* Custom Code Changes End *********************************

            A.RgbColorModelHex rgbColorModelHex2101 = new A.RgbColorModelHex() { Val = colorCode2101 };

            solidFill2101.Append(rgbColorModelHex2101);

            tableCellProperties8.Append(leftBorderLineProperties8);
            tableCellProperties8.Append(rightBorderLineProperties8);
            tableCellProperties8.Append(topBorderLineProperties8);
            tableCellProperties8.Append(bottomBorderLineProperties8);
            tableCellProperties8.Append(topLeftToBottomRightBorderLineProperties8);
            tableCellProperties8.Append(bottomLeftToTopRightBorderLineProperties8);
            tableCellProperties8.Append(solidFill2101);

            tableCell8.Append(textBody13);
            tableCell8.Append(tableCellProperties8);

            A.TableCell tableCell9 = new A.TableCell();

            A.TextBody textBody14 = new A.TextBody();
            A.BodyProperties bodyProperties14 = new A.BodyProperties();
            A.ListStyle listStyle14 = new A.ListStyle();

            A.Paragraph paragraph14 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties9 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing9 = new A.LineSpacing();
            A.SpacingPercent spacingPercent9 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing9.Append(spacingPercent9);

            A.SpaceBefore spaceBefore7 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints13 = new A.SpacingPoints() { Val = 0 };

            spaceBefore7.Append(spacingPoints13);

            A.SpaceAfter spaceAfter7 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints14 = new A.SpacingPoints() { Val = 0 };

            spaceAfter7.Append(spacingPoints14);
            A.BulletColorText bulletColorText7 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText7 = new A.BulletSizeText();
            A.BulletFontText bulletFontText7 = new A.BulletFontText();
            A.NoBullet noBullet6 = new A.NoBullet();
            A.TabStopList tabStopList7 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties17 = new A.DefaultRunProperties();

            paragraphProperties9.Append(lineSpacing9);
            paragraphProperties9.Append(spaceBefore7);
            paragraphProperties9.Append(spaceAfter7);
            paragraphProperties9.Append(bulletColorText7);
            paragraphProperties9.Append(bulletSizeText7);
            paragraphProperties9.Append(bulletFontText7);
            paragraphProperties9.Append(noBullet6);
            paragraphProperties9.Append(tabStopList7);
            paragraphProperties9.Append(defaultRunProperties17);

            A.EndParagraphRunProperties endParagraphRunProperties11 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline9 = new A.Outline();
            A.NoFill noFill57 = new A.NoFill();

            outline9.Append(noFill57);

            A.SolidFill solidFill32 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "EAA300" };

            solidFill32.Append(rgbColorModelHex7);
            A.EffectList effectList11 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText6 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText6 = new A.UnderlineFillText();
            A.LatinFont latinFont25 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont21 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont25 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties11.Append(outline9);
            endParagraphRunProperties11.Append(solidFill32);
            endParagraphRunProperties11.Append(effectList11);
            endParagraphRunProperties11.Append(underlineFollowsText6);
            endParagraphRunProperties11.Append(underlineFillText6);
            endParagraphRunProperties11.Append(latinFont25);
            endParagraphRunProperties11.Append(eastAsianFont21);
            endParagraphRunProperties11.Append(complexScriptFont25);

            paragraph14.Append(paragraphProperties9);
            paragraph14.Append(endParagraphRunProperties11);

            textBody14.Append(bodyProperties14);
            textBody14.Append(listStyle14);
            textBody14.Append(paragraph14);

            A.TableCellProperties tableCellProperties9 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties9 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill58 = new A.NoFill();

            leftBorderLineProperties9.Append(noFill58);

            A.RightBorderLineProperties rightBorderLineProperties9 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill59 = new A.NoFill();

            rightBorderLineProperties9.Append(noFill59);

            A.TopBorderLineProperties topBorderLineProperties9 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill60 = new A.NoFill();
            A.PresetDash presetDash30 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round14 = new A.Round();
            A.HeadEnd headEnd14 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd14 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties9.Append(noFill60);
            topBorderLineProperties9.Append(presetDash30);
            topBorderLineProperties9.Append(round14);
            topBorderLineProperties9.Append(headEnd14);
            topBorderLineProperties9.Append(tailEnd14);

            A.BottomBorderLineProperties bottomBorderLineProperties9 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill61 = new A.NoFill();
            A.PresetDash presetDash31 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round15 = new A.Round();
            A.HeadEnd headEnd15 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd15 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties9.Append(noFill61);
            bottomBorderLineProperties9.Append(presetDash31);
            bottomBorderLineProperties9.Append(round15);
            bottomBorderLineProperties9.Append(headEnd15);
            bottomBorderLineProperties9.Append(tailEnd15);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties9 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill62 = new A.NoFill();
            A.PresetDash presetDash32 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties9.Append(noFill62);
            topLeftToBottomRightBorderLineProperties9.Append(presetDash32);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties9 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill63 = new A.NoFill();
            A.PresetDash presetDash33 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties9.Append(noFill63);
            bottomLeftToTopRightBorderLineProperties9.Append(presetDash33);

            A.SolidFill solidFill2201 = new A.SolidFill();
            // ********************************* Custom Code Changes Start *********************************
            string colorCode2201 = goalsArrayForSlide.Count > 1 && goalsArrayForSlide.ElementAt(1).isMainGoal ? "F5F5F5" : "FFFFFF";
            // ********************************* Custom Code Changes End *********************************

            A.RgbColorModelHex rgbColorModelHex2201 = new A.RgbColorModelHex() { Val = colorCode2201 };

            solidFill2201.Append(rgbColorModelHex2201);

            tableCellProperties9.Append(leftBorderLineProperties9);
            tableCellProperties9.Append(rightBorderLineProperties9);
            tableCellProperties9.Append(topBorderLineProperties9);
            tableCellProperties9.Append(bottomBorderLineProperties9);
            tableCellProperties9.Append(topLeftToBottomRightBorderLineProperties9);
            tableCellProperties9.Append(bottomLeftToTopRightBorderLineProperties9);
            tableCellProperties9.Append(solidFill2201);

            tableCell9.Append(textBody14);
            tableCell9.Append(tableCellProperties9);

            A.ExtensionList extensionList6 = new A.ExtensionList();

            A.Extension extension6 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement15 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"3520631283\" />");

            // extension6.Append(openXmlUnknownElement15);

            extensionList6.Append(extension6);

            tableRow3.Append(tableCell7);
            tableRow3.Append(tableCell8);
            tableRow3.Append(tableCell9);
            tableRow3.Append(extensionList6);

            A.TableRow tableRow4 = new A.TableRow() { Height = 766152L };

            A.TableCell tableCell10 = new A.TableCell();

            A.TextBody textBody3001 = new A.TextBody();
            A.BodyProperties bodyProperties3001 = new A.BodyProperties();
            A.ListStyle listStyle3001 = new A.ListStyle();

            A.Paragraph paragraph3001 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties3001 = new A.ParagraphProperties() { LeftMargin = 171450, RightMargin = 0, Indent = -171450, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing3001 = new A.LineSpacing();
            A.SpacingPercent spacingPercent3001 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing3001.Append(spacingPercent3001);

            A.SpaceBefore spaceBefore3001 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints3001 = new A.SpacingPoints() { Val = 0 };

            spaceBefore3001.Append(spacingPoints3001);

            A.SpaceAfter spaceAfter3001 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints3002 = new A.SpacingPoints() { Val = 0 };

            spaceAfter3001.Append(spacingPoints3002);
            A.BulletColorText bulletColorText3001 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText3001 = new A.BulletSizeText();
            A.BulletFontText bulletFontText3001 = new A.BulletFontText();

            A.PictureBullet pictureBullet3001 = new A.PictureBullet();

            A.Blip blip3001 = new A.Blip() { Embed = "rId4" };

            A.BlipExtensionList blipExtensionList3001 = new A.BlipExtensionList();

            A.BlipExtension blipExtension3001 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            blipExtensionList3001.Append(blipExtension3001);

            blip3001.Append(blipExtensionList3001);

            pictureBullet3001.Append(blip3001);
            A.TabStopList tabStopList3001 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties3001 = new A.DefaultRunProperties();

            paragraphProperties3001.Append(lineSpacing3001);
            paragraphProperties3001.Append(spaceBefore3001);
            paragraphProperties3001.Append(spaceAfter3001);
            paragraphProperties3001.Append(bulletColorText3001);
            paragraphProperties3001.Append(bulletSizeText3001);
            paragraphProperties3001.Append(bulletFontText3001);
            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count > 2 && !goalsArrayForSlide.ElementAt(2).isMainGoal)
                paragraphProperties3001.Append(pictureBullet3001);
            // ********************************* Custom Code Changes End *********************************
            paragraphProperties3001.Append(tabStopList3001);
            paragraphProperties3001.Append(defaultRunProperties3001);

            A.Run run3001 = new A.Run();

            // ********************************* Custom Code Changes Start *********************************
            int fontSize3001 = goalsArrayForSlide.Count > 2 && goalsArrayForSlide.ElementAt(2).isMainGoal ? 1300 : 1200;
            bool isBold3001 = goalsArrayForSlide.Count > 2 && goalsArrayForSlide.ElementAt(2).isMainGoal;
            string fontName3001 = goalsArrayForSlide.Count > 2 && goalsArrayForSlide.ElementAt(2).isMainGoal ? "Segoe UI Semibold" : "Segoe UI";
            // ********************************* Custom Code Changes End *********************************

            A.RunProperties runProperties3001 = new A.RunProperties() { Language = "en-US", FontSize = fontSize3001, Bold = isBold3001, Italic = false };
            A.EffectList effectList3001 = new A.EffectList();
            A.LatinFont latinFont3001 = new A.LatinFont() { Typeface = fontName3001 };
            A.ComplexScriptFont complexScriptFont3001 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            runProperties3001.Append(effectList3001);
            runProperties3001.Append(latinFont3001);
            runProperties3001.Append(complexScriptFont3001);
            A.Text text3001 = new A.Text();

            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count() > 2)
                text3001.Text = goalsArrayForSlide.ElementAt(2).goalName ?? "";
            else
                text3001.Text = "NO DATA";
            // ********************************* Custom Code Changes End *********************************

            run3001.Append(runProperties3001);
            run3001.Append(text3001);

            A.EndParagraphRunProperties endParagraphRunProperties3001 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false };
            A.LatinFont latinFont3002 = new A.LatinFont() { Typeface = "Segoe UI" };
            A.ComplexScriptFont complexScriptFont3002 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            endParagraphRunProperties3001.Append(latinFont3002);
            endParagraphRunProperties3001.Append(complexScriptFont3002);

            paragraph3001.Append(paragraphProperties3001);
            paragraph3001.Append(run3001);
            paragraph3001.Append(endParagraphRunProperties3001);

            textBody3001.Append(bodyProperties3001);
            textBody3001.Append(listStyle3001);
            textBody3001.Append(paragraph3001);

            A.TableCellProperties tableCellProperties3001 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties3001 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill3001 = new A.NoFill();

            leftBorderLineProperties3001.Append(noFill3001);

            A.RightBorderLineProperties rightBorderLineProperties3001 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill3002 = new A.NoFill();

            rightBorderLineProperties3001.Append(noFill3002);

            A.TopBorderLineProperties topBorderLineProperties3001 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill3003 = new A.NoFill();
            A.PresetDash presetDash3001 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round3001 = new A.Round();
            A.HeadEnd headEnd3001 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd3001 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties3001.Append(noFill3003);
            topBorderLineProperties3001.Append(presetDash3001);
            topBorderLineProperties3001.Append(round3001);
            topBorderLineProperties3001.Append(headEnd3001);
            topBorderLineProperties3001.Append(tailEnd3001);

            A.BottomBorderLineProperties bottomBorderLineProperties3001 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill3004 = new A.NoFill();
            A.PresetDash presetDash3002 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round3002 = new A.Round();
            A.HeadEnd headEnd3002 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd3002 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties3001.Append(noFill3004);
            bottomBorderLineProperties3001.Append(presetDash3002);
            bottomBorderLineProperties3001.Append(round3002);
            bottomBorderLineProperties3001.Append(headEnd3002);
            bottomBorderLineProperties3001.Append(tailEnd3002);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties3001 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill3005 = new A.NoFill();
            A.PresetDash presetDash3003 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties3001.Append(noFill3005);
            topLeftToBottomRightBorderLineProperties3001.Append(presetDash3003);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties3001 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill3006 = new A.NoFill();
            A.PresetDash presetDash3004 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties3001.Append(noFill3006);
            bottomLeftToTopRightBorderLineProperties3001.Append(presetDash3004);

            A.SolidFill solidFill3001 = new A.SolidFill();

            // ********************************* Custom Code Changes Start *********************************
            // Adding the background color according to the goal3 type
            string goalColor3001 = goalsArrayForSlide.Count > 2 && goalsArrayForSlide.ElementAt(2).isMainGoal ? "F5F5F5" : "FFFFFF";
            A.RgbColorModelHex rgbColorModelHex3001 = new A.RgbColorModelHex() { Val = goalColor3001 };
            // ********************************* Custom Code Changes End *********************************

            solidFill3001.Append(rgbColorModelHex3001);

            tableCellProperties3001.Append(leftBorderLineProperties3001);
            tableCellProperties3001.Append(rightBorderLineProperties3001);
            tableCellProperties3001.Append(topBorderLineProperties3001);
            tableCellProperties3001.Append(bottomBorderLineProperties3001);
            tableCellProperties3001.Append(topLeftToBottomRightBorderLineProperties3001);
            tableCellProperties3001.Append(bottomLeftToTopRightBorderLineProperties3001);
            tableCellProperties3001.Append(solidFill3001);

            tableCell10.Append(textBody3001);
            tableCell10.Append(tableCellProperties3001);


            A.TableCell tableCell11 = new A.TableCell();

            A.TextBody textBody16 = new A.TextBody();
            A.BodyProperties bodyProperties16 = new A.BodyProperties();
            A.ListStyle listStyle16 = new A.ListStyle();

            A.Paragraph paragraph16 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties11 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing11 = new A.LineSpacing();
            A.SpacingPercent spacingPercent11 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing11.Append(spacingPercent11);

            A.SpaceBefore spaceBefore9 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints17 = new A.SpacingPoints() { Val = 0 };

            spaceBefore9.Append(spacingPoints17);

            A.SpaceAfter spaceAfter9 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints18 = new A.SpacingPoints() { Val = 0 };

            spaceAfter9.Append(spacingPoints18);
            A.BulletColorText bulletColorText9 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText9 = new A.BulletSizeText();
            A.BulletFontText bulletFontText9 = new A.BulletFontText();
            A.NoBullet noBullet7 = new A.NoBullet();
            A.TabStopList tabStopList8 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties18 = new A.DefaultRunProperties();

            paragraphProperties11.Append(lineSpacing11);
            paragraphProperties11.Append(spaceBefore9);
            paragraphProperties11.Append(spaceAfter9);
            paragraphProperties11.Append(bulletColorText9);
            paragraphProperties11.Append(bulletSizeText9);
            paragraphProperties11.Append(bulletFontText9);
            paragraphProperties11.Append(noBullet7);
            paragraphProperties11.Append(tabStopList8);
            paragraphProperties11.Append(defaultRunProperties18);

            A.EndParagraphRunProperties endParagraphRunProperties12 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline10 = new A.Outline();
            A.NoFill noFill70 = new A.NoFill();

            outline10.Append(noFill70);

            A.SolidFill solidFill35 = new A.SolidFill();
            A.PresetColor presetColor5 = new A.PresetColor() { Val = A.PresetColorValues.Black };

            solidFill35.Append(presetColor5);
            A.EffectList effectList13 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText7 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText7 = new A.UnderlineFillText();
            A.LatinFont latinFont27 = new A.LatinFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont22 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont27 = new A.ComplexScriptFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties12.Append(outline10);
            endParagraphRunProperties12.Append(solidFill35);
            endParagraphRunProperties12.Append(effectList13);
            endParagraphRunProperties12.Append(underlineFollowsText7);
            endParagraphRunProperties12.Append(underlineFillText7);
            endParagraphRunProperties12.Append(latinFont27);
            endParagraphRunProperties12.Append(eastAsianFont22);
            endParagraphRunProperties12.Append(complexScriptFont27);

            paragraph16.Append(paragraphProperties11);
            paragraph16.Append(endParagraphRunProperties12);

            textBody16.Append(bodyProperties16);
            textBody16.Append(listStyle16);
            textBody16.Append(paragraph16);

            A.TableCellProperties tableCellProperties11 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties11 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill71 = new A.NoFill();

            leftBorderLineProperties11.Append(noFill71);

            A.RightBorderLineProperties rightBorderLineProperties11 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill72 = new A.NoFill();

            rightBorderLineProperties11.Append(noFill72);

            A.TopBorderLineProperties topBorderLineProperties11 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill73 = new A.NoFill();
            A.PresetDash presetDash38 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round18 = new A.Round();
            A.HeadEnd headEnd18 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd18 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties11.Append(noFill73);
            topBorderLineProperties11.Append(presetDash38);
            topBorderLineProperties11.Append(round18);
            topBorderLineProperties11.Append(headEnd18);
            topBorderLineProperties11.Append(tailEnd18);

            A.BottomBorderLineProperties bottomBorderLineProperties11 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill74 = new A.NoFill();
            A.PresetDash presetDash39 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round19 = new A.Round();
            A.HeadEnd headEnd19 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd19 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties11.Append(noFill74);
            bottomBorderLineProperties11.Append(presetDash39);
            bottomBorderLineProperties11.Append(round19);
            bottomBorderLineProperties11.Append(headEnd19);
            bottomBorderLineProperties11.Append(tailEnd19);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties11 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill75 = new A.NoFill();
            A.PresetDash presetDash40 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties11.Append(noFill75);
            topLeftToBottomRightBorderLineProperties11.Append(presetDash40);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties11 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill76 = new A.NoFill();
            A.PresetDash presetDash41 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties11.Append(noFill76);
            bottomLeftToTopRightBorderLineProperties11.Append(presetDash41);

            A.SolidFill solidFill3101 = new A.SolidFill();
            // ********************************* Custom Code Changes Start *********************************
            string colorCode3101 = goalsArrayForSlide.Count > 2 && goalsArrayForSlide.ElementAt(2).isMainGoal ? "F5F5F5" : "FFFFFF";
            // ********************************* Custom Code Changes End *********************************

            A.RgbColorModelHex rgbColorModelHex3101 = new A.RgbColorModelHex() { Val = colorCode3101 };

            solidFill3101.Append(rgbColorModelHex3101);

            tableCellProperties11.Append(leftBorderLineProperties11);
            tableCellProperties11.Append(rightBorderLineProperties11);
            tableCellProperties11.Append(topBorderLineProperties11);
            tableCellProperties11.Append(bottomBorderLineProperties11);
            tableCellProperties11.Append(topLeftToBottomRightBorderLineProperties11);
            tableCellProperties11.Append(bottomLeftToTopRightBorderLineProperties11);
            tableCellProperties11.Append(solidFill3101);

            tableCell11.Append(textBody16);
            tableCell11.Append(tableCellProperties11);

            A.TableCell tableCell12 = new A.TableCell();

            A.TextBody textBody17 = new A.TextBody();
            A.BodyProperties bodyProperties17 = new A.BodyProperties();
            A.ListStyle listStyle17 = new A.ListStyle();

            A.Paragraph paragraph17 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties12 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing12 = new A.LineSpacing();
            A.SpacingPercent spacingPercent12 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing12.Append(spacingPercent12);

            A.SpaceBefore spaceBefore10 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints19 = new A.SpacingPoints() { Val = 0 };

            spaceBefore10.Append(spacingPoints19);

            A.SpaceAfter spaceAfter10 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints20 = new A.SpacingPoints() { Val = 0 };

            spaceAfter10.Append(spacingPoints20);
            A.BulletColorText bulletColorText10 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText10 = new A.BulletSizeText();
            A.BulletFontText bulletFontText10 = new A.BulletFontText();
            A.NoBullet noBullet8 = new A.NoBullet();
            A.TabStopList tabStopList9 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties19 = new A.DefaultRunProperties();

            paragraphProperties12.Append(lineSpacing12);
            paragraphProperties12.Append(spaceBefore10);
            paragraphProperties12.Append(spaceAfter10);
            paragraphProperties12.Append(bulletColorText10);
            paragraphProperties12.Append(bulletSizeText10);
            paragraphProperties12.Append(bulletFontText10);
            paragraphProperties12.Append(noBullet8);
            paragraphProperties12.Append(tabStopList9);
            paragraphProperties12.Append(defaultRunProperties19);

            A.EndParagraphRunProperties endParagraphRunProperties13 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline11 = new A.Outline();
            A.NoFill noFill77 = new A.NoFill();

            outline11.Append(noFill77);

            A.SolidFill solidFill37 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "5EC65A" };

            solidFill37.Append(rgbColorModelHex8);
            A.EffectList effectList14 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText8 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText8 = new A.UnderlineFillText();
            A.LatinFont latinFont28 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont23 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont28 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties13.Append(outline11);
            endParagraphRunProperties13.Append(solidFill37);
            endParagraphRunProperties13.Append(effectList14);
            endParagraphRunProperties13.Append(underlineFollowsText8);
            endParagraphRunProperties13.Append(underlineFillText8);
            endParagraphRunProperties13.Append(latinFont28);
            endParagraphRunProperties13.Append(eastAsianFont23);
            endParagraphRunProperties13.Append(complexScriptFont28);

            paragraph17.Append(paragraphProperties12);
            paragraph17.Append(endParagraphRunProperties13);

            textBody17.Append(bodyProperties17);
            textBody17.Append(listStyle17);
            textBody17.Append(paragraph17);

            A.TableCellProperties tableCellProperties12 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties12 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill78 = new A.NoFill();

            leftBorderLineProperties12.Append(noFill78);

            A.RightBorderLineProperties rightBorderLineProperties12 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill79 = new A.NoFill();

            rightBorderLineProperties12.Append(noFill79);

            A.TopBorderLineProperties topBorderLineProperties12 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill80 = new A.NoFill();
            A.PresetDash presetDash42 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round20 = new A.Round();
            A.HeadEnd headEnd20 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd20 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties12.Append(noFill80);
            topBorderLineProperties12.Append(presetDash42);
            topBorderLineProperties12.Append(round20);
            topBorderLineProperties12.Append(headEnd20);
            topBorderLineProperties12.Append(tailEnd20);

            A.BottomBorderLineProperties bottomBorderLineProperties12 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill81 = new A.NoFill();
            A.PresetDash presetDash43 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round21 = new A.Round();
            A.HeadEnd headEnd21 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd21 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties12.Append(noFill81);
            bottomBorderLineProperties12.Append(presetDash43);
            bottomBorderLineProperties12.Append(round21);
            bottomBorderLineProperties12.Append(headEnd21);
            bottomBorderLineProperties12.Append(tailEnd21);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties12 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill82 = new A.NoFill();
            A.PresetDash presetDash44 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties12.Append(noFill82);
            topLeftToBottomRightBorderLineProperties12.Append(presetDash44);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties12 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill83 = new A.NoFill();
            A.PresetDash presetDash45 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties12.Append(noFill83);
            bottomLeftToTopRightBorderLineProperties12.Append(presetDash45);

            A.SolidFill solidFill3201 = new A.SolidFill();
            // ********************************* Custom Code Changes Start *********************************
            string colorCode3201 = goalsArrayForSlide.Count > 2 && goalsArrayForSlide.ElementAt(2).isMainGoal ? "F5F5F5" : "FFFFFF";
            // ********************************* Custom Code Changes End *********************************

            A.RgbColorModelHex rgbColorModelHex3201 = new A.RgbColorModelHex() { Val = colorCode3201 };

            solidFill3201.Append(rgbColorModelHex3201);

            tableCellProperties12.Append(leftBorderLineProperties12);
            tableCellProperties12.Append(rightBorderLineProperties12);
            tableCellProperties12.Append(topBorderLineProperties12);
            tableCellProperties12.Append(bottomBorderLineProperties12);
            tableCellProperties12.Append(topLeftToBottomRightBorderLineProperties12);
            tableCellProperties12.Append(bottomLeftToTopRightBorderLineProperties12);
            tableCellProperties12.Append(solidFill3201);

            tableCell12.Append(textBody17);
            tableCell12.Append(tableCellProperties12);

            A.ExtensionList extensionList7 = new A.ExtensionList();

            A.Extension extension7 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement17 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"3020163789\" />");

            // extension7.Append(openXmlUnknownElement17);

            extensionList7.Append(extension7);

            tableRow4.Append(tableCell10);
            tableRow4.Append(tableCell11);
            tableRow4.Append(tableCell12);
            tableRow4.Append(extensionList7);

            A.TableRow tableRow5 = new A.TableRow() { Height = 747144L };

            A.TableCell tableCell13 = new A.TableCell();

            A.TextBody textBody4001 = new A.TextBody();
            A.BodyProperties bodyProperties4001 = new A.BodyProperties();
            A.ListStyle listStyle4001 = new A.ListStyle();

            A.Paragraph paragraph4001 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties4001 = new A.ParagraphProperties() { LeftMargin = 171450, RightMargin = 0, Indent = -171450, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing4001 = new A.LineSpacing();
            A.SpacingPercent spacingPercent4001 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing4001.Append(spacingPercent4001);

            A.SpaceBefore spaceBefore4001 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints4001 = new A.SpacingPoints() { Val = 0 };

            spaceBefore4001.Append(spacingPoints4001);

            A.SpaceAfter spaceAfter4001 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints4002 = new A.SpacingPoints() { Val = 0 };

            spaceAfter4001.Append(spacingPoints4002);
            A.BulletColorText bulletColorText4001 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText4001 = new A.BulletSizeText();
            A.BulletFontText bulletFontText4001 = new A.BulletFontText();

            A.PictureBullet pictureBullet4001 = new A.PictureBullet();

            A.Blip blip4001 = new A.Blip() { Embed = "rId4" };

            A.BlipExtensionList blipExtensionList4001 = new A.BlipExtensionList();

            A.BlipExtension blipExtension4001 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            blipExtensionList4001.Append(blipExtension4001);

            blip4001.Append(blipExtensionList4001);

            pictureBullet4001.Append(blip4001);
            A.TabStopList tabStopList4001 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties4001 = new A.DefaultRunProperties();

            paragraphProperties4001.Append(lineSpacing4001);
            paragraphProperties4001.Append(spaceBefore4001);
            paragraphProperties4001.Append(spaceAfter4001);
            paragraphProperties4001.Append(bulletColorText4001);
            paragraphProperties4001.Append(bulletSizeText4001);
            paragraphProperties4001.Append(bulletFontText4001);
            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count > 3 && !goalsArrayForSlide.ElementAt(3).isMainGoal)
                paragraphProperties4001.Append(pictureBullet4001);
            // ********************************* Custom Code Changes End *********************************
            paragraphProperties4001.Append(tabStopList4001);
            paragraphProperties4001.Append(defaultRunProperties4001);

            A.Run run4001 = new A.Run();

            // ********************************* Custom Code Changes Start *********************************
            int fontSize4001 = goalsArrayForSlide.Count > 3 && goalsArrayForSlide.ElementAt(3).isMainGoal ? 1300 : 1200;
            bool isBold4001 = goalsArrayForSlide.Count > 3 && goalsArrayForSlide.ElementAt(3).isMainGoal;
            string fontName4001 = goalsArrayForSlide.Count > 3 && goalsArrayForSlide.ElementAt(3).isMainGoal ? "Segoe UI Semibold" : "Segoe UI";
            // ********************************* Custom Code Changes End *********************************

            A.RunProperties runProperties4001 = new A.RunProperties() { Language = "en-US", FontSize = fontSize4001, Bold = isBold4001, Italic = false };
            A.EffectList effectList4001 = new A.EffectList();
            A.LatinFont latinFont4001 = new A.LatinFont() { Typeface = fontName4001 };
            A.ComplexScriptFont complexScriptFont4001 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            runProperties4001.Append(effectList4001);
            runProperties4001.Append(latinFont4001);
            runProperties4001.Append(complexScriptFont4001);
            A.Text text4001 = new A.Text();

            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count() > 3)
                text4001.Text = goalsArrayForSlide.ElementAt(3).goalName ?? "";
            else
                text4001.Text = "NO DATA";
            // ********************************* Custom Code Changes End *********************************

            run4001.Append(runProperties4001);
            run4001.Append(text4001);

            A.EndParagraphRunProperties endParagraphRunProperties4001 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false };
            A.LatinFont latinFont4002 = new A.LatinFont() { Typeface = "Segoe UI" };
            A.ComplexScriptFont complexScriptFont4002 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            endParagraphRunProperties4001.Append(latinFont4002);
            endParagraphRunProperties4001.Append(complexScriptFont4002);

            paragraph4001.Append(paragraphProperties4001);
            paragraph4001.Append(run4001);
            paragraph4001.Append(endParagraphRunProperties4001);

            textBody4001.Append(bodyProperties4001);
            textBody4001.Append(listStyle4001);
            textBody4001.Append(paragraph4001);

            A.TableCellProperties tableCellProperties4001 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties4001 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill4001 = new A.NoFill();

            leftBorderLineProperties4001.Append(noFill4001);

            A.RightBorderLineProperties rightBorderLineProperties4001 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill4002 = new A.NoFill();

            rightBorderLineProperties4001.Append(noFill4002);

            A.TopBorderLineProperties topBorderLineProperties4001 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill4003 = new A.NoFill();
            A.PresetDash presetDash4001 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round4001 = new A.Round();
            A.HeadEnd headEnd4001 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd4001 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties4001.Append(noFill4003);
            topBorderLineProperties4001.Append(presetDash4001);
            topBorderLineProperties4001.Append(round4001);
            topBorderLineProperties4001.Append(headEnd4001);
            topBorderLineProperties4001.Append(tailEnd4001);

            A.BottomBorderLineProperties bottomBorderLineProperties4001 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill4004 = new A.NoFill();
            A.PresetDash presetDash4002 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round4002 = new A.Round();
            A.HeadEnd headEnd4002 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd4002 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties4001.Append(noFill4004);
            bottomBorderLineProperties4001.Append(presetDash4002);
            bottomBorderLineProperties4001.Append(round4002);
            bottomBorderLineProperties4001.Append(headEnd4002);
            bottomBorderLineProperties4001.Append(tailEnd4002);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties4001 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill4005 = new A.NoFill();
            A.PresetDash presetDash4003 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties4001.Append(noFill4005);
            topLeftToBottomRightBorderLineProperties4001.Append(presetDash4003);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties4001 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill4006 = new A.NoFill();
            A.PresetDash presetDash4004 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties4001.Append(noFill4006);
            bottomLeftToTopRightBorderLineProperties4001.Append(presetDash4004);

            A.SolidFill solidFill4001 = new A.SolidFill();

            // ********************************* Custom Code Changes Start *********************************
            // Adding the background color according to the goal4 type
            string goalColor4001 = goalsArrayForSlide.Count > 3 && goalsArrayForSlide.ElementAt(3).isMainGoal ? "F5F5F5" : "FFFFFF";
            A.RgbColorModelHex rgbColorModelHex4001 = new A.RgbColorModelHex() { Val = goalColor4001 };
            // ********************************* Custom Code Changes End *********************************

            solidFill4001.Append(rgbColorModelHex4001);

            tableCellProperties4001.Append(leftBorderLineProperties4001);
            tableCellProperties4001.Append(rightBorderLineProperties4001);
            tableCellProperties4001.Append(topBorderLineProperties4001);
            tableCellProperties4001.Append(bottomBorderLineProperties4001);
            tableCellProperties4001.Append(topLeftToBottomRightBorderLineProperties4001);
            tableCellProperties4001.Append(bottomLeftToTopRightBorderLineProperties4001);
            tableCellProperties4001.Append(solidFill4001);

            tableCell13.Append(textBody4001);
            tableCell13.Append(tableCellProperties4001);


            A.TableCell tableCell14 = new A.TableCell();

            A.TextBody textBody19 = new A.TextBody();
            A.BodyProperties bodyProperties19 = new A.BodyProperties();
            A.ListStyle listStyle19 = new A.ListStyle();

            A.Paragraph paragraph19 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties14 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing13 = new A.LineSpacing();
            A.SpacingPercent spacingPercent13 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing13.Append(spacingPercent13);

            A.SpaceBefore spaceBefore11 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints21 = new A.SpacingPoints() { Val = 0 };

            spaceBefore11.Append(spacingPoints21);

            A.SpaceAfter spaceAfter11 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints22 = new A.SpacingPoints() { Val = 0 };

            spaceAfter11.Append(spacingPoints22);
            A.BulletColorText bulletColorText11 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText11 = new A.BulletSizeText();
            A.BulletFontText bulletFontText12 = new A.BulletFontText();
            A.NoBullet noBullet9 = new A.NoBullet();
            A.TabStopList tabStopList10 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties20 = new A.DefaultRunProperties();

            paragraphProperties14.Append(lineSpacing13);
            paragraphProperties14.Append(spaceBefore11);
            paragraphProperties14.Append(spaceAfter11);
            paragraphProperties14.Append(bulletColorText11);
            paragraphProperties14.Append(bulletSizeText11);
            paragraphProperties14.Append(bulletFontText12);
            paragraphProperties14.Append(noBullet9);
            paragraphProperties14.Append(tabStopList10);
            paragraphProperties14.Append(defaultRunProperties20);

            A.EndParagraphRunProperties endParagraphRunProperties14 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline12 = new A.Outline();
            A.NoFill noFill90 = new A.NoFill();

            outline12.Append(noFill90);

            A.SolidFill solidFill40 = new A.SolidFill();
            A.PresetColor presetColor6 = new A.PresetColor() { Val = A.PresetColorValues.Black };

            solidFill40.Append(presetColor6);
            A.EffectList effectList15 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText9 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText9 = new A.UnderlineFillText();
            A.LatinFont latinFont30 = new A.LatinFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont24 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont30 = new A.ComplexScriptFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties14.Append(outline12);
            endParagraphRunProperties14.Append(solidFill40);
            endParagraphRunProperties14.Append(effectList15);
            endParagraphRunProperties14.Append(underlineFollowsText9);
            endParagraphRunProperties14.Append(underlineFillText9);
            endParagraphRunProperties14.Append(latinFont30);
            endParagraphRunProperties14.Append(eastAsianFont24);
            endParagraphRunProperties14.Append(complexScriptFont30);

            paragraph19.Append(paragraphProperties14);
            paragraph19.Append(endParagraphRunProperties14);

            textBody19.Append(bodyProperties19);
            textBody19.Append(listStyle19);
            textBody19.Append(paragraph19);

            A.TableCellProperties tableCellProperties14 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties14 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill91 = new A.NoFill();

            leftBorderLineProperties14.Append(noFill91);

            A.RightBorderLineProperties rightBorderLineProperties14 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill92 = new A.NoFill();

            rightBorderLineProperties14.Append(noFill92);

            A.TopBorderLineProperties topBorderLineProperties14 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill93 = new A.NoFill();
            A.PresetDash presetDash50 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round24 = new A.Round();
            A.HeadEnd headEnd24 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd24 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties14.Append(noFill93);
            topBorderLineProperties14.Append(presetDash50);
            topBorderLineProperties14.Append(round24);
            topBorderLineProperties14.Append(headEnd24);
            topBorderLineProperties14.Append(tailEnd24);

            A.BottomBorderLineProperties bottomBorderLineProperties14 = new A.BottomBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill94 = new A.NoFill();
            A.PresetDash presetDash51 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round25 = new A.Round();
            A.HeadEnd headEnd25 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd25 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties14.Append(noFill94);
            bottomBorderLineProperties14.Append(presetDash51);
            bottomBorderLineProperties14.Append(round25);
            bottomBorderLineProperties14.Append(headEnd25);
            bottomBorderLineProperties14.Append(tailEnd25);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties14 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill95 = new A.NoFill();
            A.PresetDash presetDash52 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties14.Append(noFill95);
            topLeftToBottomRightBorderLineProperties14.Append(presetDash52);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties14 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill96 = new A.NoFill();
            A.PresetDash presetDash53 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties14.Append(noFill96);
            bottomLeftToTopRightBorderLineProperties14.Append(presetDash53);

            A.SolidFill solidFill4101 = new A.SolidFill();
            // ********************************* Custom Code Changes Start *********************************
            string colorCode4101 = goalsArrayForSlide.Count > 3 && goalsArrayForSlide.ElementAt(3).isMainGoal ? "F5F5F5" : "FFFFFF";
            // ********************************* Custom Code Changes End *********************************

            A.RgbColorModelHex rgbColorModelHex4101 = new A.RgbColorModelHex() { Val = colorCode4101 };

            solidFill4101.Append(rgbColorModelHex4101);

            tableCellProperties14.Append(leftBorderLineProperties14);
            tableCellProperties14.Append(rightBorderLineProperties14);
            tableCellProperties14.Append(topBorderLineProperties14);
            tableCellProperties14.Append(bottomBorderLineProperties14);
            tableCellProperties14.Append(topLeftToBottomRightBorderLineProperties14);
            tableCellProperties14.Append(bottomLeftToTopRightBorderLineProperties14);
            tableCellProperties14.Append(solidFill4101);

            tableCell14.Append(textBody19);
            tableCell14.Append(tableCellProperties14);

            A.TableCell tableCell15 = new A.TableCell();

            A.TextBody textBody20 = new A.TextBody();
            A.BodyProperties bodyProperties20 = new A.BodyProperties();
            A.ListStyle listStyle20 = new A.ListStyle();

            A.Paragraph paragraph20 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties15 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing14 = new A.LineSpacing();
            A.SpacingPercent spacingPercent14 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing14.Append(spacingPercent14);

            A.SpaceBefore spaceBefore12 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints23 = new A.SpacingPoints() { Val = 0 };

            spaceBefore12.Append(spacingPoints23);

            A.SpaceAfter spaceAfter12 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints24 = new A.SpacingPoints() { Val = 0 };

            spaceAfter12.Append(spacingPoints24);
            A.BulletColorText bulletColorText12 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText12 = new A.BulletSizeText();
            A.BulletFontText bulletFontText13 = new A.BulletFontText();
            A.NoBullet noBullet10 = new A.NoBullet();
            A.TabStopList tabStopList11 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties21 = new A.DefaultRunProperties();

            paragraphProperties15.Append(lineSpacing14);
            paragraphProperties15.Append(spaceBefore12);
            paragraphProperties15.Append(spaceAfter12);
            paragraphProperties15.Append(bulletColorText12);
            paragraphProperties15.Append(bulletSizeText12);
            paragraphProperties15.Append(bulletFontText13);
            paragraphProperties15.Append(noBullet10);
            paragraphProperties15.Append(tabStopList11);
            paragraphProperties15.Append(defaultRunProperties21);

            A.EndParagraphRunProperties endParagraphRunProperties15 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline13 = new A.Outline();
            A.NoFill noFill97 = new A.NoFill();

            outline13.Append(noFill97);

            A.SolidFill solidFill42 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "5EC65A" };

            solidFill42.Append(rgbColorModelHex9);
            A.EffectList effectList16 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText10 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText10 = new A.UnderlineFillText();
            A.LatinFont latinFont31 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont25 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont31 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties15.Append(outline13);
            endParagraphRunProperties15.Append(solidFill42);
            endParagraphRunProperties15.Append(effectList16);
            endParagraphRunProperties15.Append(underlineFollowsText10);
            endParagraphRunProperties15.Append(underlineFillText10);
            endParagraphRunProperties15.Append(latinFont31);
            endParagraphRunProperties15.Append(eastAsianFont25);
            endParagraphRunProperties15.Append(complexScriptFont31);

            paragraph20.Append(paragraphProperties15);
            paragraph20.Append(endParagraphRunProperties15);

            textBody20.Append(bodyProperties20);
            textBody20.Append(listStyle20);
            textBody20.Append(paragraph20);

            A.TableCellProperties tableCellProperties15 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties15 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill98 = new A.NoFill();

            leftBorderLineProperties15.Append(noFill98);

            A.RightBorderLineProperties rightBorderLineProperties15 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill99 = new A.NoFill();

            rightBorderLineProperties15.Append(noFill99);

            A.TopBorderLineProperties topBorderLineProperties15 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill100 = new A.NoFill();
            A.PresetDash presetDash54 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round26 = new A.Round();
            A.HeadEnd headEnd26 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd26 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties15.Append(noFill100);
            topBorderLineProperties15.Append(presetDash54);
            topBorderLineProperties15.Append(round26);
            topBorderLineProperties15.Append(headEnd26);
            topBorderLineProperties15.Append(tailEnd26);

            A.BottomBorderLineProperties bottomBorderLineProperties15 = new A.BottomBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill101 = new A.NoFill();
            A.PresetDash presetDash55 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round27 = new A.Round();
            A.HeadEnd headEnd27 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd27 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties15.Append(noFill101);
            bottomBorderLineProperties15.Append(presetDash55);
            bottomBorderLineProperties15.Append(round27);
            bottomBorderLineProperties15.Append(headEnd27);
            bottomBorderLineProperties15.Append(tailEnd27);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties15 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill102 = new A.NoFill();
            A.PresetDash presetDash56 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties15.Append(noFill102);
            topLeftToBottomRightBorderLineProperties15.Append(presetDash56);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties15 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill103 = new A.NoFill();
            A.PresetDash presetDash57 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties15.Append(noFill103);
            bottomLeftToTopRightBorderLineProperties15.Append(presetDash57);

            A.SolidFill solidFill4201 = new A.SolidFill();
            // ********************************* Custom Code Changes Start *********************************
            string colorCode4201 = goalsArrayForSlide.Count > 3 && goalsArrayForSlide.ElementAt(3).isMainGoal ? "F5F5F5" : "FFFFFF";
            // ********************************* Custom Code Changes End *********************************

            A.RgbColorModelHex rgbColorModelHex4201 = new A.RgbColorModelHex() { Val = colorCode4201 };

            solidFill4201.Append(rgbColorModelHex4201);

            tableCellProperties15.Append(leftBorderLineProperties15);
            tableCellProperties15.Append(rightBorderLineProperties15);
            tableCellProperties15.Append(topBorderLineProperties15);
            tableCellProperties15.Append(bottomBorderLineProperties15);
            tableCellProperties15.Append(topLeftToBottomRightBorderLineProperties15);
            tableCellProperties15.Append(bottomLeftToTopRightBorderLineProperties15);
            tableCellProperties15.Append(solidFill4201);

            tableCell15.Append(textBody20);
            tableCell15.Append(tableCellProperties15);

            A.ExtensionList extensionList8 = new A.ExtensionList();

            A.Extension extension8 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement19 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"3973020123\" />");

            // extension8.Append(openXmlUnknownElement19);

            extensionList8.Append(extension8);

            tableRow5.Append(tableCell13);
            tableRow5.Append(tableCell14);
            tableRow5.Append(tableCell15);
            tableRow5.Append(extensionList8);

            A.TableRow tableRow6 = new A.TableRow() { Height = 777502L };

            A.TableCell tableCell16 = new A.TableCell();

            A.TextBody textBody5001 = new A.TextBody();
            A.BodyProperties bodyProperties5001 = new A.BodyProperties();
            A.ListStyle listStyle5001 = new A.ListStyle();

            A.Paragraph paragraph5001 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties5001 = new A.ParagraphProperties() { LeftMargin = 171450, RightMargin = 0, Indent = -171450, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing5001 = new A.LineSpacing();
            A.SpacingPercent spacingPercent5001 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing5001.Append(spacingPercent5001);

            A.SpaceBefore spaceBefore5001 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints5001 = new A.SpacingPoints() { Val = 0 };

            spaceBefore5001.Append(spacingPoints5001);

            A.SpaceAfter spaceAfter5001 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints5002 = new A.SpacingPoints() { Val = 0 };

            spaceAfter5001.Append(spacingPoints5002);
            A.BulletColorText bulletColorText5001 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText5001 = new A.BulletSizeText();
            A.BulletFontText bulletFontText5001 = new A.BulletFontText();

            A.PictureBullet pictureBullet5001 = new A.PictureBullet();

            A.Blip blip5001 = new A.Blip() { Embed = "rId4" };

            A.BlipExtensionList blipExtensionList5001 = new A.BlipExtensionList();

            A.BlipExtension blipExtension5001 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement5001 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId5\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension5001.Append(openXmlUnknownElement5001);

            blipExtensionList5001.Append(blipExtension5001);

            blip5001.Append(blipExtensionList5001);

            pictureBullet5001.Append(blip5001);
            A.TabStopList tabStopList5001 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties5001 = new A.DefaultRunProperties();

            paragraphProperties5001.Append(lineSpacing5001);
            paragraphProperties5001.Append(spaceBefore5001);
            paragraphProperties5001.Append(spaceAfter5001);
            paragraphProperties5001.Append(bulletColorText5001);
            paragraphProperties5001.Append(bulletSizeText5001);
            paragraphProperties5001.Append(bulletFontText5001);
            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count > 4 && !goalsArrayForSlide.ElementAt(4).isMainGoal)
                paragraphProperties5001.Append(pictureBullet5001);
            // ********************************* Custom Code Changes End *********************************
            paragraphProperties5001.Append(tabStopList5001);
            paragraphProperties5001.Append(defaultRunProperties5001);

            A.Run run5001 = new A.Run();

            // ********************************* Custom Code Changes Start *********************************
            int fontSize5001 = goalsArrayForSlide.Count > 4 && goalsArrayForSlide.ElementAt(4).isMainGoal ? 1300 : 1200;
            bool isBold5001 = goalsArrayForSlide.Count > 4 && goalsArrayForSlide.ElementAt(4).isMainGoal;
            string fontName5001 = goalsArrayForSlide.Count > 4 && goalsArrayForSlide.ElementAt(4).isMainGoal ? "Segoe UI Semibold" : "Segoe UI";
            // ********************************* Custom Code Changes End *********************************

            A.RunProperties runProperties5001 = new A.RunProperties() { Language = "en-US", FontSize = fontSize5001, Bold = isBold5001, Italic = false };
            A.EffectList effectList5001 = new A.EffectList();
            A.LatinFont latinFont5001 = new A.LatinFont() { Typeface = fontName5001 };
            A.ComplexScriptFont complexScriptFont5001 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            runProperties5001.Append(effectList5001);
            runProperties5001.Append(latinFont5001);
            runProperties5001.Append(complexScriptFont5001);
            A.Text text5001 = new A.Text();

            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count() > 4)
                text5001.Text = goalsArrayForSlide.ElementAt(4).goalName ?? "";
            else
                text5001.Text = "NO DATA";
            // ********************************* Custom Code Changes End *********************************

            run5001.Append(runProperties5001);
            run5001.Append(text5001);

            A.EndParagraphRunProperties endParagraphRunProperties5001 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false };
            A.LatinFont latinFont5002 = new A.LatinFont() { Typeface = "Segoe UI" };
            A.ComplexScriptFont complexScriptFont5002 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            endParagraphRunProperties5001.Append(latinFont5002);
            endParagraphRunProperties5001.Append(complexScriptFont5002);

            paragraph5001.Append(paragraphProperties5001);
            paragraph5001.Append(run5001);
            paragraph5001.Append(endParagraphRunProperties5001);

            textBody5001.Append(bodyProperties5001);
            textBody5001.Append(listStyle5001);
            textBody5001.Append(paragraph5001);

            A.TableCellProperties tableCellProperties5001 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties5001 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill5001 = new A.NoFill();

            leftBorderLineProperties5001.Append(noFill5001);

            A.RightBorderLineProperties rightBorderLineProperties5001 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill5002 = new A.NoFill();

            rightBorderLineProperties5001.Append(noFill5002);

            A.TopBorderLineProperties topBorderLineProperties5001 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill5003 = new A.NoFill();
            A.PresetDash presetDash5001 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round5001 = new A.Round();
            A.HeadEnd headEnd5001 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd5001 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties5001.Append(noFill5003);
            topBorderLineProperties5001.Append(presetDash5001);
            topBorderLineProperties5001.Append(round5001);
            topBorderLineProperties5001.Append(headEnd5001);
            topBorderLineProperties5001.Append(tailEnd5001);

            A.BottomBorderLineProperties bottomBorderLineProperties5001 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill5004 = new A.NoFill();
            A.PresetDash presetDash5002 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round5002 = new A.Round();
            A.HeadEnd headEnd5002 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd5002 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties5001.Append(noFill5004);
            bottomBorderLineProperties5001.Append(presetDash5002);
            bottomBorderLineProperties5001.Append(round5002);
            bottomBorderLineProperties5001.Append(headEnd5002);
            bottomBorderLineProperties5001.Append(tailEnd5002);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties5001 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill5005 = new A.NoFill();
            A.PresetDash presetDash5003 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties5001.Append(noFill5005);
            topLeftToBottomRightBorderLineProperties5001.Append(presetDash5003);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties5001 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill5006 = new A.NoFill();
            A.PresetDash presetDash5004 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties5001.Append(noFill5006);
            bottomLeftToTopRightBorderLineProperties5001.Append(presetDash5004);

            A.SolidFill solidFill5001 = new A.SolidFill();

            // ********************************* Custom Code Changes Start *********************************
            // Adding the background color according to the goal4 type
            string goalColor5001 = goalsArrayForSlide.Count > 4 && goalsArrayForSlide.ElementAt(4).isMainGoal ? "F5F5F5" : "FFFFFF";
            A.RgbColorModelHex rgbColorModelHex5001 = new A.RgbColorModelHex() { Val = goalColor5001 };
            // ********************************* Custom Code Changes End *********************************

            solidFill5001.Append(rgbColorModelHex5001);

            tableCellProperties5001.Append(leftBorderLineProperties5001);
            tableCellProperties5001.Append(rightBorderLineProperties5001);
            tableCellProperties5001.Append(topBorderLineProperties5001);
            tableCellProperties5001.Append(bottomBorderLineProperties5001);
            tableCellProperties5001.Append(topLeftToBottomRightBorderLineProperties5001);
            tableCellProperties5001.Append(bottomLeftToTopRightBorderLineProperties5001);
            tableCellProperties5001.Append(solidFill5001);

            tableCell16.Append(textBody5001);
            tableCell16.Append(tableCellProperties5001);

            A.TableCell tableCell17 = new A.TableCell();

            A.TextBody textBody22 = new A.TextBody();
            A.BodyProperties bodyProperties22 = new A.BodyProperties();
            A.ListStyle listStyle22 = new A.ListStyle();

            A.Paragraph paragraph22 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties16 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing15 = new A.LineSpacing();
            A.SpacingPercent spacingPercent15 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing15.Append(spacingPercent15);

            A.SpaceBefore spaceBefore13 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints25 = new A.SpacingPoints() { Val = 0 };

            spaceBefore13.Append(spacingPoints25);

            A.SpaceAfter spaceAfter13 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints26 = new A.SpacingPoints() { Val = 0 };

            spaceAfter13.Append(spacingPoints26);
            A.BulletColorText bulletColorText13 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText13 = new A.BulletSizeText();
            A.BulletFontText bulletFontText14 = new A.BulletFontText();
            A.NoBullet noBullet11 = new A.NoBullet();
            A.TabStopList tabStopList12 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties22 = new A.DefaultRunProperties();

            paragraphProperties16.Append(lineSpacing15);
            paragraphProperties16.Append(spaceBefore13);
            paragraphProperties16.Append(spaceAfter13);
            paragraphProperties16.Append(bulletColorText13);
            paragraphProperties16.Append(bulletSizeText13);
            paragraphProperties16.Append(bulletFontText14);
            paragraphProperties16.Append(noBullet11);
            paragraphProperties16.Append(tabStopList12);
            paragraphProperties16.Append(defaultRunProperties22);

            A.EndParagraphRunProperties endParagraphRunProperties17 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false, Dirty = false };

            A.Outline outline14 = new A.Outline();
            A.NoFill noFill110 = new A.NoFill();

            outline14.Append(noFill110);

            A.SolidFill solidFill46 = new A.SolidFill();
            A.PresetColor presetColor7 = new A.PresetColor() { Val = A.PresetColorValues.Black };

            solidFill46.Append(presetColor7);
            A.EffectList effectList18 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText11 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText11 = new A.UnderlineFillText();
            A.LatinFont latinFont34 = new A.LatinFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont26 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont34 = new A.ComplexScriptFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties17.Append(outline14);
            endParagraphRunProperties17.Append(solidFill46);
            endParagraphRunProperties17.Append(effectList18);
            endParagraphRunProperties17.Append(underlineFollowsText11);
            endParagraphRunProperties17.Append(underlineFillText11);
            endParagraphRunProperties17.Append(latinFont34);
            endParagraphRunProperties17.Append(eastAsianFont26);
            endParagraphRunProperties17.Append(complexScriptFont34);

            paragraph22.Append(paragraphProperties16);
            paragraph22.Append(endParagraphRunProperties17);

            textBody22.Append(bodyProperties22);
            textBody22.Append(listStyle22);
            textBody22.Append(paragraph22);

            A.TableCellProperties tableCellProperties17 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties17 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill111 = new A.NoFill();

            leftBorderLineProperties17.Append(noFill111);

            A.RightBorderLineProperties rightBorderLineProperties17 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill112 = new A.NoFill();

            rightBorderLineProperties17.Append(noFill112);

            A.TopBorderLineProperties topBorderLineProperties17 = new A.TopBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill113 = new A.NoFill();
            A.PresetDash presetDash62 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round30 = new A.Round();
            A.HeadEnd headEnd30 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd30 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties17.Append(noFill113);
            topBorderLineProperties17.Append(presetDash62);
            topBorderLineProperties17.Append(round30);
            topBorderLineProperties17.Append(headEnd30);
            topBorderLineProperties17.Append(tailEnd30);

            A.BottomBorderLineProperties bottomBorderLineProperties17 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill114 = new A.NoFill();
            A.PresetDash presetDash63 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round31 = new A.Round();
            A.HeadEnd headEnd31 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd31 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties17.Append(noFill114);
            bottomBorderLineProperties17.Append(presetDash63);
            bottomBorderLineProperties17.Append(round31);
            bottomBorderLineProperties17.Append(headEnd31);
            bottomBorderLineProperties17.Append(tailEnd31);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties17 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill115 = new A.NoFill();
            A.PresetDash presetDash64 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties17.Append(noFill115);
            topLeftToBottomRightBorderLineProperties17.Append(presetDash64);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties17 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill116 = new A.NoFill();
            A.PresetDash presetDash65 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties17.Append(noFill116);
            bottomLeftToTopRightBorderLineProperties17.Append(presetDash65);

            A.SolidFill solidFill5101 = new A.SolidFill();
            // ********************************* Custom Code Changes Start *********************************
            string colorCode5101 = goalsArrayForSlide.Count > 4 && goalsArrayForSlide.ElementAt(4).isMainGoal ? "F5F5F5" : "FFFFFF";
            // ********************************* Custom Code Changes End *********************************

            A.RgbColorModelHex rgbColorModelHex5101 = new A.RgbColorModelHex() { Val = colorCode5101 };

            solidFill5101.Append(rgbColorModelHex5101);

            tableCellProperties17.Append(leftBorderLineProperties17);
            tableCellProperties17.Append(rightBorderLineProperties17);
            tableCellProperties17.Append(topBorderLineProperties17);
            tableCellProperties17.Append(bottomBorderLineProperties17);
            tableCellProperties17.Append(topLeftToBottomRightBorderLineProperties17);
            tableCellProperties17.Append(bottomLeftToTopRightBorderLineProperties17);
            tableCellProperties17.Append(solidFill5101);

            tableCell17.Append(textBody22);
            tableCell17.Append(tableCellProperties17);

            A.TableCell tableCell18 = new A.TableCell();

            A.TextBody textBody23 = new A.TextBody();
            A.BodyProperties bodyProperties23 = new A.BodyProperties();
            A.ListStyle listStyle23 = new A.ListStyle();

            A.Paragraph paragraph23 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties17 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing16 = new A.LineSpacing();
            A.SpacingPercent spacingPercent16 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing16.Append(spacingPercent16);

            A.SpaceBefore spaceBefore14 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints27 = new A.SpacingPoints() { Val = 0 };

            spaceBefore14.Append(spacingPoints27);

            A.SpaceAfter spaceAfter14 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints28 = new A.SpacingPoints() { Val = 0 };

            spaceAfter14.Append(spacingPoints28);
            A.BulletColorText bulletColorText14 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText14 = new A.BulletSizeText();
            A.BulletFontText bulletFontText15 = new A.BulletFontText();
            A.NoBullet noBullet12 = new A.NoBullet();
            A.TabStopList tabStopList13 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties23 = new A.DefaultRunProperties();

            paragraphProperties17.Append(lineSpacing16);
            paragraphProperties17.Append(spaceBefore14);
            paragraphProperties17.Append(spaceAfter14);
            paragraphProperties17.Append(bulletColorText14);
            paragraphProperties17.Append(bulletSizeText14);
            paragraphProperties17.Append(bulletFontText15);
            paragraphProperties17.Append(noBullet12);
            paragraphProperties17.Append(tabStopList13);
            paragraphProperties17.Append(defaultRunProperties23);

            A.EndParagraphRunProperties endParagraphRunProperties18 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline15 = new A.Outline();
            A.NoFill noFill117 = new A.NoFill();

            outline15.Append(noFill117);

            A.SolidFill solidFill48 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "C50F1F" };

            solidFill48.Append(rgbColorModelHex12);
            A.EffectList effectList19 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText12 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText12 = new A.UnderlineFillText();
            A.LatinFont latinFont35 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont27 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont35 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties18.Append(outline15);
            endParagraphRunProperties18.Append(solidFill48);
            endParagraphRunProperties18.Append(effectList19);
            endParagraphRunProperties18.Append(underlineFollowsText12);
            endParagraphRunProperties18.Append(underlineFillText12);
            endParagraphRunProperties18.Append(latinFont35);
            endParagraphRunProperties18.Append(eastAsianFont27);
            endParagraphRunProperties18.Append(complexScriptFont35);

            paragraph23.Append(paragraphProperties17);
            paragraph23.Append(endParagraphRunProperties18);

            textBody23.Append(bodyProperties23);
            textBody23.Append(listStyle23);
            textBody23.Append(paragraph23);

            A.TableCellProperties tableCellProperties18 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties18 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill118 = new A.NoFill();

            leftBorderLineProperties18.Append(noFill118);

            A.RightBorderLineProperties rightBorderLineProperties18 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill119 = new A.NoFill();

            rightBorderLineProperties18.Append(noFill119);

            A.TopBorderLineProperties topBorderLineProperties18 = new A.TopBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill120 = new A.NoFill();
            A.PresetDash presetDash66 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round32 = new A.Round();
            A.HeadEnd headEnd32 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd32 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties18.Append(noFill120);
            topBorderLineProperties18.Append(presetDash66);
            topBorderLineProperties18.Append(round32);
            topBorderLineProperties18.Append(headEnd32);
            topBorderLineProperties18.Append(tailEnd32);

            A.BottomBorderLineProperties bottomBorderLineProperties18 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill121 = new A.NoFill();
            A.PresetDash presetDash67 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round33 = new A.Round();
            A.HeadEnd headEnd33 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd33 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties18.Append(noFill121);
            bottomBorderLineProperties18.Append(presetDash67);
            bottomBorderLineProperties18.Append(round33);
            bottomBorderLineProperties18.Append(headEnd33);
            bottomBorderLineProperties18.Append(tailEnd33);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties18 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill122 = new A.NoFill();
            A.PresetDash presetDash68 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties18.Append(noFill122);
            topLeftToBottomRightBorderLineProperties18.Append(presetDash68);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties18 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill123 = new A.NoFill();
            A.PresetDash presetDash69 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties18.Append(noFill123);
            bottomLeftToTopRightBorderLineProperties18.Append(presetDash69);

            A.SolidFill solidFill5201 = new A.SolidFill();
            // ********************************* Custom Code Changes Start *********************************
            string colorCode5201 = goalsArrayForSlide.Count > 4 && goalsArrayForSlide.ElementAt(4).isMainGoal ? "F5F5F5" : "FFFFFF";
            // ********************************* Custom Code Changes End *********************************

            A.RgbColorModelHex rgbColorModelHex5201 = new A.RgbColorModelHex() { Val = colorCode5201 };

            solidFill5201.Append(rgbColorModelHex5201);

            tableCellProperties18.Append(leftBorderLineProperties18);
            tableCellProperties18.Append(rightBorderLineProperties18);
            tableCellProperties18.Append(topBorderLineProperties18);
            tableCellProperties18.Append(bottomBorderLineProperties18);
            tableCellProperties18.Append(topLeftToBottomRightBorderLineProperties18);
            tableCellProperties18.Append(bottomLeftToTopRightBorderLineProperties18);
            tableCellProperties18.Append(solidFill5201);

            tableCell18.Append(textBody23);
            tableCell18.Append(tableCellProperties18);

            A.ExtensionList extensionList9 = new A.ExtensionList();

            A.Extension extension9 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement20 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"3811565737\" />");

            // extension9.Append(openXmlUnknownElement20);

            extensionList9.Append(extension9);

            tableRow6.Append(tableCell16);
            tableRow6.Append(tableCell17);
            tableRow6.Append(tableCell18);
            tableRow6.Append(extensionList9);

            A.TableRow tableRow7 = new A.TableRow() { Height = 891827L };

            A.TableCell tableCell19 = new A.TableCell();

            A.TextBody textBody6001 = new A.TextBody();
            A.BodyProperties bodyProperties6001 = new A.BodyProperties();
            A.ListStyle listStyle6001 = new A.ListStyle();

            A.Paragraph paragraph6001 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties6001 = new A.ParagraphProperties() { LeftMargin = 171450, RightMargin = 0, Indent = -171450, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing6001 = new A.LineSpacing();
            A.SpacingPercent spacingPercent6001 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing6001.Append(spacingPercent6001);

            A.SpaceBefore spaceBefore6001 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints6001 = new A.SpacingPoints() { Val = 0 };

            spaceBefore6001.Append(spacingPoints6001);

            A.SpaceAfter spaceAfter6001 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints6002 = new A.SpacingPoints() { Val = 0 };

            spaceAfter6001.Append(spacingPoints6002);
            A.BulletColorText bulletColorText6001 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText6001 = new A.BulletSizeText();
            A.BulletFontText bulletFontText6001 = new A.BulletFontText();

            A.PictureBullet pictureBullet6001 = new A.PictureBullet();

            A.Blip blip6001 = new A.Blip() { Embed = "rId4" };

            A.BlipExtensionList blipExtensionList6001 = new A.BlipExtensionList();

            A.BlipExtension blipExtension6001 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement6001 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId5\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension6001.Append(openXmlUnknownElement6001);

            blipExtensionList6001.Append(blipExtension6001);

            blip6001.Append(blipExtensionList6001);

            pictureBullet6001.Append(blip6001);
            A.TabStopList tabStopList6001 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties6001 = new A.DefaultRunProperties();

            paragraphProperties6001.Append(lineSpacing6001);
            paragraphProperties6001.Append(spaceBefore6001);
            paragraphProperties6001.Append(spaceAfter6001);
            paragraphProperties6001.Append(bulletColorText6001);
            paragraphProperties6001.Append(bulletSizeText6001);
            paragraphProperties6001.Append(bulletFontText6001);
            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count > 5 && !goalsArrayForSlide.ElementAt(5).isMainGoal)
                paragraphProperties6001.Append(pictureBullet6001);
            // ********************************* Custom Code Changes End *********************************
            paragraphProperties6001.Append(tabStopList6001);
            paragraphProperties6001.Append(defaultRunProperties6001);

            A.Run run6001 = new A.Run();

            // ********************************* Custom Code Changes Start *********************************
            int fontSize6001 = goalsArrayForSlide.Count > 5 && goalsArrayForSlide.ElementAt(5).isMainGoal ? 1300 : 1200;
            bool isBold6001 = goalsArrayForSlide.Count > 5 && goalsArrayForSlide.ElementAt(5).isMainGoal;
            string fontName6001 = goalsArrayForSlide.Count > 5 && goalsArrayForSlide.ElementAt(5).isMainGoal ? "Segoe UI Semibold" : "Segoe UI";
            // ********************************* Custom Code Changes End *********************************

            A.RunProperties runProperties6001 = new A.RunProperties() { Language = "en-US", FontSize = fontSize6001, Bold = isBold6001, Italic = false };
            A.EffectList effectList6001 = new A.EffectList();
            A.LatinFont latinFont6001 = new A.LatinFont() { Typeface = fontName6001 };
            A.ComplexScriptFont complexScriptFont6001 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            runProperties6001.Append(effectList6001);
            runProperties6001.Append(latinFont6001);
            runProperties6001.Append(complexScriptFont6001);
            A.Text text6001 = new A.Text();

            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count() > 5)
                text6001.Text = goalsArrayForSlide.ElementAt(5).goalName ?? "";
            else
                text6001.Text = "NO DATA";
            // ********************************* Custom Code Changes End *********************************

            run6001.Append(runProperties6001);
            run6001.Append(text6001);

            A.EndParagraphRunProperties endParagraphRunProperties6001 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false };
            A.LatinFont latinFont6002 = new A.LatinFont() { Typeface = "Segoe UI" };
            A.ComplexScriptFont complexScriptFont6002 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            endParagraphRunProperties6001.Append(latinFont6002);
            endParagraphRunProperties6001.Append(complexScriptFont6002);

            paragraph6001.Append(paragraphProperties6001);
            paragraph6001.Append(run6001);
            paragraph6001.Append(endParagraphRunProperties6001);

            textBody6001.Append(bodyProperties6001);
            textBody6001.Append(listStyle6001);
            textBody6001.Append(paragraph6001);

            A.TableCellProperties tableCellProperties6001 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties6001 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill6001 = new A.NoFill();

            leftBorderLineProperties6001.Append(noFill6001);

            A.RightBorderLineProperties rightBorderLineProperties6001 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill6002 = new A.NoFill();

            rightBorderLineProperties6001.Append(noFill6002);

            A.TopBorderLineProperties topBorderLineProperties6001 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill6003 = new A.NoFill();
            A.PresetDash presetDash6001 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round6001 = new A.Round();
            A.HeadEnd headEnd6001 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd6001 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties6001.Append(noFill6003);
            topBorderLineProperties6001.Append(presetDash6001);
            topBorderLineProperties6001.Append(round6001);
            topBorderLineProperties6001.Append(headEnd6001);
            topBorderLineProperties6001.Append(tailEnd6001);

            A.BottomBorderLineProperties bottomBorderLineProperties6001 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill6004 = new A.NoFill();
            A.PresetDash presetDash6002 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round6002 = new A.Round();
            A.HeadEnd headEnd6002 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd6002 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties6001.Append(noFill6004);
            bottomBorderLineProperties6001.Append(presetDash6002);
            bottomBorderLineProperties6001.Append(round6002);
            bottomBorderLineProperties6001.Append(headEnd6002);
            bottomBorderLineProperties6001.Append(tailEnd6002);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties6001 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill6005 = new A.NoFill();
            A.PresetDash presetDash6003 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties6001.Append(noFill6005);
            topLeftToBottomRightBorderLineProperties6001.Append(presetDash6003);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties6001 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill6006 = new A.NoFill();
            A.PresetDash presetDash6004 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties6001.Append(noFill6006);
            bottomLeftToTopRightBorderLineProperties6001.Append(presetDash6004);

            A.SolidFill solidFill6001 = new A.SolidFill();

            // ********************************* Custom Code Changes Start *********************************
            // Adding the background color according to the goal6 type
            string goalColor6001 = goalsArrayForSlide.Count > 5 && goalsArrayForSlide.ElementAt(5).isMainGoal ? "F5F5F5" : "FFFFFF";
            A.RgbColorModelHex rgbColorModelHex6001 = new A.RgbColorModelHex() { Val = goalColor6001 };
            // ********************************* Custom Code Changes End *********************************

            solidFill6001.Append(rgbColorModelHex6001);

            tableCellProperties6001.Append(leftBorderLineProperties6001);
            tableCellProperties6001.Append(rightBorderLineProperties6001);
            tableCellProperties6001.Append(topBorderLineProperties6001);
            tableCellProperties6001.Append(bottomBorderLineProperties6001);
            tableCellProperties6001.Append(topLeftToBottomRightBorderLineProperties6001);
            tableCellProperties6001.Append(bottomLeftToTopRightBorderLineProperties6001);
            tableCellProperties6001.Append(solidFill6001);

            tableCell19.Append(textBody6001);
            tableCell19.Append(tableCellProperties6001);


            A.TableCell tableCell20 = new A.TableCell();

            A.TextBody textBody25 = new A.TextBody();
            A.BodyProperties bodyProperties25 = new A.BodyProperties();
            A.ListStyle listStyle25 = new A.ListStyle();

            A.Paragraph paragraph25 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties19 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing17 = new A.LineSpacing();
            A.SpacingPercent spacingPercent17 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing17.Append(spacingPercent17);

            A.SpaceBefore spaceBefore15 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints29 = new A.SpacingPoints() { Val = 0 };

            spaceBefore15.Append(spacingPoints29);

            A.SpaceAfter spaceAfter15 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints30 = new A.SpacingPoints() { Val = 0 };

            spaceAfter15.Append(spacingPoints30);
            A.BulletColorText bulletColorText15 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText15 = new A.BulletSizeText();
            A.BulletFontText bulletFontText17 = new A.BulletFontText();
            A.NoBullet noBullet13 = new A.NoBullet();
            A.TabStopList tabStopList14 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties24 = new A.DefaultRunProperties();

            paragraphProperties19.Append(lineSpacing17);
            paragraphProperties19.Append(spaceBefore15);
            paragraphProperties19.Append(spaceAfter15);
            paragraphProperties19.Append(bulletColorText15);
            paragraphProperties19.Append(bulletSizeText15);
            paragraphProperties19.Append(bulletFontText17);
            paragraphProperties19.Append(noBullet13);
            paragraphProperties19.Append(tabStopList14);
            paragraphProperties19.Append(defaultRunProperties24);

            A.EndParagraphRunProperties endParagraphRunProperties20 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline16 = new A.Outline();
            A.NoFill noFill130 = new A.NoFill();

            outline16.Append(noFill130);

            A.SolidFill solidFill51 = new A.SolidFill();
            A.PresetColor presetColor8 = new A.PresetColor() { Val = A.PresetColorValues.Black };

            solidFill51.Append(presetColor8);
            A.EffectList effectList21 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText13 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText13 = new A.UnderlineFillText();
            A.LatinFont latinFont38 = new A.LatinFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont28 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont38 = new A.ComplexScriptFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties20.Append(outline16);
            endParagraphRunProperties20.Append(solidFill51);
            endParagraphRunProperties20.Append(effectList21);
            endParagraphRunProperties20.Append(underlineFollowsText13);
            endParagraphRunProperties20.Append(underlineFillText13);
            endParagraphRunProperties20.Append(latinFont38);
            endParagraphRunProperties20.Append(eastAsianFont28);
            endParagraphRunProperties20.Append(complexScriptFont38);

            paragraph25.Append(paragraphProperties19);
            paragraph25.Append(endParagraphRunProperties20);

            textBody25.Append(bodyProperties25);
            textBody25.Append(listStyle25);
            textBody25.Append(paragraph25);

            A.TableCellProperties tableCellProperties20 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties20 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill131 = new A.NoFill();

            leftBorderLineProperties20.Append(noFill131);

            A.RightBorderLineProperties rightBorderLineProperties20 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill132 = new A.NoFill();

            rightBorderLineProperties20.Append(noFill132);

            A.TopBorderLineProperties topBorderLineProperties20 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill133 = new A.NoFill();
            A.PresetDash presetDash74 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round36 = new A.Round();
            A.HeadEnd headEnd36 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd36 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties20.Append(noFill133);
            topBorderLineProperties20.Append(presetDash74);
            topBorderLineProperties20.Append(round36);
            topBorderLineProperties20.Append(headEnd36);
            topBorderLineProperties20.Append(tailEnd36);

            A.BottomBorderLineProperties bottomBorderLineProperties20 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill134 = new A.NoFill();
            A.PresetDash presetDash75 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round37 = new A.Round();
            A.HeadEnd headEnd37 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd37 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties20.Append(noFill134);
            bottomBorderLineProperties20.Append(presetDash75);
            bottomBorderLineProperties20.Append(round37);
            bottomBorderLineProperties20.Append(headEnd37);
            bottomBorderLineProperties20.Append(tailEnd37);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties20 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill135 = new A.NoFill();
            A.PresetDash presetDash76 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties20.Append(noFill135);
            topLeftToBottomRightBorderLineProperties20.Append(presetDash76);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties20 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill136 = new A.NoFill();
            A.PresetDash presetDash77 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties20.Append(noFill136);
            bottomLeftToTopRightBorderLineProperties20.Append(presetDash77);

            A.SolidFill solidFill6101 = new A.SolidFill();
            // ********************************* Custom Code Changes Start *********************************
            string colorCode6101 = goalsArrayForSlide.Count > 5 && goalsArrayForSlide.ElementAt(5).isMainGoal ? "F5F5F5" : "FFFFFF";
            // ********************************* Custom Code Changes End *********************************

            A.RgbColorModelHex rgbColorModelHex6101 = new A.RgbColorModelHex() { Val = colorCode6101 };

            solidFill6101.Append(rgbColorModelHex6101);

            tableCellProperties20.Append(leftBorderLineProperties20);
            tableCellProperties20.Append(rightBorderLineProperties20);
            tableCellProperties20.Append(topBorderLineProperties20);
            tableCellProperties20.Append(bottomBorderLineProperties20);
            tableCellProperties20.Append(topLeftToBottomRightBorderLineProperties20);
            tableCellProperties20.Append(bottomLeftToTopRightBorderLineProperties20);
            tableCellProperties20.Append(solidFill6101);

            tableCell20.Append(textBody25);
            tableCell20.Append(tableCellProperties20);

            A.TableCell tableCell21 = new A.TableCell();

            A.TextBody textBody26 = new A.TextBody();
            A.BodyProperties bodyProperties26 = new A.BodyProperties();
            A.ListStyle listStyle26 = new A.ListStyle();

            A.Paragraph paragraph26 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties20 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing18 = new A.LineSpacing();
            A.SpacingPercent spacingPercent18 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing18.Append(spacingPercent18);

            A.SpaceBefore spaceBefore16 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints31 = new A.SpacingPoints() { Val = 0 };

            spaceBefore16.Append(spacingPoints31);

            A.SpaceAfter spaceAfter16 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints32 = new A.SpacingPoints() { Val = 0 };

            spaceAfter16.Append(spacingPoints32);
            A.BulletColorText bulletColorText16 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText16 = new A.BulletSizeText();
            A.BulletFontText bulletFontText18 = new A.BulletFontText();
            A.NoBullet noBullet14 = new A.NoBullet();
            A.TabStopList tabStopList15 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties25 = new A.DefaultRunProperties();

            paragraphProperties20.Append(lineSpacing18);
            paragraphProperties20.Append(spaceBefore16);
            paragraphProperties20.Append(spaceAfter16);
            paragraphProperties20.Append(bulletColorText16);
            paragraphProperties20.Append(bulletSizeText16);
            paragraphProperties20.Append(bulletFontText18);
            paragraphProperties20.Append(noBullet14);
            paragraphProperties20.Append(tabStopList15);
            paragraphProperties20.Append(defaultRunProperties25);

            A.EndParagraphRunProperties endParagraphRunProperties21 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false, Dirty = false };

            A.Outline outline17 = new A.Outline();
            A.NoFill noFill137 = new A.NoFill();

            outline17.Append(noFill137);

            A.SolidFill solidFill53 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex14 = new A.RgbColorModelHex() { Val = "5EC65A" };

            solidFill53.Append(rgbColorModelHex14);
            A.EffectList effectList22 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText14 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText14 = new A.UnderlineFillText();
            A.LatinFont latinFont39 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont29 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont39 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties21.Append(outline17);
            endParagraphRunProperties21.Append(solidFill53);
            endParagraphRunProperties21.Append(effectList22);
            endParagraphRunProperties21.Append(underlineFollowsText14);
            endParagraphRunProperties21.Append(underlineFillText14);
            endParagraphRunProperties21.Append(latinFont39);
            endParagraphRunProperties21.Append(eastAsianFont29);
            endParagraphRunProperties21.Append(complexScriptFont39);

            paragraph26.Append(paragraphProperties20);
            paragraph26.Append(endParagraphRunProperties21);

            textBody26.Append(bodyProperties26);
            textBody26.Append(listStyle26);
            textBody26.Append(paragraph26);

            A.TableCellProperties tableCellProperties21 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties21 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill138 = new A.NoFill();

            leftBorderLineProperties21.Append(noFill138);

            A.RightBorderLineProperties rightBorderLineProperties21 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill139 = new A.NoFill();

            rightBorderLineProperties21.Append(noFill139);

            A.TopBorderLineProperties topBorderLineProperties21 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill140 = new A.NoFill();
            A.PresetDash presetDash78 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round38 = new A.Round();
            A.HeadEnd headEnd38 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd38 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties21.Append(noFill140);
            topBorderLineProperties21.Append(presetDash78);
            topBorderLineProperties21.Append(round38);
            topBorderLineProperties21.Append(headEnd38);
            topBorderLineProperties21.Append(tailEnd38);

            A.BottomBorderLineProperties bottomBorderLineProperties21 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill141 = new A.NoFill();
            A.PresetDash presetDash79 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round39 = new A.Round();
            A.HeadEnd headEnd39 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd39 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties21.Append(noFill141);
            bottomBorderLineProperties21.Append(presetDash79);
            bottomBorderLineProperties21.Append(round39);
            bottomBorderLineProperties21.Append(headEnd39);
            bottomBorderLineProperties21.Append(tailEnd39);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties21 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill142 = new A.NoFill();
            A.PresetDash presetDash80 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties21.Append(noFill142);
            topLeftToBottomRightBorderLineProperties21.Append(presetDash80);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties21 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill143 = new A.NoFill();
            A.PresetDash presetDash81 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties21.Append(noFill143);
            bottomLeftToTopRightBorderLineProperties21.Append(presetDash81);

            A.SolidFill solidFill6201 = new A.SolidFill();
            // ********************************* Custom Code Changes Start *********************************
            string colorCode6201 = goalsArrayForSlide.Count > 5 && goalsArrayForSlide.ElementAt(5).isMainGoal ? "F5F5F5" : "FFFFFF";
            // ********************************* Custom Code Changes End *********************************

            A.RgbColorModelHex rgbColorModelHex6201 = new A.RgbColorModelHex() { Val = colorCode6201 };

            solidFill6201.Append(rgbColorModelHex6201);

            tableCellProperties21.Append(leftBorderLineProperties21);
            tableCellProperties21.Append(rightBorderLineProperties21);
            tableCellProperties21.Append(topBorderLineProperties21);
            tableCellProperties21.Append(bottomBorderLineProperties21);
            tableCellProperties21.Append(topLeftToBottomRightBorderLineProperties21);
            tableCellProperties21.Append(bottomLeftToTopRightBorderLineProperties21);
            tableCellProperties21.Append(solidFill6201);

            tableCell21.Append(textBody26);
            tableCell21.Append(tableCellProperties21);

            A.ExtensionList extensionList10 = new A.ExtensionList();

            A.Extension extension10 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement22 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"21430747\" />");

            // extension10.Append(openXmlUnknownElement22);

            extensionList10.Append(extension10);

            tableRow7.Append(tableCell19);
            tableRow7.Append(tableCell20);
            tableRow7.Append(tableCell21);
            tableRow7.Append(extensionList10);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);

            // ********************************* Custom Code Changes Start *********************************
            A.TableRow[] tableRows = [tableRow2, tableRow3, tableRow4, tableRow5, tableRow6, tableRow7];

            for (int i = 0; i < goalsArrayForSlide.Count(); i++)
            {
                table1.Append(tableRows[i]);
            }
            // ********************************* Custom Code Changes End *********************************

            graphicData1.Append(table1);

            graphic1.Append(graphicData1);

            graphicFrame1.Append(nonVisualGraphicFrameProperties1);
            graphicFrame1.Append(transform1);
            graphicFrame1.Append(graphic1);

            Shape shape6 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties6 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties10 = new NonVisualDrawingProperties() { Id = (UInt32Value)10U, Name = "Text 46" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList9 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension9 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement23 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{55972656-B818-B1DD-4415-6E95B0D8F09A}\" />");

            // nonVisualDrawingPropertiesExtension9.Append(openXmlUnknownElement23);

            nonVisualDrawingPropertiesExtensionList9.Append(nonVisualDrawingPropertiesExtension9);

            nonVisualDrawingProperties10.Append(nonVisualDrawingPropertiesExtensionList9);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties6 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties10 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties6.Append(nonVisualDrawingProperties10);
            nonVisualShapeProperties6.Append(nonVisualShapeDrawingProperties6);
            nonVisualShapeProperties6.Append(applicationNonVisualDrawingProperties10);

            ShapeProperties shapeProperties8 = new ShapeProperties();

            A.Transform2D transform2D6 = new A.Transform2D();
            A.Offset offset8 = new A.Offset() { X = 8355350L, Y = 2553866L };
            A.Extents extents8 = new A.Extents() { Cx = 1280895L, Cy = 228571L };

            transform2D6.Append(offset8);
            transform2D6.Append(extents8);

            A.PresetGeometry presetGeometry6 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList6 = new A.AdjustValueList();

            presetGeometry6.Append(adjustValueList6);
            A.NoFill noFill144 = new A.NoFill();
            A.Outline outline18 = new A.Outline();

            shapeProperties8.Append(transform2D6);
            shapeProperties8.Append(presetGeometry6);
            shapeProperties8.Append(noFill144);
            shapeProperties8.Append(outline18);

            TextBody textBody27 = new TextBody();
            A.BodyProperties bodyProperties27 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle27 = new A.ListStyle();

            A.Paragraph paragraph27 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties21 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing19 = new A.LineSpacing();
            A.SpacingPercent spacingPercent19 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing19.Append(spacingPercent19);

            paragraphProperties21.Append(lineSpacing19);

            A.Run run15 = new A.Run();

            A.RunProperties runProperties15 = new A.RunProperties() { Language = "en-US", FontSize = 1080 };

            A.SolidFill solidFill55 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex15 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha4 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex15.Append(alpha4);

            solidFill55.Append(rgbColorModelHex15);
            A.LatinFont latinFont40 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont30 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont40 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties15.Append(solidFill55);
            runProperties15.Append(latinFont40);
            runProperties15.Append(eastAsianFont30);
            runProperties15.Append(complexScriptFont40);
            A.Text text15 = new A.Text();
            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count() > 1)
                text15.Text = goalsArrayForSlide.ElementAt(1).goalOwner ?? "";
            else
                text15.Text = "NO DATA";
            // ********************************* Custom Code Changes End *********************************

            run15.Append(runProperties15);
            run15.Append(text15);
            A.EndParagraphRunProperties endParagraphRunProperties22 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080 };

            paragraph27.Append(paragraphProperties21);
            paragraph27.Append(run15);
            paragraph27.Append(endParagraphRunProperties22);

            textBody27.Append(bodyProperties27);
            textBody27.Append(listStyle27);
            textBody27.Append(paragraph27);

            shape6.Append(nonVisualShapeProperties6);
            shape6.Append(shapeProperties8);
            shape6.Append(textBody27);

            Shape shape7 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties7 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties11 = new NonVisualDrawingProperties() { Id = (UInt32Value)11U, Name = "Text 46" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList10 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension10 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement24 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{071EE7F8-2AE4-842D-F2C4-3ED832EE2ABE}\" />");

            // nonVisualDrawingPropertiesExtension10.Append(openXmlUnknownElement24);

            nonVisualDrawingPropertiesExtensionList10.Append(nonVisualDrawingPropertiesExtension10);

            nonVisualDrawingProperties11.Append(nonVisualDrawingPropertiesExtensionList10);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties7 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties11 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties7.Append(nonVisualDrawingProperties11);
            nonVisualShapeProperties7.Append(nonVisualShapeDrawingProperties7);
            nonVisualShapeProperties7.Append(applicationNonVisualDrawingProperties11);

            ShapeProperties shapeProperties9 = new ShapeProperties();

            A.Transform2D transform2D7 = new A.Transform2D();
            A.Offset offset9 = new A.Offset() { X = 8380023L, Y = 3310389L };
            A.Extents extents9 = new A.Extents() { Cx = 1280895L, Cy = 228571L };

            transform2D7.Append(offset9);
            transform2D7.Append(extents9);

            A.PresetGeometry presetGeometry7 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList7 = new A.AdjustValueList();

            presetGeometry7.Append(adjustValueList7);
            A.NoFill noFill145 = new A.NoFill();
            A.Outline outline19 = new A.Outline();

            shapeProperties9.Append(transform2D7);
            shapeProperties9.Append(presetGeometry7);
            shapeProperties9.Append(noFill145);
            shapeProperties9.Append(outline19);

            TextBody textBody28 = new TextBody();
            A.BodyProperties bodyProperties28 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle28 = new A.ListStyle();

            A.Paragraph paragraph28 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties22 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing20 = new A.LineSpacing();
            A.SpacingPercent spacingPercent20 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing20.Append(spacingPercent20);

            paragraphProperties22.Append(lineSpacing20);

            A.Run run16 = new A.Run();

            A.RunProperties runProperties16 = new A.RunProperties() { Language = "en-US", FontSize = 1080 };

            A.SolidFill solidFill56 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex16 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha5 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex16.Append(alpha5);

            solidFill56.Append(rgbColorModelHex16);
            A.LatinFont latinFont41 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont31 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont41 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties16.Append(solidFill56);
            runProperties16.Append(latinFont41);
            runProperties16.Append(eastAsianFont31);
            runProperties16.Append(complexScriptFont41);
            A.Text text16 = new A.Text();
            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count() > 2)
                text16.Text = goalsArrayForSlide.ElementAt(2).goalOwner ?? "";
            else
                text16.Text = "NO DATA";
            // ********************************* Custom Code Changes End *********************************

            run16.Append(runProperties16);
            run16.Append(text16);
            A.EndParagraphRunProperties endParagraphRunProperties23 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080 };

            paragraph28.Append(paragraphProperties22);
            paragraph28.Append(run16);
            paragraph28.Append(endParagraphRunProperties23);

            textBody28.Append(bodyProperties28);
            textBody28.Append(listStyle28);
            textBody28.Append(paragraph28);

            shape7.Append(nonVisualShapeProperties7);
            shape7.Append(shapeProperties9);
            shape7.Append(textBody28);

            Shape shape8 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties8 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties12 = new NonVisualDrawingProperties() { Id = (UInt32Value)12U, Name = "Text 46" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList11 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension11 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement25 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{7010F11F-5CFE-C13E-372F-ED251C1B3BCE}\" />");

            // nonVisualDrawingPropertiesExtension11.Append(openXmlUnknownElement25);

            nonVisualDrawingPropertiesExtensionList11.Append(nonVisualDrawingPropertiesExtension11);

            nonVisualDrawingProperties12.Append(nonVisualDrawingPropertiesExtensionList11);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties8 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties12 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties8.Append(nonVisualDrawingProperties12);
            nonVisualShapeProperties8.Append(nonVisualShapeDrawingProperties8);
            nonVisualShapeProperties8.Append(applicationNonVisualDrawingProperties12);

            ShapeProperties shapeProperties10 = new ShapeProperties();

            A.Transform2D transform2D8 = new A.Transform2D();
            A.Offset offset10 = new A.Offset() { X = 8353878L, Y = 4031547L };
            A.Extents extents10 = new A.Extents() { Cx = 1280895L, Cy = 228571L };

            transform2D8.Append(offset10);
            transform2D8.Append(extents10);

            A.PresetGeometry presetGeometry8 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList8 = new A.AdjustValueList();

            presetGeometry8.Append(adjustValueList8);
            A.NoFill noFill146 = new A.NoFill();
            A.Outline outline20 = new A.Outline();

            shapeProperties10.Append(transform2D8);
            shapeProperties10.Append(presetGeometry8);
            shapeProperties10.Append(noFill146);
            shapeProperties10.Append(outline20);

            TextBody textBody29 = new TextBody();
            A.BodyProperties bodyProperties29 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle29 = new A.ListStyle();

            A.Paragraph paragraph29 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties23 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing21 = new A.LineSpacing();
            A.SpacingPercent spacingPercent21 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing21.Append(spacingPercent21);

            paragraphProperties23.Append(lineSpacing21);

            A.Run run17 = new A.Run();

            A.RunProperties runProperties17 = new A.RunProperties() { Language = "en-US", FontSize = 1080 };

            A.SolidFill solidFill57 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex17 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha6 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex17.Append(alpha6);

            solidFill57.Append(rgbColorModelHex17);
            A.LatinFont latinFont42 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont32 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont42 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties17.Append(solidFill57);
            runProperties17.Append(latinFont42);
            runProperties17.Append(eastAsianFont32);
            runProperties17.Append(complexScriptFont42);
            A.Text text17 = new A.Text();
            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count() > 3)
                text17.Text = goalsArrayForSlide.ElementAt(3).goalOwner ?? "";
            else
                text17.Text = "NO DATA";
            // ********************************* Custom Code Changes End *********************************

            run17.Append(runProperties17);
            run17.Append(text17);
            A.EndParagraphRunProperties endParagraphRunProperties24 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080 };

            paragraph29.Append(paragraphProperties23);
            paragraph29.Append(run17);
            paragraph29.Append(endParagraphRunProperties24);

            textBody29.Append(bodyProperties29);
            textBody29.Append(listStyle29);
            textBody29.Append(paragraph29);

            shape8.Append(nonVisualShapeProperties8);
            shape8.Append(shapeProperties10);
            shape8.Append(textBody29);

            Shape shape9 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties9 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties13 = new NonVisualDrawingProperties() { Id = (UInt32Value)13U, Name = "Text 46" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList12 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension12 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement26 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{AC854CCD-B672-BC6B-9990-E57A66BE4135}\" />");

            // nonVisualDrawingPropertiesExtension12.Append(openXmlUnknownElement26);

            nonVisualDrawingPropertiesExtensionList12.Append(nonVisualDrawingPropertiesExtension12);

            nonVisualDrawingProperties13.Append(nonVisualDrawingPropertiesExtensionList12);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties9 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties13 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties9.Append(nonVisualDrawingProperties13);
            nonVisualShapeProperties9.Append(nonVisualShapeDrawingProperties9);
            nonVisualShapeProperties9.Append(applicationNonVisualDrawingProperties13);

            ShapeProperties shapeProperties11 = new ShapeProperties();

            A.Transform2D transform2D9 = new A.Transform2D();
            A.Offset offset11 = new A.Offset() { X = 8355349L, Y = 5654066L };
            A.Extents extents11 = new A.Extents() { Cx = 1280895L, Cy = 228571L };

            transform2D9.Append(offset11);
            transform2D9.Append(extents11);

            A.PresetGeometry presetGeometry9 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList9 = new A.AdjustValueList();

            presetGeometry9.Append(adjustValueList9);
            A.NoFill noFill147 = new A.NoFill();
            A.Outline outline21 = new A.Outline();

            shapeProperties11.Append(transform2D9);
            shapeProperties11.Append(presetGeometry9);
            shapeProperties11.Append(noFill147);
            shapeProperties11.Append(outline21);

            TextBody textBody30 = new TextBody();
            A.BodyProperties bodyProperties30 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle30 = new A.ListStyle();

            A.Paragraph paragraph30 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties24 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing22 = new A.LineSpacing();
            A.SpacingPercent spacingPercent22 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing22.Append(spacingPercent22);

            paragraphProperties24.Append(lineSpacing22);

            A.Run run18 = new A.Run();

            A.RunProperties runProperties18 = new A.RunProperties() { Language = "en-US", FontSize = 1080 };

            A.SolidFill solidFill58 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex18 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha7 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex18.Append(alpha7);

            solidFill58.Append(rgbColorModelHex18);
            A.LatinFont latinFont43 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont33 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont43 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties18.Append(solidFill58);
            runProperties18.Append(latinFont43);
            runProperties18.Append(eastAsianFont33);
            runProperties18.Append(complexScriptFont43);
            A.Text text18 = new A.Text();
            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count() > 5)
                text18.Text = goalsArrayForSlide.ElementAt(5).goalOwner ?? "";
            else
                text18.Text = "NO DATA";
            // ********************************* Custom Code Changes End *********************************

            run18.Append(runProperties18);
            run18.Append(text18);
            A.EndParagraphRunProperties endParagraphRunProperties25 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080 };

            paragraph30.Append(paragraphProperties24);
            paragraph30.Append(run18);
            paragraph30.Append(endParagraphRunProperties25);

            textBody30.Append(bodyProperties30);
            textBody30.Append(listStyle30);
            textBody30.Append(paragraph30);

            shape9.Append(nonVisualShapeProperties9);
            shape9.Append(shapeProperties11);
            shape9.Append(textBody30);

            GroupShape groupShape1 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties2 = new NonVisualGroupShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties14 = new NonVisualDrawingProperties() { Id = (UInt32Value)14U, Name = "Group 13" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList13 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension13 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement27 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{2EFC0D7C-20CA-ECB5-F7FD-2F8EB2546CDC}\" />");

            // nonVisualDrawingPropertiesExtension13.Append(openXmlUnknownElement27);

            nonVisualDrawingPropertiesExtensionList13.Append(nonVisualDrawingPropertiesExtension13);

            nonVisualDrawingProperties14.Append(nonVisualDrawingPropertiesExtensionList13);
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties2 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties14 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties2.Append(nonVisualDrawingProperties14);
            nonVisualGroupShapeProperties2.Append(nonVisualGroupShapeDrawingProperties2);
            nonVisualGroupShapeProperties2.Append(applicationNonVisualDrawingProperties14);

            GroupShapeProperties groupShapeProperties2 = new GroupShapeProperties();

            A.TransformGroup transformGroup2 = new A.TransformGroup();
            A.Offset offset12 = new A.Offset() { X = 10344259L, Y = 1783073L };
            A.Extents extents12 = new A.Extents() { Cx = 731939L, Cy = 228571L };
            A.ChildOffset childOffset2 = new A.ChildOffset() { X = 10453015L, Y = 2651915L };
            A.ChildExtents childExtents2 = new A.ChildExtents() { Cx = 731939L, Cy = 228571L };

            transformGroup2.Append(offset12);
            transformGroup2.Append(extents12);
            transformGroup2.Append(childOffset2);
            transformGroup2.Append(childExtents2);

            groupShapeProperties2.Append(transformGroup2);

            Picture picture3 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties3 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties15 = new NonVisualDrawingProperties() { Id = (UInt32Value)15U, Name = "Image 24", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList14 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension14 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement28 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{FBD43D53-88D2-1229-D616-0E4A0A47F56B}\" />");

            // nonVisualDrawingPropertiesExtension14.Append(openXmlUnknownElement28);

            nonVisualDrawingPropertiesExtensionList14.Append(nonVisualDrawingPropertiesExtension14);

            nonVisualDrawingProperties15.Append(nonVisualDrawingPropertiesExtensionList14);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties3 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties3.Append(pictureLocks1);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties15 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties3.Append(nonVisualDrawingProperties15);
            nonVisualPictureProperties3.Append(nonVisualPictureDrawingProperties3);
            nonVisualPictureProperties3.Append(applicationNonVisualDrawingProperties15);

            BlipFill blipFill3 = new BlipFill();

            A.Blip blip7 = new A.Blip() { Embed = "rId6" };

            A.BlipExtensionList blipExtensionList5 = new A.BlipExtensionList();

            A.BlipExtension blipExtension5 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement29 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId7\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension5.Append(openXmlUnknownElement29);

            blipExtensionList5.Append(blipExtension5);

            blip7.Append(blipExtensionList5);

            A.Stretch stretch3 = new A.Stretch();
            A.FillRectangle fillRectangle3 = new A.FillRectangle();

            stretch3.Append(fillRectangle3);

            blipFill3.Append(blip7);
            blipFill3.Append(stretch3);

            ShapeProperties shapeProperties12 = new ShapeProperties();

            A.Transform2D transform2D10 = new A.Transform2D();
            A.Offset offset13 = new A.Offset() { X = 10453015L, Y = 2709059L };
            A.Extents extents13 = new A.Extents() { Cx = 137250L, Cy = 137157L };

            transform2D10.Append(offset13);
            transform2D10.Append(extents13);

            A.PresetGeometry presetGeometry10 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList10 = new A.AdjustValueList();

            presetGeometry10.Append(adjustValueList10);

            shapeProperties12.Append(transform2D10);
            shapeProperties12.Append(presetGeometry10);

            picture3.Append(nonVisualPictureProperties3);
            picture3.Append(blipFill3);
            picture3.Append(shapeProperties12);

            Shape shape10 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties10 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties16 = new NonVisualDrawingProperties() { Id = (UInt32Value)16U, Name = "Text 48" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList15 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension15 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement30 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{5DBDD316-B408-BF96-45AA-39E230D26B47}\" />");

            // nonVisualDrawingPropertiesExtension15.Append(openXmlUnknownElement30);

            nonVisualDrawingPropertiesExtensionList15.Append(nonVisualDrawingPropertiesExtension15);

            nonVisualDrawingProperties16.Append(nonVisualDrawingPropertiesExtensionList15);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties10 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties16 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties10.Append(nonVisualDrawingProperties16);
            nonVisualShapeProperties10.Append(nonVisualShapeDrawingProperties10);
            nonVisualShapeProperties10.Append(applicationNonVisualDrawingProperties16);

            ShapeProperties shapeProperties13 = new ShapeProperties();

            A.Transform2D transform2D11 = new A.Transform2D();
            A.Offset offset14 = new A.Offset() { X = 10658873L, Y = 2651915L };
            A.Extents extents14 = new A.Extents() { Cx = 526081L, Cy = 228571L };

            transform2D11.Append(offset14);
            transform2D11.Append(extents14);

            A.PresetGeometry presetGeometry11 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList11 = new A.AdjustValueList();

            presetGeometry11.Append(adjustValueList11);
            A.NoFill noFill148 = new A.NoFill();
            A.Outline outline22 = new A.Outline();

            shapeProperties13.Append(transform2D11);
            shapeProperties13.Append(presetGeometry11);
            shapeProperties13.Append(noFill148);
            shapeProperties13.Append(outline22);

            TextBody textBody31 = new TextBody();
            A.BodyProperties bodyProperties31 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle31 = new A.ListStyle();

            A.Paragraph paragraph31 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties25 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing23 = new A.LineSpacing();
            A.SpacingPercent spacingPercent23 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing23.Append(spacingPercent23);

            paragraphProperties25.Append(lineSpacing23);

            A.Run run19 = new A.Run();

            A.RunProperties runProperties19 = new A.RunProperties() { Language = "en-US", FontSize = 1050, Bold = true };

            A.SolidFill solidFill59 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex19 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha8 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex19.Append(alpha8);

            solidFill59.Append(rgbColorModelHex19);
            A.LatinFont latinFont44 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont34 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont44 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            runProperties19.Append(solidFill59);
            runProperties19.Append(latinFont44);
            runProperties19.Append(eastAsianFont34);
            runProperties19.Append(complexScriptFont44);
            A.Text text19 = new A.Text();

            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count() > 0)
                text19.Text = goalsArrayForSlide.ElementAt(0).target ?? "";
            else
                text19.Text = "NO DATA";
            // ********************************* Custom Code Changes End *********************************

            run19.Append(runProperties19);
            run19.Append(text19);

            A.EndParagraphRunProperties endParagraphRunProperties26 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1050, Bold = true };
            A.LatinFont latinFont45 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont45 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties26.Append(latinFont45);
            endParagraphRunProperties26.Append(complexScriptFont45);

            paragraph31.Append(paragraphProperties25);
            paragraph31.Append(run19);
            paragraph31.Append(endParagraphRunProperties26);

            textBody31.Append(bodyProperties31);
            textBody31.Append(listStyle31);
            textBody31.Append(paragraph31);

            shape10.Append(nonVisualShapeProperties10);
            shape10.Append(shapeProperties13);
            shape10.Append(textBody31);

            groupShape1.Append(nonVisualGroupShapeProperties2);
            groupShape1.Append(groupShapeProperties2);
            groupShape1.Append(picture3);
            groupShape1.Append(shape10);

            GroupShape groupShape2 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties3 = new NonVisualGroupShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties17 = new NonVisualDrawingProperties() { Id = (UInt32Value)17U, Name = "Group 16" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList16 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension16 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement31 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{EFE55359-869F-8B61-793A-543539D9D8C6}\" />");

            // nonVisualDrawingPropertiesExtension16.Append(openXmlUnknownElement31);

            nonVisualDrawingPropertiesExtensionList16.Append(nonVisualDrawingPropertiesExtension16);

            nonVisualDrawingProperties17.Append(nonVisualDrawingPropertiesExtensionList16);
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties3 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties17 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties3.Append(nonVisualDrawingProperties17);
            nonVisualGroupShapeProperties3.Append(nonVisualGroupShapeDrawingProperties3);
            nonVisualGroupShapeProperties3.Append(applicationNonVisualDrawingProperties17);

            GroupShapeProperties groupShapeProperties3 = new GroupShapeProperties();

            A.TransformGroup transformGroup3 = new A.TransformGroup();
            A.Offset offset15 = new A.Offset() { X = 10344259L, Y = 2537221L };
            A.Extents extents15 = new A.Extents() { Cx = 731939L, Cy = 228571L };
            A.ChildOffset childOffset3 = new A.ChildOffset() { X = 10453015L, Y = 2651915L };
            A.ChildExtents childExtents3 = new A.ChildExtents() { Cx = 731939L, Cy = 228571L };

            transformGroup3.Append(offset15);
            transformGroup3.Append(extents15);
            transformGroup3.Append(childOffset3);
            transformGroup3.Append(childExtents3);

            groupShapeProperties3.Append(transformGroup3);

            Picture picture4 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties4 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties18 = new NonVisualDrawingProperties() { Id = (UInt32Value)18U, Name = "Image 24", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList17 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension17 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement32 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{8A240DBE-E256-336F-46B0-3E9E272B7AEE}\" />");

            // nonVisualDrawingPropertiesExtension17.Append(openXmlUnknownElement32);

            nonVisualDrawingPropertiesExtensionList17.Append(nonVisualDrawingPropertiesExtension17);

            nonVisualDrawingProperties18.Append(nonVisualDrawingPropertiesExtensionList17);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties4 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks2 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties4.Append(pictureLocks2);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties18 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties4.Append(nonVisualDrawingProperties18);
            nonVisualPictureProperties4.Append(nonVisualPictureDrawingProperties4);
            nonVisualPictureProperties4.Append(applicationNonVisualDrawingProperties18);

            BlipFill blipFill4 = new BlipFill();

            A.Blip blip8 = new A.Blip() { Embed = "rId6" };

            A.BlipExtensionList blipExtensionList6 = new A.BlipExtensionList();

            A.BlipExtension blipExtension6 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement33 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId7\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension6.Append(openXmlUnknownElement33);

            blipExtensionList6.Append(blipExtension6);

            blip8.Append(blipExtensionList6);

            A.Stretch stretch4 = new A.Stretch();
            A.FillRectangle fillRectangle4 = new A.FillRectangle();

            stretch4.Append(fillRectangle4);

            blipFill4.Append(blip8);
            blipFill4.Append(stretch4);

            ShapeProperties shapeProperties14 = new ShapeProperties();

            A.Transform2D transform2D12 = new A.Transform2D();
            A.Offset offset16 = new A.Offset() { X = 10453015L, Y = 2709059L };
            A.Extents extents16 = new A.Extents() { Cx = 137250L, Cy = 137157L };

            transform2D12.Append(offset16);
            transform2D12.Append(extents16);

            A.PresetGeometry presetGeometry12 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList12 = new A.AdjustValueList();

            presetGeometry12.Append(adjustValueList12);

            shapeProperties14.Append(transform2D12);
            shapeProperties14.Append(presetGeometry12);

            picture4.Append(nonVisualPictureProperties4);
            picture4.Append(blipFill4);
            picture4.Append(shapeProperties14);

            Shape shape11 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties11 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties19 = new NonVisualDrawingProperties() { Id = (UInt32Value)19U, Name = "Text 48" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList18 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension18 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement34 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{389F9B76-9DC9-2111-B828-CFE0586CEE47}\" />");

            // nonVisualDrawingPropertiesExtension18.Append(openXmlUnknownElement34);

            nonVisualDrawingPropertiesExtensionList18.Append(nonVisualDrawingPropertiesExtension18);

            nonVisualDrawingProperties19.Append(nonVisualDrawingPropertiesExtensionList18);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties11 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties19 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties11.Append(nonVisualDrawingProperties19);
            nonVisualShapeProperties11.Append(nonVisualShapeDrawingProperties11);
            nonVisualShapeProperties11.Append(applicationNonVisualDrawingProperties19);

            ShapeProperties shapeProperties15 = new ShapeProperties();

            A.Transform2D transform2D13 = new A.Transform2D();
            A.Offset offset17 = new A.Offset() { X = 10658873L, Y = 2651915L };
            A.Extents extents17 = new A.Extents() { Cx = 526081L, Cy = 228571L };

            transform2D13.Append(offset17);
            transform2D13.Append(extents17);

            A.PresetGeometry presetGeometry13 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList13 = new A.AdjustValueList();

            presetGeometry13.Append(adjustValueList13);
            A.NoFill noFill149 = new A.NoFill();
            A.Outline outline23 = new A.Outline();

            shapeProperties15.Append(transform2D13);
            shapeProperties15.Append(presetGeometry13);
            shapeProperties15.Append(noFill149);
            shapeProperties15.Append(outline23);

            TextBody textBody32 = new TextBody();
            A.BodyProperties bodyProperties32 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle32 = new A.ListStyle();

            A.Paragraph paragraph32 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties26 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing24 = new A.LineSpacing();
            A.SpacingPercent spacingPercent24 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing24.Append(spacingPercent24);

            paragraphProperties26.Append(lineSpacing24);

            A.Run run20 = new A.Run();

            A.RunProperties runProperties20 = new A.RunProperties() { Language = "en-US", FontSize = 1080, Bold = true };

            A.SolidFill solidFill60 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex20 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha9 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex20.Append(alpha9);

            solidFill60.Append(rgbColorModelHex20);
            A.LatinFont latinFont46 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont35 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont46 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            runProperties20.Append(solidFill60);
            runProperties20.Append(latinFont46);
            runProperties20.Append(eastAsianFont35);
            runProperties20.Append(complexScriptFont46);
            A.Text text20 = new A.Text();
            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count() > 1)
                text20.Text = goalsArrayForSlide.ElementAt(1).target ?? "";
            else
                text20.Text = "NO DATA";
            // ********************************* Custom Code Changes End *********************************

            run20.Append(runProperties20);
            run20.Append(text20);

            A.EndParagraphRunProperties endParagraphRunProperties27 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080, Bold = true };
            A.LatinFont latinFont47 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont47 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties27.Append(latinFont47);
            endParagraphRunProperties27.Append(complexScriptFont47);

            paragraph32.Append(paragraphProperties26);
            paragraph32.Append(run20);
            paragraph32.Append(endParagraphRunProperties27);

            textBody32.Append(bodyProperties32);
            textBody32.Append(listStyle32);
            textBody32.Append(paragraph32);

            shape11.Append(nonVisualShapeProperties11);
            shape11.Append(shapeProperties15);
            shape11.Append(textBody32);

            groupShape2.Append(nonVisualGroupShapeProperties3);
            groupShape2.Append(groupShapeProperties3);
            groupShape2.Append(picture4);
            groupShape2.Append(shape11);

            GroupShape groupShape3 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties4 = new NonVisualGroupShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties20 = new NonVisualDrawingProperties() { Id = (UInt32Value)20U, Name = "Group 19" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList19 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension19 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement35 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{D8987891-0D8A-39FC-775F-9EC24BEBA006}\" />");

            // nonVisualDrawingPropertiesExtension19.Append(openXmlUnknownElement35);

            nonVisualDrawingPropertiesExtensionList19.Append(nonVisualDrawingPropertiesExtension19);

            nonVisualDrawingProperties20.Append(nonVisualDrawingPropertiesExtensionList19);
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties4 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties20 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties4.Append(nonVisualDrawingProperties20);
            nonVisualGroupShapeProperties4.Append(nonVisualGroupShapeDrawingProperties4);
            nonVisualGroupShapeProperties4.Append(applicationNonVisualDrawingProperties20);

            GroupShapeProperties groupShapeProperties4 = new GroupShapeProperties();

            A.TransformGroup transformGroup4 = new A.TransformGroup();
            A.Offset offset18 = new A.Offset() { X = 10344259L, Y = 3283601L };
            A.Extents extents18 = new A.Extents() { Cx = 731939L, Cy = 228571L };
            A.ChildOffset childOffset4 = new A.ChildOffset() { X = 10453015L, Y = 2651915L };
            A.ChildExtents childExtents4 = new A.ChildExtents() { Cx = 731939L, Cy = 228571L };

            transformGroup4.Append(offset18);
            transformGroup4.Append(extents18);
            transformGroup4.Append(childOffset4);
            transformGroup4.Append(childExtents4);

            groupShapeProperties4.Append(transformGroup4);

            Picture picture5 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties5 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties21 = new NonVisualDrawingProperties() { Id = (UInt32Value)21U, Name = "Image 24", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList20 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension20 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement36 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{FC52E511-F997-4B84-3F36-F431F761CBAB}\" />");

            // nonVisualDrawingPropertiesExtension20.Append(openXmlUnknownElement36);

            nonVisualDrawingPropertiesExtensionList20.Append(nonVisualDrawingPropertiesExtension20);

            nonVisualDrawingProperties21.Append(nonVisualDrawingPropertiesExtensionList20);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties5 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks3 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties5.Append(pictureLocks3);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties21 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties5.Append(nonVisualDrawingProperties21);
            nonVisualPictureProperties5.Append(nonVisualPictureDrawingProperties5);
            nonVisualPictureProperties5.Append(applicationNonVisualDrawingProperties21);

            BlipFill blipFill5 = new BlipFill();

            A.Blip blip9 = new A.Blip() { Embed = "rId6" };

            A.BlipExtensionList blipExtensionList7 = new A.BlipExtensionList();

            A.BlipExtension blipExtension7 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement37 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId7\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension7.Append(openXmlUnknownElement37);

            blipExtensionList7.Append(blipExtension7);

            blip9.Append(blipExtensionList7);

            A.Stretch stretch5 = new A.Stretch();
            A.FillRectangle fillRectangle5 = new A.FillRectangle();

            stretch5.Append(fillRectangle5);

            blipFill5.Append(blip9);
            blipFill5.Append(stretch5);

            ShapeProperties shapeProperties16 = new ShapeProperties();

            A.Transform2D transform2D14 = new A.Transform2D();
            A.Offset offset19 = new A.Offset() { X = 10453015L, Y = 2709059L };
            A.Extents extents19 = new A.Extents() { Cx = 137250L, Cy = 137157L };

            transform2D14.Append(offset19);
            transform2D14.Append(extents19);

            A.PresetGeometry presetGeometry14 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList14 = new A.AdjustValueList();

            presetGeometry14.Append(adjustValueList14);

            shapeProperties16.Append(transform2D14);
            shapeProperties16.Append(presetGeometry14);

            picture5.Append(nonVisualPictureProperties5);
            picture5.Append(blipFill5);
            picture5.Append(shapeProperties16);

            Shape shape12 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties12 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties22 = new NonVisualDrawingProperties() { Id = (UInt32Value)22U, Name = "Text 48" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList21 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension21 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement38 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{E74AE8BB-1E5A-E855-A26A-5E5CEF058C2C}\" />");

            // nonVisualDrawingPropertiesExtension21.Append(openXmlUnknownElement38);

            nonVisualDrawingPropertiesExtensionList21.Append(nonVisualDrawingPropertiesExtension21);

            nonVisualDrawingProperties22.Append(nonVisualDrawingPropertiesExtensionList21);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties12 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties22 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties12.Append(nonVisualDrawingProperties22);
            nonVisualShapeProperties12.Append(nonVisualShapeDrawingProperties12);
            nonVisualShapeProperties12.Append(applicationNonVisualDrawingProperties22);

            ShapeProperties shapeProperties17 = new ShapeProperties();

            A.Transform2D transform2D15 = new A.Transform2D();
            A.Offset offset20 = new A.Offset() { X = 10658873L, Y = 2651915L };
            A.Extents extents20 = new A.Extents() { Cx = 526081L, Cy = 228571L };

            transform2D15.Append(offset20);
            transform2D15.Append(extents20);

            A.PresetGeometry presetGeometry15 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList15 = new A.AdjustValueList();

            presetGeometry15.Append(adjustValueList15);
            A.NoFill noFill150 = new A.NoFill();
            A.Outline outline24 = new A.Outline();

            shapeProperties17.Append(transform2D15);
            shapeProperties17.Append(presetGeometry15);
            shapeProperties17.Append(noFill150);
            shapeProperties17.Append(outline24);

            TextBody textBody33 = new TextBody();
            A.BodyProperties bodyProperties33 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle33 = new A.ListStyle();

            A.Paragraph paragraph33 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties27 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing25 = new A.LineSpacing();
            A.SpacingPercent spacingPercent25 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing25.Append(spacingPercent25);

            paragraphProperties27.Append(lineSpacing25);

            A.Run run21 = new A.Run();

            A.RunProperties runProperties21 = new A.RunProperties() { Language = "en-US", FontSize = 1080, Bold = true };

            A.SolidFill solidFill61 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex21 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha10 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex21.Append(alpha10);

            solidFill61.Append(rgbColorModelHex21);
            A.LatinFont latinFont48 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont36 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont48 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            runProperties21.Append(solidFill61);
            runProperties21.Append(latinFont48);
            runProperties21.Append(eastAsianFont36);
            runProperties21.Append(complexScriptFont48);
            A.Text text21 = new A.Text();
            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count() > 2)
                text21.Text = goalsArrayForSlide.ElementAt(2).target ?? "";
            else
                text21.Text = "NO DATA";
            // ********************************* Custom Code Changes End *********************************

            run21.Append(runProperties21);
            run21.Append(text21);

            A.EndParagraphRunProperties endParagraphRunProperties28 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080, Bold = true };
            A.LatinFont latinFont49 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont49 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties28.Append(latinFont49);
            endParagraphRunProperties28.Append(complexScriptFont49);

            paragraph33.Append(paragraphProperties27);
            paragraph33.Append(run21);
            paragraph33.Append(endParagraphRunProperties28);

            textBody33.Append(bodyProperties33);
            textBody33.Append(listStyle33);
            textBody33.Append(paragraph33);

            shape12.Append(nonVisualShapeProperties12);
            shape12.Append(shapeProperties17);
            shape12.Append(textBody33);

            groupShape3.Append(nonVisualGroupShapeProperties4);
            groupShape3.Append(groupShapeProperties4);
            groupShape3.Append(picture5);
            groupShape3.Append(shape12);

            GroupShape groupShape4 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties5 = new NonVisualGroupShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties23 = new NonVisualDrawingProperties() { Id = (UInt32Value)23U, Name = "Group 22" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList22 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension22 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement39 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{94DA84A7-3D6B-D8F0-9374-4D690A6656AD}\" />");

            // nonVisualDrawingPropertiesExtension22.Append(openXmlUnknownElement39);

            nonVisualDrawingPropertiesExtensionList22.Append(nonVisualDrawingPropertiesExtension22);

            nonVisualDrawingProperties23.Append(nonVisualDrawingPropertiesExtensionList22);
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties5 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties23 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties5.Append(nonVisualDrawingProperties23);
            nonVisualGroupShapeProperties5.Append(nonVisualGroupShapeDrawingProperties5);
            nonVisualGroupShapeProperties5.Append(applicationNonVisualDrawingProperties23);

            GroupShapeProperties groupShapeProperties5 = new GroupShapeProperties();

            A.TransformGroup transformGroup5 = new A.TransformGroup();
            A.Offset offset21 = new A.Offset() { X = 10344259L, Y = 4009400L };
            A.Extents extents21 = new A.Extents() { Cx = 731939L, Cy = 228571L };
            A.ChildOffset childOffset5 = new A.ChildOffset() { X = 10453015L, Y = 2651915L };
            A.ChildExtents childExtents5 = new A.ChildExtents() { Cx = 731939L, Cy = 228571L };

            transformGroup5.Append(offset21);
            transformGroup5.Append(extents21);
            transformGroup5.Append(childOffset5);
            transformGroup5.Append(childExtents5);

            groupShapeProperties5.Append(transformGroup5);

            Picture picture6 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties6 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties24 = new NonVisualDrawingProperties() { Id = (UInt32Value)24U, Name = "Image 24", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList23 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension23 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement40 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{30EC9A71-FEFA-1EA2-6705-598B5BEEAE4E}\" />");

            // nonVisualDrawingPropertiesExtension23.Append(openXmlUnknownElement40);

            nonVisualDrawingPropertiesExtensionList23.Append(nonVisualDrawingPropertiesExtension23);

            nonVisualDrawingProperties24.Append(nonVisualDrawingPropertiesExtensionList23);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties6 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks4 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties6.Append(pictureLocks4);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties24 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties6.Append(nonVisualDrawingProperties24);
            nonVisualPictureProperties6.Append(nonVisualPictureDrawingProperties6);
            nonVisualPictureProperties6.Append(applicationNonVisualDrawingProperties24);

            BlipFill blipFill6 = new BlipFill();

            A.Blip blip10 = new A.Blip() { Embed = "rId6" };

            A.BlipExtensionList blipExtensionList8 = new A.BlipExtensionList();

            A.BlipExtension blipExtension8 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement41 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId7\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension8.Append(openXmlUnknownElement41);

            blipExtensionList8.Append(blipExtension8);

            blip10.Append(blipExtensionList8);

            A.Stretch stretch6 = new A.Stretch();
            A.FillRectangle fillRectangle6 = new A.FillRectangle();

            stretch6.Append(fillRectangle6);

            blipFill6.Append(blip10);
            blipFill6.Append(stretch6);

            ShapeProperties shapeProperties18 = new ShapeProperties();

            A.Transform2D transform2D16 = new A.Transform2D();
            A.Offset offset22 = new A.Offset() { X = 10453015L, Y = 2709059L };
            A.Extents extents22 = new A.Extents() { Cx = 137250L, Cy = 137157L };

            transform2D16.Append(offset22);
            transform2D16.Append(extents22);

            A.PresetGeometry presetGeometry16 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList16 = new A.AdjustValueList();

            presetGeometry16.Append(adjustValueList16);

            shapeProperties18.Append(transform2D16);
            shapeProperties18.Append(presetGeometry16);

            picture6.Append(nonVisualPictureProperties6);
            picture6.Append(blipFill6);
            picture6.Append(shapeProperties18);

            Shape shape13 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties13 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties25 = new NonVisualDrawingProperties() { Id = (UInt32Value)25U, Name = "Text 48" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList24 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension24 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement42 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{E27F3243-7005-457C-3438-5515CF5049E5}\" />");

            // nonVisualDrawingPropertiesExtension24.Append(openXmlUnknownElement42);

            nonVisualDrawingPropertiesExtensionList24.Append(nonVisualDrawingPropertiesExtension24);

            nonVisualDrawingProperties25.Append(nonVisualDrawingPropertiesExtensionList24);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties13 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties25 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties13.Append(nonVisualDrawingProperties25);
            nonVisualShapeProperties13.Append(nonVisualShapeDrawingProperties13);
            nonVisualShapeProperties13.Append(applicationNonVisualDrawingProperties25);

            ShapeProperties shapeProperties19 = new ShapeProperties();

            A.Transform2D transform2D17 = new A.Transform2D();
            A.Offset offset23 = new A.Offset() { X = 10658873L, Y = 2651915L };
            A.Extents extents23 = new A.Extents() { Cx = 526081L, Cy = 228571L };

            transform2D17.Append(offset23);
            transform2D17.Append(extents23);

            A.PresetGeometry presetGeometry17 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList17 = new A.AdjustValueList();

            presetGeometry17.Append(adjustValueList17);
            A.NoFill noFill151 = new A.NoFill();
            A.Outline outline25 = new A.Outline();

            shapeProperties19.Append(transform2D17);
            shapeProperties19.Append(presetGeometry17);
            shapeProperties19.Append(noFill151);
            shapeProperties19.Append(outline25);

            TextBody textBody34 = new TextBody();
            A.BodyProperties bodyProperties34 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle34 = new A.ListStyle();

            A.Paragraph paragraph34 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties28 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing26 = new A.LineSpacing();
            A.SpacingPercent spacingPercent26 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing26.Append(spacingPercent26);

            paragraphProperties28.Append(lineSpacing26);

            A.Run run22 = new A.Run();

            A.RunProperties runProperties22 = new A.RunProperties() { Language = "en-US", FontSize = 1080, Bold = true };

            A.SolidFill solidFill62 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex22 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha11 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex22.Append(alpha11);

            solidFill62.Append(rgbColorModelHex22);
            A.LatinFont latinFont50 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont37 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont50 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            runProperties22.Append(solidFill62);
            runProperties22.Append(latinFont50);
            runProperties22.Append(eastAsianFont37);
            runProperties22.Append(complexScriptFont50);
            A.Text text22 = new A.Text();
            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count() > 3)
                text22.Text = goalsArrayForSlide.ElementAt(3).target ?? "";
            else
                text22.Text = "NO DATA";
            // ********************************* Custom Code Changes End *********************************

            run22.Append(runProperties22);
            run22.Append(text22);

            A.EndParagraphRunProperties endParagraphRunProperties29 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080, Bold = true };
            A.LatinFont latinFont51 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont51 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties29.Append(latinFont51);
            endParagraphRunProperties29.Append(complexScriptFont51);

            paragraph34.Append(paragraphProperties28);
            paragraph34.Append(run22);
            paragraph34.Append(endParagraphRunProperties29);

            textBody34.Append(bodyProperties34);
            textBody34.Append(listStyle34);
            textBody34.Append(paragraph34);

            shape13.Append(nonVisualShapeProperties13);
            shape13.Append(shapeProperties19);
            shape13.Append(textBody34);

            groupShape4.Append(nonVisualGroupShapeProperties5);
            groupShape4.Append(groupShapeProperties5);
            groupShape4.Append(picture6);
            groupShape4.Append(shape13);

            GroupShape groupShape5 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties6 = new NonVisualGroupShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties26 = new NonVisualDrawingProperties() { Id = (UInt32Value)26U, Name = "Group 25" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList25 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension25 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement43 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{14CAB9C8-2F17-B974-F712-F855E3BE886C}\" />");

            // nonVisualDrawingPropertiesExtension25.Append(openXmlUnknownElement43);

            nonVisualDrawingPropertiesExtensionList25.Append(nonVisualDrawingPropertiesExtension25);

            nonVisualDrawingProperties26.Append(nonVisualDrawingPropertiesExtensionList25);
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties6 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties26 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties6.Append(nonVisualDrawingProperties26);
            nonVisualGroupShapeProperties6.Append(nonVisualGroupShapeDrawingProperties6);
            nonVisualGroupShapeProperties6.Append(applicationNonVisualDrawingProperties26);

            GroupShapeProperties groupShapeProperties6 = new GroupShapeProperties();

            A.TransformGroup transformGroup6 = new A.TransformGroup();
            A.Offset offset24 = new A.Offset() { X = 10344259L, Y = 5642630L };
            A.Extents extents24 = new A.Extents() { Cx = 731939L, Cy = 228571L };
            A.ChildOffset childOffset6 = new A.ChildOffset() { X = 10453015L, Y = 2651915L };
            A.ChildExtents childExtents6 = new A.ChildExtents() { Cx = 731939L, Cy = 228571L };

            transformGroup6.Append(offset24);
            transformGroup6.Append(extents24);
            transformGroup6.Append(childOffset6);
            transformGroup6.Append(childExtents6);

            groupShapeProperties6.Append(transformGroup6);

            Picture picture7 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties7 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties27 = new NonVisualDrawingProperties() { Id = (UInt32Value)27U, Name = "Image 24", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList26 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension26 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement44 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{76886C98-B980-96AD-7C95-7C02A0BA71EE}\" />");

            // nonVisualDrawingPropertiesExtension26.Append(openXmlUnknownElement44);

            nonVisualDrawingPropertiesExtensionList26.Append(nonVisualDrawingPropertiesExtension26);

            nonVisualDrawingProperties27.Append(nonVisualDrawingPropertiesExtensionList26);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties7 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks5 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties7.Append(pictureLocks5);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties27 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties7.Append(nonVisualDrawingProperties27);
            nonVisualPictureProperties7.Append(nonVisualPictureDrawingProperties7);
            nonVisualPictureProperties7.Append(applicationNonVisualDrawingProperties27);

            BlipFill blipFill7 = new BlipFill();

            A.Blip blip11 = new A.Blip() { Embed = "rId6" };

            A.BlipExtensionList blipExtensionList9 = new A.BlipExtensionList();

            A.BlipExtension blipExtension9 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement45 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId7\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension9.Append(openXmlUnknownElement45);

            blipExtensionList9.Append(blipExtension9);

            blip11.Append(blipExtensionList9);

            A.Stretch stretch7 = new A.Stretch();
            A.FillRectangle fillRectangle7 = new A.FillRectangle();

            stretch7.Append(fillRectangle7);

            blipFill7.Append(blip11);
            blipFill7.Append(stretch7);

            ShapeProperties shapeProperties20 = new ShapeProperties();

            A.Transform2D transform2D18 = new A.Transform2D();
            A.Offset offset25 = new A.Offset() { X = 10453015L, Y = 2709059L };
            A.Extents extents25 = new A.Extents() { Cx = 137250L, Cy = 137157L };

            transform2D18.Append(offset25);
            transform2D18.Append(extents25);

            A.PresetGeometry presetGeometry18 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList18 = new A.AdjustValueList();

            presetGeometry18.Append(adjustValueList18);

            shapeProperties20.Append(transform2D18);
            shapeProperties20.Append(presetGeometry18);

            picture7.Append(nonVisualPictureProperties7);
            picture7.Append(blipFill7);
            picture7.Append(shapeProperties20);

            Shape shape14 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties14 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties28 = new NonVisualDrawingProperties() { Id = (UInt32Value)28U, Name = "Text 48" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList27 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension27 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement46 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{EE3F3508-60F5-C86D-3AEE-5EE7F82B2A7B}\" />");

            // nonVisualDrawingPropertiesExtension27.Append(openXmlUnknownElement46);

            nonVisualDrawingPropertiesExtensionList27.Append(nonVisualDrawingPropertiesExtension27);

            nonVisualDrawingProperties28.Append(nonVisualDrawingPropertiesExtensionList27);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties14 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties28 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties14.Append(nonVisualDrawingProperties28);
            nonVisualShapeProperties14.Append(nonVisualShapeDrawingProperties14);
            nonVisualShapeProperties14.Append(applicationNonVisualDrawingProperties28);

            ShapeProperties shapeProperties21 = new ShapeProperties();

            A.Transform2D transform2D19 = new A.Transform2D();
            A.Offset offset26 = new A.Offset() { X = 10658873L, Y = 2651915L };
            A.Extents extents26 = new A.Extents() { Cx = 526081L, Cy = 228571L };

            transform2D19.Append(offset26);
            transform2D19.Append(extents26);

            A.PresetGeometry presetGeometry19 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList19 = new A.AdjustValueList();

            presetGeometry19.Append(adjustValueList19);
            A.NoFill noFill152 = new A.NoFill();
            A.Outline outline26 = new A.Outline();

            shapeProperties21.Append(transform2D19);
            shapeProperties21.Append(presetGeometry19);
            shapeProperties21.Append(noFill152);
            shapeProperties21.Append(outline26);

            TextBody textBody35 = new TextBody();
            A.BodyProperties bodyProperties35 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle35 = new A.ListStyle();

            A.Paragraph paragraph35 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties29 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing27 = new A.LineSpacing();
            A.SpacingPercent spacingPercent27 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing27.Append(spacingPercent27);

            paragraphProperties29.Append(lineSpacing27);

            A.Run run23 = new A.Run();

            A.RunProperties runProperties23 = new A.RunProperties() { Language = "en-US", FontSize = 1080, Bold = true };

            A.SolidFill solidFill63 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex23 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha12 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex23.Append(alpha12);

            solidFill63.Append(rgbColorModelHex23);
            A.LatinFont latinFont52 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont38 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont52 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            runProperties23.Append(solidFill63);
            runProperties23.Append(latinFont52);
            runProperties23.Append(eastAsianFont38);
            runProperties23.Append(complexScriptFont52);
            A.Text text23 = new A.Text();
            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count() > 5)
                text23.Text = goalsArrayForSlide.ElementAt(5).target ?? "";
            else
                text23.Text = "NO DATA";
            // ********************************* Custom Code Changes End *********************************

            run23.Append(runProperties23);
            run23.Append(text23);

            A.EndParagraphRunProperties endParagraphRunProperties30 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080, Bold = true };
            A.LatinFont latinFont53 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont53 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties30.Append(latinFont53);
            endParagraphRunProperties30.Append(complexScriptFont53);

            paragraph35.Append(paragraphProperties29);
            paragraph35.Append(run23);
            paragraph35.Append(endParagraphRunProperties30);

            textBody35.Append(bodyProperties35);
            textBody35.Append(listStyle35);
            textBody35.Append(paragraph35);

            shape14.Append(nonVisualShapeProperties14);
            shape14.Append(shapeProperties21);
            shape14.Append(textBody35);

            groupShape5.Append(nonVisualGroupShapeProperties6);
            groupShape5.Append(groupShapeProperties6);
            groupShape5.Append(picture7);
            groupShape5.Append(shape14);

            GroupShape groupShape6 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties7 = new NonVisualGroupShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties29 = new NonVisualDrawingProperties() { Id = (UInt32Value)29U, Name = "Group 28" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList28 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension28 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement47 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{B7BF25A5-D2AB-2BDF-FB2C-69514102F186}\" />");

            // nonVisualDrawingPropertiesExtension28.Append(openXmlUnknownElement47);

            nonVisualDrawingPropertiesExtensionList28.Append(nonVisualDrawingPropertiesExtension28);

            nonVisualDrawingProperties29.Append(nonVisualDrawingPropertiesExtensionList28);
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties7 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties29 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties7.Append(nonVisualDrawingProperties29);
            nonVisualGroupShapeProperties7.Append(nonVisualGroupShapeDrawingProperties7);
            nonVisualGroupShapeProperties7.Append(applicationNonVisualDrawingProperties29);

            GroupShapeProperties groupShapeProperties7 = new GroupShapeProperties();

            A.TransformGroup transformGroup7 = new A.TransformGroup();
            A.Offset offset27 = new A.Offset() { X = 10339639L, Y = 4811742L };
            A.Extents extents27 = new A.Extents() { Cx = 731939L, Cy = 228571L };
            A.ChildOffset childOffset7 = new A.ChildOffset() { X = 10453015L, Y = 2651915L };
            A.ChildExtents childExtents7 = new A.ChildExtents() { Cx = 731939L, Cy = 228571L };

            transformGroup7.Append(offset27);
            transformGroup7.Append(extents27);
            transformGroup7.Append(childOffset7);
            transformGroup7.Append(childExtents7);

            groupShapeProperties7.Append(transformGroup7);

            Picture picture8 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties8 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties30 = new NonVisualDrawingProperties() { Id = (UInt32Value)30U, Name = "Image 24", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList29 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension29 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement48 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{C6F4F0E2-AAA4-D0A1-0A58-AE68B37DCC5B}\" />");

            // nonVisualDrawingPropertiesExtension29.Append(openXmlUnknownElement48);

            nonVisualDrawingPropertiesExtensionList29.Append(nonVisualDrawingPropertiesExtension29);

            nonVisualDrawingProperties30.Append(nonVisualDrawingPropertiesExtensionList29);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties8 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks6 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties8.Append(pictureLocks6);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties30 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties8.Append(nonVisualDrawingProperties30);
            nonVisualPictureProperties8.Append(nonVisualPictureDrawingProperties8);
            nonVisualPictureProperties8.Append(applicationNonVisualDrawingProperties30);

            BlipFill blipFill8 = new BlipFill();

            A.Blip blip12 = new A.Blip() { Embed = "rId6" };

            A.BlipExtensionList blipExtensionList10 = new A.BlipExtensionList();

            A.BlipExtension blipExtension10 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement49 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId7\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension10.Append(openXmlUnknownElement49);

            blipExtensionList10.Append(blipExtension10);

            blip12.Append(blipExtensionList10);

            A.Stretch stretch8 = new A.Stretch();
            A.FillRectangle fillRectangle8 = new A.FillRectangle();

            stretch8.Append(fillRectangle8);

            blipFill8.Append(blip12);
            blipFill8.Append(stretch8);

            ShapeProperties shapeProperties22 = new ShapeProperties();

            A.Transform2D transform2D20 = new A.Transform2D();
            A.Offset offset28 = new A.Offset() { X = 10453015L, Y = 2709059L };
            A.Extents extents28 = new A.Extents() { Cx = 137250L, Cy = 137157L };

            transform2D20.Append(offset28);
            transform2D20.Append(extents28);

            A.PresetGeometry presetGeometry20 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList20 = new A.AdjustValueList();

            presetGeometry20.Append(adjustValueList20);

            shapeProperties22.Append(transform2D20);
            shapeProperties22.Append(presetGeometry20);

            picture8.Append(nonVisualPictureProperties8);
            picture8.Append(blipFill8);
            picture8.Append(shapeProperties22);

            Shape shape15 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties15 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties31 = new NonVisualDrawingProperties() { Id = (UInt32Value)31U, Name = "Text 48" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList30 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension30 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement50 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{CBE381C6-AAAB-E70B-8CD0-2FC297C9D0E8}\" />");

            // nonVisualDrawingPropertiesExtension30.Append(openXmlUnknownElement50);

            nonVisualDrawingPropertiesExtensionList30.Append(nonVisualDrawingPropertiesExtension30);

            nonVisualDrawingProperties31.Append(nonVisualDrawingPropertiesExtensionList30);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties15 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties31 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties15.Append(nonVisualDrawingProperties31);
            nonVisualShapeProperties15.Append(nonVisualShapeDrawingProperties15);
            nonVisualShapeProperties15.Append(applicationNonVisualDrawingProperties31);

            ShapeProperties shapeProperties23 = new ShapeProperties();

            A.Transform2D transform2D21 = new A.Transform2D();
            A.Offset offset29 = new A.Offset() { X = 10658873L, Y = 2651915L };
            A.Extents extents29 = new A.Extents() { Cx = 526081L, Cy = 228571L };

            transform2D21.Append(offset29);
            transform2D21.Append(extents29);

            A.PresetGeometry presetGeometry21 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList21 = new A.AdjustValueList();

            presetGeometry21.Append(adjustValueList21);
            A.NoFill noFill153 = new A.NoFill();
            A.Outline outline27 = new A.Outline();

            shapeProperties23.Append(transform2D21);
            shapeProperties23.Append(presetGeometry21);
            shapeProperties23.Append(noFill153);
            shapeProperties23.Append(outline27);

            TextBody textBody36 = new TextBody();
            A.BodyProperties bodyProperties36 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle36 = new A.ListStyle();

            A.Paragraph paragraph36 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties30 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing28 = new A.LineSpacing();
            A.SpacingPercent spacingPercent28 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing28.Append(spacingPercent28);

            paragraphProperties30.Append(lineSpacing28);

            A.Run run24 = new A.Run();

            A.RunProperties runProperties24 = new A.RunProperties() { Language = "en-US", FontSize = 1080, Bold = true };

            A.SolidFill solidFill64 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex24 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha13 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex24.Append(alpha13);

            solidFill64.Append(rgbColorModelHex24);
            A.LatinFont latinFont54 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont39 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont54 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            runProperties24.Append(solidFill64);
            runProperties24.Append(latinFont54);
            runProperties24.Append(eastAsianFont39);
            runProperties24.Append(complexScriptFont54);
            A.Text text24 = new A.Text();
            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count() > 4)
                text24.Text = goalsArrayForSlide.ElementAt(4).target ?? "";
            else
                text24.Text = "NO DATA";
            // ********************************* Custom Code Changes End *********************************

            run24.Append(runProperties24);
            run24.Append(text24);

            A.EndParagraphRunProperties endParagraphRunProperties31 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080, Bold = true };
            A.LatinFont latinFont55 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.ComplexScriptFont complexScriptFont55 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties31.Append(latinFont55);
            endParagraphRunProperties31.Append(complexScriptFont55);

            paragraph36.Append(paragraphProperties30);
            paragraph36.Append(run24);
            paragraph36.Append(endParagraphRunProperties31);

            textBody36.Append(bodyProperties36);
            textBody36.Append(listStyle36);
            textBody36.Append(paragraph36);

            shape15.Append(nonVisualShapeProperties15);
            shape15.Append(shapeProperties23);
            shape15.Append(textBody36);

            groupShape6.Append(nonVisualGroupShapeProperties7);
            groupShape6.Append(groupShapeProperties7);
            groupShape6.Append(picture8);
            groupShape6.Append(shape15);

            Shape shape16 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties16 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties32 = new NonVisualDrawingProperties() { Id = (UInt32Value)32U, Name = "Text 46" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList31 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension31 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement51 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{22EDD6C5-8CCA-33F3-5572-74AF9BB5B674}\" />");

            // nonVisualDrawingPropertiesExtension31.Append(openXmlUnknownElement51);

            nonVisualDrawingPropertiesExtensionList31.Append(nonVisualDrawingPropertiesExtension31);

            nonVisualDrawingProperties32.Append(nonVisualDrawingPropertiesExtensionList31);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties16 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties32 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties16.Append(nonVisualDrawingProperties32);
            nonVisualShapeProperties16.Append(nonVisualShapeDrawingProperties16);
            nonVisualShapeProperties16.Append(applicationNonVisualDrawingProperties32);

            ShapeProperties shapeProperties24 = new ShapeProperties();

            A.Transform2D transform2D22 = new A.Transform2D();
            A.Offset offset30 = new A.Offset() { X = 8355707L, Y = 1773625L };
            A.Extents extents30 = new A.Extents() { Cx = 1280895L, Cy = 228571L };

            transform2D22.Append(offset30);
            transform2D22.Append(extents30);

            A.PresetGeometry presetGeometry22 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList22 = new A.AdjustValueList();

            presetGeometry22.Append(adjustValueList22);
            A.NoFill noFill154 = new A.NoFill();
            A.Outline outline28 = new A.Outline();

            shapeProperties24.Append(transform2D22);
            shapeProperties24.Append(presetGeometry22);
            shapeProperties24.Append(noFill154);
            shapeProperties24.Append(outline28);

            TextBody textBody37 = new TextBody();
            A.BodyProperties bodyProperties37 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle37 = new A.ListStyle();

            A.Paragraph paragraph37 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties31 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing29 = new A.LineSpacing();
            A.SpacingPercent spacingPercent29 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing29.Append(spacingPercent29);

            paragraphProperties31.Append(lineSpacing29);

            A.Run run25 = new A.Run();

            A.RunProperties runProperties25 = new A.RunProperties() { Language = "en-US", FontSize = 1080 };

            A.SolidFill solidFill65 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex25 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha14 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex25.Append(alpha14);

            solidFill65.Append(rgbColorModelHex25);
            A.LatinFont latinFont56 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont40 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont56 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties25.Append(solidFill65);
            runProperties25.Append(latinFont56);
            runProperties25.Append(eastAsianFont40);
            runProperties25.Append(complexScriptFont56);
            A.Text text25 = new A.Text();
            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count() > 0)
                text25.Text = goalsArrayForSlide.ElementAt(0).goalOwner ?? "";
            else
                text25.Text = "NO DATA";
            // ********************************* Custom Code Changes End *********************************


            run25.Append(runProperties25);
            run25.Append(text25);
            A.EndParagraphRunProperties endParagraphRunProperties32 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080 };

            paragraph37.Append(paragraphProperties31);
            paragraph37.Append(run25);
            paragraph37.Append(endParagraphRunProperties32);

            textBody37.Append(bodyProperties37);
            textBody37.Append(listStyle37);
            textBody37.Append(paragraph37);

            shape16.Append(nonVisualShapeProperties16);
            shape16.Append(shapeProperties24);
            shape16.Append(textBody37);

            Shape shape17 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties17 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties33 = new NonVisualDrawingProperties() { Id = (UInt32Value)33U, Name = "Text 46" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList32 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension32 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement52 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{7C80EB8C-25AC-DA07-511F-BF291184ADE1}\" />");

            // nonVisualDrawingPropertiesExtension32.Append(openXmlUnknownElement52);

            nonVisualDrawingPropertiesExtensionList32.Append(nonVisualDrawingPropertiesExtension32);

            nonVisualDrawingProperties33.Append(nonVisualDrawingPropertiesExtensionList32);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties17 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties33 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties17.Append(nonVisualDrawingProperties33);
            nonVisualShapeProperties17.Append(nonVisualShapeDrawingProperties17);
            nonVisualShapeProperties17.Append(applicationNonVisualDrawingProperties33);

            ShapeProperties shapeProperties25 = new ShapeProperties();

            A.Transform2D transform2D23 = new A.Transform2D();
            A.Offset offset31 = new A.Offset() { X = 8353878L, Y = 4845394L };
            A.Extents extents31 = new A.Extents() { Cx = 1280895L, Cy = 228571L };

            transform2D23.Append(offset31);
            transform2D23.Append(extents31);

            A.PresetGeometry presetGeometry23 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList23 = new A.AdjustValueList();

            presetGeometry23.Append(adjustValueList23);
            A.NoFill noFill155 = new A.NoFill();
            A.Outline outline29 = new A.Outline();

            shapeProperties25.Append(transform2D23);
            shapeProperties25.Append(presetGeometry23);
            shapeProperties25.Append(noFill155);
            shapeProperties25.Append(outline29);

            TextBody textBody38 = new TextBody();
            A.BodyProperties bodyProperties38 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle38 = new A.ListStyle();

            A.Paragraph paragraph38 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties32 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing30 = new A.LineSpacing();
            A.SpacingPercent spacingPercent30 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing30.Append(spacingPercent30);

            paragraphProperties32.Append(lineSpacing30);

            A.Run run26 = new A.Run();

            A.RunProperties runProperties26 = new A.RunProperties() { Language = "en-US", FontSize = 1080 };

            A.SolidFill solidFill66 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex26 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha15 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex26.Append(alpha15);

            solidFill66.Append(rgbColorModelHex26);
            A.LatinFont latinFont57 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont41 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont57 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties26.Append(solidFill66);
            runProperties26.Append(latinFont57);
            runProperties26.Append(eastAsianFont41);
            runProperties26.Append(complexScriptFont57);
            A.Text text26 = new A.Text();
            // ********************************* Custom Code Changes Start *********************************
            if (goalsArrayForSlide.Count() > 4)
                text26.Text = goalsArrayForSlide.ElementAt(4).goalOwner ?? "";
            else
                text26.Text = "NO DATA";
            // ********************************* Custom Code Changes End *********************************

            run26.Append(runProperties26);
            run26.Append(text26);
            A.EndParagraphRunProperties endParagraphRunProperties33 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080 };

            paragraph38.Append(paragraphProperties32);
            paragraph38.Append(run26);
            paragraph38.Append(endParagraphRunProperties33);

            textBody38.Append(bodyProperties38);
            textBody38.Append(listStyle38);
            textBody38.Append(paragraph38);

            shape17.Append(nonVisualShapeProperties17);
            shape17.Append(shapeProperties25);
            shape17.Append(textBody38);

            shapeTree1.Append(nonVisualGroupShapeProperties1);
            shapeTree1.Append(groupShapeProperties1);
            shapeTree1.Append(picture1);
            shapeTree1.Append(shape3);
            shapeTree1.Append(shape4);
            shapeTree1.Append(shape5);
            shapeTree1.Append(picture2);
            shapeTree1.Append(graphicFrame1);

            // ****************************** Custom Code Changes Start for Owner Cells *************************************
            // shape 16->Cell 1
            // shape 6->Cell 2
            // shape 7->Cell 3
            // shape 8->Cell 4
            // shape 17->Cell 5
            // shape 9->Cell 6
            Shape[] ownerShapes = [shape16, shape6, shape7, shape8, shape17, shape9];

            for (int i = 0; i < goalsArrayForSlide.Count(); i++)
            {
                shapeTree1.Append(ownerShapes[i]);
            }
            // ****************************** Custom Code Changes End for Owner Cells *************************************

            // ****************************** Custom Code Changes Start for Target Cells *************************************
            // groupShape 1->Cell 1
            // groupShape 2->Cell 2
            // groupShape 3->Cell 3
            // groupShape 4->Cell 4
            // groupShape 6->Cell 5
            // groupShape 5->Cell 6
            GroupShape[] targetGroupShapes = [groupShape1, groupShape2, groupShape3, groupShape4, groupShape6, groupShape5];

            for (int i = 0; i < goalsArrayForSlide.Count(); i++)
            {
                shapeTree1.Append(targetGroupShapes[i]);
            }
            // ****************************** Custom Code Changes End for Target Cells *************************************

            CommonSlideDataExtensionList commonSlideDataExtensionList1 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension1 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId1 = new P14.CreationId() { Val = (UInt32Value)589903529U };
            creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension1.Append(creationId1);

            commonSlideDataExtensionList1.Append(commonSlideDataExtension1);

            commonSlideData1.Append(shapeTree1);
            commonSlideData1.Append(commonSlideDataExtensionList1);

            ColorMapOverride colorMapOverride1 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping1 = new A.MasterColorMapping();

            colorMapOverride1.Append(masterColorMapping1);

            slide1.Append(commonSlideData1);
            slide1.Append(colorMapOverride1);

            slidePart1.Slide = slide1;
        }

        // Generates content of imagePart1.
        private void GenerateImagePart1Content(ImagePart imagePart1)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart1Data);
            imagePart1.FeedData(data);
            data.Close();
        }

        // Generates content of imagePart2.
        private void GenerateImagePart2Content(ImagePart imagePart2)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart2Data);
            imagePart2.FeedData(data);
            data.Close();
        }

        // Generates content of imagePart3.
        private void GenerateImagePart3Content(ImagePart imagePart3)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart3Data);
            imagePart3.FeedData(data);
            data.Close();
        }

        // Generates content of slideLayoutPart1.
        private void GenerateSlideLayoutPart1Content(SlideLayoutPart slideLayoutPart1)
        {
            SlideLayout slideLayout1 = new SlideLayout() { Type = SlideLayoutValues.Title, Preserve = true };
            slideLayout1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData2 = new CommonSlideData() { Name = "Title Slide" };

            ShapeTree shapeTree2 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties8 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties34 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties8 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties34 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties8.Append(nonVisualDrawingProperties34);
            nonVisualGroupShapeProperties8.Append(nonVisualGroupShapeDrawingProperties8);
            nonVisualGroupShapeProperties8.Append(applicationNonVisualDrawingProperties34);

            GroupShapeProperties groupShapeProperties8 = new GroupShapeProperties();

            A.TransformGroup transformGroup8 = new A.TransformGroup();
            A.Offset offset32 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents32 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset8 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents8 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup8.Append(offset32);
            transformGroup8.Append(extents32);
            transformGroup8.Append(childOffset8);
            transformGroup8.Append(childExtents8);

            groupShapeProperties8.Append(transformGroup8);

            Shape shape18 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties18 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties35 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList33 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension33 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement53 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{E26AE054-2C5B-DE0D-2495-95C698D54DA2}\" />");

            // nonVisualDrawingPropertiesExtension33.Append(openXmlUnknownElement53);

            nonVisualDrawingPropertiesExtensionList33.Append(nonVisualDrawingPropertiesExtension33);

            nonVisualDrawingProperties35.Append(nonVisualDrawingPropertiesExtensionList33);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties18 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks3 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties18.Append(shapeLocks3);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties35 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape3 = new PlaceholderShape() { Type = PlaceholderValues.CenteredTitle };

            applicationNonVisualDrawingProperties35.Append(placeholderShape3);

            nonVisualShapeProperties18.Append(nonVisualDrawingProperties35);
            nonVisualShapeProperties18.Append(nonVisualShapeDrawingProperties18);
            nonVisualShapeProperties18.Append(applicationNonVisualDrawingProperties35);

            ShapeProperties shapeProperties26 = new ShapeProperties();

            A.Transform2D transform2D24 = new A.Transform2D();
            A.Offset offset33 = new A.Offset() { X = 1524000L, Y = 1122363L };
            A.Extents extents33 = new A.Extents() { Cx = 9144000L, Cy = 2387600L };

            transform2D24.Append(offset33);
            transform2D24.Append(extents33);

            shapeProperties26.Append(transform2D24);

            TextBody textBody39 = new TextBody();
            A.BodyProperties bodyProperties39 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle39 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties2 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };
            A.DefaultRunProperties defaultRunProperties26 = new A.DefaultRunProperties() { FontSize = 6000 };

            level1ParagraphProperties2.Append(defaultRunProperties26);

            listStyle39.Append(level1ParagraphProperties2);

            A.Paragraph paragraph39 = new A.Paragraph();

            A.Run run27 = new A.Run();
            A.RunProperties runProperties27 = new A.RunProperties() { Language = "en-GB" };
            A.Text text27 = new A.Text();
            text27.Text = "Click to edit Master title style";

            run27.Append(runProperties27);
            run27.Append(text27);
            A.EndParagraphRunProperties endParagraphRunProperties34 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph39.Append(run27);
            paragraph39.Append(endParagraphRunProperties34);

            textBody39.Append(bodyProperties39);
            textBody39.Append(listStyle39);
            textBody39.Append(paragraph39);

            shape18.Append(nonVisualShapeProperties18);
            shape18.Append(shapeProperties26);
            shape18.Append(textBody39);

            Shape shape19 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties19 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties36 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Subtitle 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList34 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension34 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement54 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{6A0B1F01-DF1C-6CE1-FC40-F178EE72915F}\" />");

            // nonVisualDrawingPropertiesExtension34.Append(openXmlUnknownElement54);

            nonVisualDrawingPropertiesExtensionList34.Append(nonVisualDrawingPropertiesExtension34);

            nonVisualDrawingProperties36.Append(nonVisualDrawingPropertiesExtensionList34);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties19 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks4 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties19.Append(shapeLocks4);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties36 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape4 = new PlaceholderShape() { Type = PlaceholderValues.SubTitle, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties36.Append(placeholderShape4);

            nonVisualShapeProperties19.Append(nonVisualDrawingProperties36);
            nonVisualShapeProperties19.Append(nonVisualShapeDrawingProperties19);
            nonVisualShapeProperties19.Append(applicationNonVisualDrawingProperties36);

            ShapeProperties shapeProperties27 = new ShapeProperties();

            A.Transform2D transform2D25 = new A.Transform2D();
            A.Offset offset34 = new A.Offset() { X = 1524000L, Y = 3602038L };
            A.Extents extents34 = new A.Extents() { Cx = 9144000L, Cy = 1655762L };

            transform2D25.Append(offset34);
            transform2D25.Append(extents34);

            shapeProperties27.Append(transform2D25);

            TextBody textBody40 = new TextBody();
            A.BodyProperties bodyProperties40 = new A.BodyProperties();

            A.ListStyle listStyle40 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties3 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet15 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties27 = new A.DefaultRunProperties() { FontSize = 2400 };

            level1ParagraphProperties3.Append(noBullet15);
            level1ParagraphProperties3.Append(defaultRunProperties27);

            A.Level2ParagraphProperties level2ParagraphProperties2 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet16 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties28 = new A.DefaultRunProperties() { FontSize = 2000 };

            level2ParagraphProperties2.Append(noBullet16);
            level2ParagraphProperties2.Append(defaultRunProperties28);

            A.Level3ParagraphProperties level3ParagraphProperties2 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet17 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties29 = new A.DefaultRunProperties() { FontSize = 1800 };

            level3ParagraphProperties2.Append(noBullet17);
            level3ParagraphProperties2.Append(defaultRunProperties29);

            A.Level4ParagraphProperties level4ParagraphProperties2 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet18 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties30 = new A.DefaultRunProperties() { FontSize = 1600 };

            level4ParagraphProperties2.Append(noBullet18);
            level4ParagraphProperties2.Append(defaultRunProperties30);

            A.Level5ParagraphProperties level5ParagraphProperties2 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet19 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties31 = new A.DefaultRunProperties() { FontSize = 1600 };

            level5ParagraphProperties2.Append(noBullet19);
            level5ParagraphProperties2.Append(defaultRunProperties31);

            A.Level6ParagraphProperties level6ParagraphProperties2 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet20 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties32 = new A.DefaultRunProperties() { FontSize = 1600 };

            level6ParagraphProperties2.Append(noBullet20);
            level6ParagraphProperties2.Append(defaultRunProperties32);

            A.Level7ParagraphProperties level7ParagraphProperties2 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet21 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties33 = new A.DefaultRunProperties() { FontSize = 1600 };

            level7ParagraphProperties2.Append(noBullet21);
            level7ParagraphProperties2.Append(defaultRunProperties33);

            A.Level8ParagraphProperties level8ParagraphProperties2 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet22 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties34 = new A.DefaultRunProperties() { FontSize = 1600 };

            level8ParagraphProperties2.Append(noBullet22);
            level8ParagraphProperties2.Append(defaultRunProperties34);

            A.Level9ParagraphProperties level9ParagraphProperties2 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet23 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties35 = new A.DefaultRunProperties() { FontSize = 1600 };

            level9ParagraphProperties2.Append(noBullet23);
            level9ParagraphProperties2.Append(defaultRunProperties35);

            listStyle40.Append(level1ParagraphProperties3);
            listStyle40.Append(level2ParagraphProperties2);
            listStyle40.Append(level3ParagraphProperties2);
            listStyle40.Append(level4ParagraphProperties2);
            listStyle40.Append(level5ParagraphProperties2);
            listStyle40.Append(level6ParagraphProperties2);
            listStyle40.Append(level7ParagraphProperties2);
            listStyle40.Append(level8ParagraphProperties2);
            listStyle40.Append(level9ParagraphProperties2);

            A.Paragraph paragraph40 = new A.Paragraph();

            A.Run run28 = new A.Run();
            A.RunProperties runProperties28 = new A.RunProperties() { Language = "en-GB" };
            A.Text text28 = new A.Text();
            text28.Text = "Click to edit Master subtitle style";

            run28.Append(runProperties28);
            run28.Append(text28);
            A.EndParagraphRunProperties endParagraphRunProperties35 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph40.Append(run28);
            paragraph40.Append(endParagraphRunProperties35);

            textBody40.Append(bodyProperties40);
            textBody40.Append(listStyle40);
            textBody40.Append(paragraph40);

            shape19.Append(nonVisualShapeProperties19);
            shape19.Append(shapeProperties27);
            shape19.Append(textBody40);

            Shape shape20 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties20 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties37 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList35 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension35 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement55 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{B7B2C852-B3C3-56C1-F089-BAB3223BE621}\" />");

            // nonVisualDrawingPropertiesExtension35.Append(openXmlUnknownElement55);

            nonVisualDrawingPropertiesExtensionList35.Append(nonVisualDrawingPropertiesExtension35);

            nonVisualDrawingProperties37.Append(nonVisualDrawingPropertiesExtensionList35);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties20 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks5 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties20.Append(shapeLocks5);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties37 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape5 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties37.Append(placeholderShape5);

            nonVisualShapeProperties20.Append(nonVisualDrawingProperties37);
            nonVisualShapeProperties20.Append(nonVisualShapeDrawingProperties20);
            nonVisualShapeProperties20.Append(applicationNonVisualDrawingProperties37);
            ShapeProperties shapeProperties28 = new ShapeProperties();

            TextBody textBody41 = new TextBody();
            A.BodyProperties bodyProperties41 = new A.BodyProperties();
            A.ListStyle listStyle41 = new A.ListStyle();

            A.Paragraph paragraph41 = new A.Paragraph();

            A.Field field1 = new A.Field() { Id = "{C453C8C6-607B-FF49-A93F-867C282599C0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties29 = new A.RunProperties() { Language = "en-US" };
            runProperties29.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text29 = new A.Text();
            text29.Text = "5/3/2024";

            field1.Append(runProperties29);
            field1.Append(text29);
            A.EndParagraphRunProperties endParagraphRunProperties36 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph41.Append(field1);
            paragraph41.Append(endParagraphRunProperties36);

            textBody41.Append(bodyProperties41);
            textBody41.Append(listStyle41);
            textBody41.Append(paragraph41);

            shape20.Append(nonVisualShapeProperties20);
            shape20.Append(shapeProperties28);
            shape20.Append(textBody41);

            Shape shape21 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties21 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties38 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList36 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension36 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement56 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{1BAC4464-FA41-3684-228B-645C8B8E7CD9}\" />");

            // nonVisualDrawingPropertiesExtension36.Append(openXmlUnknownElement56);

            nonVisualDrawingPropertiesExtensionList36.Append(nonVisualDrawingPropertiesExtension36);

            nonVisualDrawingProperties38.Append(nonVisualDrawingPropertiesExtensionList36);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties21 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks6 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties21.Append(shapeLocks6);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties38 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape6 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties38.Append(placeholderShape6);

            nonVisualShapeProperties21.Append(nonVisualDrawingProperties38);
            nonVisualShapeProperties21.Append(nonVisualShapeDrawingProperties21);
            nonVisualShapeProperties21.Append(applicationNonVisualDrawingProperties38);
            ShapeProperties shapeProperties29 = new ShapeProperties();

            TextBody textBody42 = new TextBody();
            A.BodyProperties bodyProperties42 = new A.BodyProperties();
            A.ListStyle listStyle42 = new A.ListStyle();

            A.Paragraph paragraph42 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties37 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph42.Append(endParagraphRunProperties37);

            textBody42.Append(bodyProperties42);
            textBody42.Append(listStyle42);
            textBody42.Append(paragraph42);

            shape21.Append(nonVisualShapeProperties21);
            shape21.Append(shapeProperties29);
            shape21.Append(textBody42);

            Shape shape22 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties22 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties39 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList37 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension37 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement57 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{697483A4-82E9-25E9-8D01-46A7BAA77CA9}\" />");

            // nonVisualDrawingPropertiesExtension37.Append(openXmlUnknownElement57);

            nonVisualDrawingPropertiesExtensionList37.Append(nonVisualDrawingPropertiesExtension37);

            nonVisualDrawingProperties39.Append(nonVisualDrawingPropertiesExtensionList37);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties22 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks7 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties22.Append(shapeLocks7);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties39 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape7 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties39.Append(placeholderShape7);

            nonVisualShapeProperties22.Append(nonVisualDrawingProperties39);
            nonVisualShapeProperties22.Append(nonVisualShapeDrawingProperties22);
            nonVisualShapeProperties22.Append(applicationNonVisualDrawingProperties39);
            ShapeProperties shapeProperties30 = new ShapeProperties();

            TextBody textBody43 = new TextBody();
            A.BodyProperties bodyProperties43 = new A.BodyProperties();
            A.ListStyle listStyle43 = new A.ListStyle();

            A.Paragraph paragraph43 = new A.Paragraph();

            A.Field field2 = new A.Field() { Id = "{402C1264-1A6A-2D47-B539-28043A88F412}", Type = "slidenum" };

            A.RunProperties runProperties30 = new A.RunProperties() { Language = "en-US" };
            runProperties30.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text30 = new A.Text();
            text30.Text = "‹#›";

            field2.Append(runProperties30);
            field2.Append(text30);
            A.EndParagraphRunProperties endParagraphRunProperties38 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph43.Append(field2);
            paragraph43.Append(endParagraphRunProperties38);

            textBody43.Append(bodyProperties43);
            textBody43.Append(listStyle43);
            textBody43.Append(paragraph43);

            shape22.Append(nonVisualShapeProperties22);
            shape22.Append(shapeProperties30);
            shape22.Append(textBody43);

            shapeTree2.Append(nonVisualGroupShapeProperties8);
            shapeTree2.Append(groupShapeProperties8);
            shapeTree2.Append(shape18);
            shapeTree2.Append(shape19);
            shapeTree2.Append(shape20);
            shapeTree2.Append(shape21);
            shapeTree2.Append(shape22);

            CommonSlideDataExtensionList commonSlideDataExtensionList2 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension2 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId2 = new P14.CreationId() { Val = (UInt32Value)3014896269U };
            creationId2.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension2.Append(creationId2);

            commonSlideDataExtensionList2.Append(commonSlideDataExtension2);

            commonSlideData2.Append(shapeTree2);
            commonSlideData2.Append(commonSlideDataExtensionList2);

            ColorMapOverride colorMapOverride2 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping2 = new A.MasterColorMapping();

            colorMapOverride2.Append(masterColorMapping2);

            slideLayout1.Append(commonSlideData2);
            slideLayout1.Append(colorMapOverride2);

            slideLayoutPart1.SlideLayout = slideLayout1;
        }

        // Generates content of slideMasterPart1.
        private void GenerateSlideMasterPart1Content(SlideMasterPart slideMasterPart1)
        {
            SlideMaster slideMaster1 = new SlideMaster();
            slideMaster1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideMaster1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideMaster1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData3 = new CommonSlideData();

            Background background1 = new Background();

            BackgroundStyleReference backgroundStyleReference1 = new BackgroundStyleReference() { Index = (UInt32Value)1001U };
            A.SchemeColor schemeColor34 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            backgroundStyleReference1.Append(schemeColor34);

            background1.Append(backgroundStyleReference1);

            ShapeTree shapeTree3 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties9 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties40 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties9 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties40 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties9.Append(nonVisualDrawingProperties40);
            nonVisualGroupShapeProperties9.Append(nonVisualGroupShapeDrawingProperties9);
            nonVisualGroupShapeProperties9.Append(applicationNonVisualDrawingProperties40);

            GroupShapeProperties groupShapeProperties9 = new GroupShapeProperties();

            A.TransformGroup transformGroup9 = new A.TransformGroup();
            A.Offset offset35 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents35 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset9 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents9 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup9.Append(offset35);
            transformGroup9.Append(extents35);
            transformGroup9.Append(childOffset9);
            transformGroup9.Append(childExtents9);

            groupShapeProperties9.Append(transformGroup9);

            Shape shape23 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties23 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties41 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList38 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension38 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement58 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{F1ADB2EF-0371-0BC2-FECB-7CEE0A92CCE9}\" />");

            // nonVisualDrawingPropertiesExtension38.Append(openXmlUnknownElement58);

            nonVisualDrawingPropertiesExtensionList38.Append(nonVisualDrawingPropertiesExtension38);

            nonVisualDrawingProperties41.Append(nonVisualDrawingPropertiesExtensionList38);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties23 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks8 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties23.Append(shapeLocks8);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties41 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape8 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties41.Append(placeholderShape8);

            nonVisualShapeProperties23.Append(nonVisualDrawingProperties41);
            nonVisualShapeProperties23.Append(nonVisualShapeDrawingProperties23);
            nonVisualShapeProperties23.Append(applicationNonVisualDrawingProperties41);

            ShapeProperties shapeProperties31 = new ShapeProperties();

            A.Transform2D transform2D26 = new A.Transform2D();
            A.Offset offset36 = new A.Offset() { X = 838200L, Y = 365125L };
            A.Extents extents36 = new A.Extents() { Cx = 10515600L, Cy = 1325563L };

            transform2D26.Append(offset36);
            transform2D26.Append(extents36);

            A.PresetGeometry presetGeometry24 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList24 = new A.AdjustValueList();

            presetGeometry24.Append(adjustValueList24);

            shapeProperties31.Append(transform2D26);
            shapeProperties31.Append(presetGeometry24);

            TextBody textBody44 = new TextBody();

            A.BodyProperties bodyProperties44 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.NormalAutoFit normalAutoFit1 = new A.NormalAutoFit();

            bodyProperties44.Append(normalAutoFit1);
            A.ListStyle listStyle44 = new A.ListStyle();

            A.Paragraph paragraph44 = new A.Paragraph();

            A.Run run29 = new A.Run();
            A.RunProperties runProperties31 = new A.RunProperties() { Language = "en-GB" };
            A.Text text31 = new A.Text();
            text31.Text = "Click to edit Master title style";

            run29.Append(runProperties31);
            run29.Append(text31);
            A.EndParagraphRunProperties endParagraphRunProperties39 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph44.Append(run29);
            paragraph44.Append(endParagraphRunProperties39);

            textBody44.Append(bodyProperties44);
            textBody44.Append(listStyle44);
            textBody44.Append(paragraph44);

            shape23.Append(nonVisualShapeProperties23);
            shape23.Append(shapeProperties31);
            shape23.Append(textBody44);

            Shape shape24 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties24 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties42 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Text Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList39 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension39 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement59 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{02A5A080-7632-4B44-0D25-2CCC0391D35F}\" />");

            // nonVisualDrawingPropertiesExtension39.Append(openXmlUnknownElement59);

            nonVisualDrawingPropertiesExtensionList39.Append(nonVisualDrawingPropertiesExtension39);

            nonVisualDrawingProperties42.Append(nonVisualDrawingPropertiesExtensionList39);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties24 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks9 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties24.Append(shapeLocks9);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties42 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape9 = new PlaceholderShape() { Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties42.Append(placeholderShape9);

            nonVisualShapeProperties24.Append(nonVisualDrawingProperties42);
            nonVisualShapeProperties24.Append(nonVisualShapeDrawingProperties24);
            nonVisualShapeProperties24.Append(applicationNonVisualDrawingProperties42);

            ShapeProperties shapeProperties32 = new ShapeProperties();

            A.Transform2D transform2D27 = new A.Transform2D();
            A.Offset offset37 = new A.Offset() { X = 838200L, Y = 1825625L };
            A.Extents extents37 = new A.Extents() { Cx = 10515600L, Cy = 4351338L };

            transform2D27.Append(offset37);
            transform2D27.Append(extents37);

            A.PresetGeometry presetGeometry25 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList25 = new A.AdjustValueList();

            presetGeometry25.Append(adjustValueList25);

            shapeProperties32.Append(transform2D27);
            shapeProperties32.Append(presetGeometry25);

            TextBody textBody45 = new TextBody();

            A.BodyProperties bodyProperties45 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false };
            A.NormalAutoFit normalAutoFit2 = new A.NormalAutoFit();

            bodyProperties45.Append(normalAutoFit2);
            A.ListStyle listStyle45 = new A.ListStyle();

            A.Paragraph paragraph45 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties33 = new A.ParagraphProperties() { Level = 0 };

            A.Run run30 = new A.Run();
            A.RunProperties runProperties32 = new A.RunProperties() { Language = "en-GB" };
            A.Text text32 = new A.Text();
            text32.Text = "Click to edit Master text styles";

            run30.Append(runProperties32);
            run30.Append(text32);

            paragraph45.Append(paragraphProperties33);
            paragraph45.Append(run30);

            A.Paragraph paragraph46 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties34 = new A.ParagraphProperties() { Level = 1 };

            A.Run run31 = new A.Run();
            A.RunProperties runProperties33 = new A.RunProperties() { Language = "en-GB" };
            A.Text text33 = new A.Text();
            text33.Text = "Second level";

            run31.Append(runProperties33);
            run31.Append(text33);

            paragraph46.Append(paragraphProperties34);
            paragraph46.Append(run31);

            A.Paragraph paragraph47 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties35 = new A.ParagraphProperties() { Level = 2 };

            A.Run run32 = new A.Run();
            A.RunProperties runProperties34 = new A.RunProperties() { Language = "en-GB" };
            A.Text text34 = new A.Text();
            text34.Text = "Third level";

            run32.Append(runProperties34);
            run32.Append(text34);

            paragraph47.Append(paragraphProperties35);
            paragraph47.Append(run32);

            A.Paragraph paragraph48 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties36 = new A.ParagraphProperties() { Level = 3 };

            A.Run run33 = new A.Run();
            A.RunProperties runProperties35 = new A.RunProperties() { Language = "en-GB" };
            A.Text text35 = new A.Text();
            text35.Text = "Fourth level";

            run33.Append(runProperties35);
            run33.Append(text35);

            paragraph48.Append(paragraphProperties36);
            paragraph48.Append(run33);

            A.Paragraph paragraph49 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties37 = new A.ParagraphProperties() { Level = 4 };

            A.Run run34 = new A.Run();
            A.RunProperties runProperties36 = new A.RunProperties() { Language = "en-GB" };
            A.Text text36 = new A.Text();
            text36.Text = "Fifth level";

            run34.Append(runProperties36);
            run34.Append(text36);
            A.EndParagraphRunProperties endParagraphRunProperties40 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph49.Append(paragraphProperties37);
            paragraph49.Append(run34);
            paragraph49.Append(endParagraphRunProperties40);

            textBody45.Append(bodyProperties45);
            textBody45.Append(listStyle45);
            textBody45.Append(paragraph45);
            textBody45.Append(paragraph46);
            textBody45.Append(paragraph47);
            textBody45.Append(paragraph48);
            textBody45.Append(paragraph49);

            shape24.Append(nonVisualShapeProperties24);
            shape24.Append(shapeProperties32);
            shape24.Append(textBody45);

            Shape shape25 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties25 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties43 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList40 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension40 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement60 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{57FCBA44-6A75-4AAB-639A-AAE9A9692CB7}\" />");

            // nonVisualDrawingPropertiesExtension40.Append(openXmlUnknownElement60);

            nonVisualDrawingPropertiesExtensionList40.Append(nonVisualDrawingPropertiesExtension40);

            nonVisualDrawingProperties43.Append(nonVisualDrawingPropertiesExtensionList40);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties25 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks10 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties25.Append(shapeLocks10);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties43 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape10 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties43.Append(placeholderShape10);

            nonVisualShapeProperties25.Append(nonVisualDrawingProperties43);
            nonVisualShapeProperties25.Append(nonVisualShapeDrawingProperties25);
            nonVisualShapeProperties25.Append(applicationNonVisualDrawingProperties43);

            ShapeProperties shapeProperties33 = new ShapeProperties();

            A.Transform2D transform2D28 = new A.Transform2D();
            A.Offset offset38 = new A.Offset() { X = 838200L, Y = 6356350L };
            A.Extents extents38 = new A.Extents() { Cx = 2743200L, Cy = 365125L };

            transform2D28.Append(offset38);
            transform2D28.Append(extents38);

            A.PresetGeometry presetGeometry26 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList26 = new A.AdjustValueList();

            presetGeometry26.Append(adjustValueList26);

            shapeProperties33.Append(transform2D28);
            shapeProperties33.Append(presetGeometry26);

            TextBody textBody46 = new TextBody();
            A.BodyProperties bodyProperties46 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle46 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties4 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Left };

            A.DefaultRunProperties defaultRunProperties36 = new A.DefaultRunProperties() { FontSize = 1200 };

            A.SolidFill solidFill67 = new A.SolidFill();

            A.SchemeColor schemeColor35 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint1 = new A.Tint() { Val = 82000 };

            schemeColor35.Append(tint1);

            solidFill67.Append(schemeColor35);

            defaultRunProperties36.Append(solidFill67);

            level1ParagraphProperties4.Append(defaultRunProperties36);

            listStyle46.Append(level1ParagraphProperties4);

            A.Paragraph paragraph50 = new A.Paragraph();

            A.Field field3 = new A.Field() { Id = "{C453C8C6-607B-FF49-A93F-867C282599C0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties37 = new A.RunProperties() { Language = "en-US" };
            runProperties37.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text37 = new A.Text();
            text37.Text = "5/3/2024";

            field3.Append(runProperties37);
            field3.Append(text37);
            A.EndParagraphRunProperties endParagraphRunProperties41 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph50.Append(field3);
            paragraph50.Append(endParagraphRunProperties41);

            textBody46.Append(bodyProperties46);
            textBody46.Append(listStyle46);
            textBody46.Append(paragraph50);

            shape25.Append(nonVisualShapeProperties25);
            shape25.Append(shapeProperties33);
            shape25.Append(textBody46);

            Shape shape26 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties26 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties44 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList41 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension41 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement61 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{24C1BE9A-A7BA-A7E0-F73C-7914597F0944}\" />");

            // nonVisualDrawingPropertiesExtension41.Append(openXmlUnknownElement61);

            nonVisualDrawingPropertiesExtensionList41.Append(nonVisualDrawingPropertiesExtension41);

            nonVisualDrawingProperties44.Append(nonVisualDrawingPropertiesExtensionList41);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties26 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks11 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties26.Append(shapeLocks11);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties44 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape11 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)3U };

            applicationNonVisualDrawingProperties44.Append(placeholderShape11);

            nonVisualShapeProperties26.Append(nonVisualDrawingProperties44);
            nonVisualShapeProperties26.Append(nonVisualShapeDrawingProperties26);
            nonVisualShapeProperties26.Append(applicationNonVisualDrawingProperties44);

            ShapeProperties shapeProperties34 = new ShapeProperties();

            A.Transform2D transform2D29 = new A.Transform2D();
            A.Offset offset39 = new A.Offset() { X = 4038600L, Y = 6356350L };
            A.Extents extents39 = new A.Extents() { Cx = 4114800L, Cy = 365125L };

            transform2D29.Append(offset39);
            transform2D29.Append(extents39);

            A.PresetGeometry presetGeometry27 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList27 = new A.AdjustValueList();

            presetGeometry27.Append(adjustValueList27);

            shapeProperties34.Append(transform2D29);
            shapeProperties34.Append(presetGeometry27);

            TextBody textBody47 = new TextBody();
            A.BodyProperties bodyProperties47 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle47 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties5 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.DefaultRunProperties defaultRunProperties37 = new A.DefaultRunProperties() { FontSize = 1200 };

            A.SolidFill solidFill68 = new A.SolidFill();

            A.SchemeColor schemeColor36 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint2 = new A.Tint() { Val = 82000 };

            schemeColor36.Append(tint2);

            solidFill68.Append(schemeColor36);

            defaultRunProperties37.Append(solidFill68);

            level1ParagraphProperties5.Append(defaultRunProperties37);

            listStyle47.Append(level1ParagraphProperties5);

            A.Paragraph paragraph51 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties42 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph51.Append(endParagraphRunProperties42);

            textBody47.Append(bodyProperties47);
            textBody47.Append(listStyle47);
            textBody47.Append(paragraph51);

            shape26.Append(nonVisualShapeProperties26);
            shape26.Append(shapeProperties34);
            shape26.Append(textBody47);

            Shape shape27 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties27 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties45 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList42 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension42 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement62 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{53B084B7-CA9B-D9A1-9FB7-8EAAD5D9BEFD}\" />");

            // nonVisualDrawingPropertiesExtension42.Append(openXmlUnknownElement62);

            nonVisualDrawingPropertiesExtensionList42.Append(nonVisualDrawingPropertiesExtension42);

            nonVisualDrawingProperties45.Append(nonVisualDrawingPropertiesExtensionList42);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties27 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks12 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties27.Append(shapeLocks12);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties45 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape12 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)4U };

            applicationNonVisualDrawingProperties45.Append(placeholderShape12);

            nonVisualShapeProperties27.Append(nonVisualDrawingProperties45);
            nonVisualShapeProperties27.Append(nonVisualShapeDrawingProperties27);
            nonVisualShapeProperties27.Append(applicationNonVisualDrawingProperties45);

            ShapeProperties shapeProperties35 = new ShapeProperties();

            A.Transform2D transform2D30 = new A.Transform2D();
            A.Offset offset40 = new A.Offset() { X = 8610600L, Y = 6356350L };
            A.Extents extents40 = new A.Extents() { Cx = 2743200L, Cy = 365125L };

            transform2D30.Append(offset40);
            transform2D30.Append(extents40);

            A.PresetGeometry presetGeometry28 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList28 = new A.AdjustValueList();

            presetGeometry28.Append(adjustValueList28);

            shapeProperties35.Append(transform2D30);
            shapeProperties35.Append(presetGeometry28);

            TextBody textBody48 = new TextBody();
            A.BodyProperties bodyProperties48 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle48 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties6 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Right };

            A.DefaultRunProperties defaultRunProperties38 = new A.DefaultRunProperties() { FontSize = 1200 };

            A.SolidFill solidFill69 = new A.SolidFill();

            A.SchemeColor schemeColor37 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint3 = new A.Tint() { Val = 82000 };

            schemeColor37.Append(tint3);

            solidFill69.Append(schemeColor37);

            defaultRunProperties38.Append(solidFill69);

            level1ParagraphProperties6.Append(defaultRunProperties38);

            listStyle48.Append(level1ParagraphProperties6);

            A.Paragraph paragraph52 = new A.Paragraph();

            A.Field field4 = new A.Field() { Id = "{402C1264-1A6A-2D47-B539-28043A88F412}", Type = "slidenum" };

            A.RunProperties runProperties38 = new A.RunProperties() { Language = "en-US" };
            runProperties38.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text38 = new A.Text();
            text38.Text = "‹#›";

            field4.Append(runProperties38);
            field4.Append(text38);
            A.EndParagraphRunProperties endParagraphRunProperties43 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph52.Append(field4);
            paragraph52.Append(endParagraphRunProperties43);

            textBody48.Append(bodyProperties48);
            textBody48.Append(listStyle48);
            textBody48.Append(paragraph52);

            shape27.Append(nonVisualShapeProperties27);
            shape27.Append(shapeProperties35);
            shape27.Append(textBody48);

            shapeTree3.Append(nonVisualGroupShapeProperties9);
            shapeTree3.Append(groupShapeProperties9);
            shapeTree3.Append(shape23);
            shapeTree3.Append(shape24);
            shapeTree3.Append(shape25);
            shapeTree3.Append(shape26);
            shapeTree3.Append(shape27);

            CommonSlideDataExtensionList commonSlideDataExtensionList3 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension3 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId3 = new P14.CreationId() { Val = (UInt32Value)2391460576U };
            creationId3.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension3.Append(creationId3);

            commonSlideDataExtensionList3.Append(commonSlideDataExtension3);

            commonSlideData3.Append(background1);
            commonSlideData3.Append(shapeTree3);
            commonSlideData3.Append(commonSlideDataExtensionList3);
            ColorMap colorMap1 = new ColorMap() { Background1 = A.ColorSchemeIndexValues.Light1, Text1 = A.ColorSchemeIndexValues.Dark1, Background2 = A.ColorSchemeIndexValues.Light2, Text2 = A.ColorSchemeIndexValues.Dark2, Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2, Accent3 = A.ColorSchemeIndexValues.Accent3, Accent4 = A.ColorSchemeIndexValues.Accent4, Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6, Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink };

            SlideLayoutIdList slideLayoutIdList1 = new SlideLayoutIdList();
            SlideLayoutId slideLayoutId1 = new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" };
            SlideLayoutId slideLayoutId2 = new SlideLayoutId() { Id = (UInt32Value)2147483650U, RelationshipId = "rId2" };
            SlideLayoutId slideLayoutId3 = new SlideLayoutId() { Id = (UInt32Value)2147483651U, RelationshipId = "rId3" };
            SlideLayoutId slideLayoutId4 = new SlideLayoutId() { Id = (UInt32Value)2147483652U, RelationshipId = "rId4" };
            SlideLayoutId slideLayoutId5 = new SlideLayoutId() { Id = (UInt32Value)2147483653U, RelationshipId = "rId5" };
            SlideLayoutId slideLayoutId6 = new SlideLayoutId() { Id = (UInt32Value)2147483654U, RelationshipId = "rId6" };
            SlideLayoutId slideLayoutId7 = new SlideLayoutId() { Id = (UInt32Value)2147483655U, RelationshipId = "rId7" };
            SlideLayoutId slideLayoutId8 = new SlideLayoutId() { Id = (UInt32Value)2147483656U, RelationshipId = "rId8" };
            SlideLayoutId slideLayoutId9 = new SlideLayoutId() { Id = (UInt32Value)2147483657U, RelationshipId = "rId9" };
            SlideLayoutId slideLayoutId10 = new SlideLayoutId() { Id = (UInt32Value)2147483658U, RelationshipId = "rId10" };
            SlideLayoutId slideLayoutId11 = new SlideLayoutId() { Id = (UInt32Value)2147483659U, RelationshipId = "rId11" };

            slideLayoutIdList1.Append(slideLayoutId1);
            slideLayoutIdList1.Append(slideLayoutId2);
            slideLayoutIdList1.Append(slideLayoutId3);
            slideLayoutIdList1.Append(slideLayoutId4);
            slideLayoutIdList1.Append(slideLayoutId5);
            slideLayoutIdList1.Append(slideLayoutId6);
            slideLayoutIdList1.Append(slideLayoutId7);
            slideLayoutIdList1.Append(slideLayoutId8);
            slideLayoutIdList1.Append(slideLayoutId9);
            slideLayoutIdList1.Append(slideLayoutId10);
            slideLayoutIdList1.Append(slideLayoutId11);

            TextStyles textStyles1 = new TextStyles();

            TitleStyle titleStyle1 = new TitleStyle();

            A.Level1ParagraphProperties level1ParagraphProperties7 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing31 = new A.LineSpacing();
            A.SpacingPercent spacingPercent31 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing31.Append(spacingPercent31);

            A.SpaceBefore spaceBefore17 = new A.SpaceBefore();
            A.SpacingPercent spacingPercent32 = new A.SpacingPercent() { Val = 0 };

            spaceBefore17.Append(spacingPercent32);
            A.NoBullet noBullet24 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties39 = new A.DefaultRunProperties() { FontSize = 4400, Kerning = 1200 };

            A.SolidFill solidFill70 = new A.SolidFill();
            A.SchemeColor schemeColor38 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill70.Append(schemeColor38);
            A.LatinFont latinFont58 = new A.LatinFont() { Typeface = "+mj-lt" };
            A.EastAsianFont eastAsianFont42 = new A.EastAsianFont() { Typeface = "+mj-ea" };
            A.ComplexScriptFont complexScriptFont58 = new A.ComplexScriptFont() { Typeface = "+mj-cs" };

            defaultRunProperties39.Append(solidFill70);
            defaultRunProperties39.Append(latinFont58);
            defaultRunProperties39.Append(eastAsianFont42);
            defaultRunProperties39.Append(complexScriptFont58);

            level1ParagraphProperties7.Append(lineSpacing31);
            level1ParagraphProperties7.Append(spaceBefore17);
            level1ParagraphProperties7.Append(noBullet24);
            level1ParagraphProperties7.Append(defaultRunProperties39);

            titleStyle1.Append(level1ParagraphProperties7);

            BodyStyle bodyStyle1 = new BodyStyle();

            A.Level1ParagraphProperties level1ParagraphProperties8 = new A.Level1ParagraphProperties() { LeftMargin = 228600, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing32 = new A.LineSpacing();
            A.SpacingPercent spacingPercent33 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing32.Append(spacingPercent33);

            A.SpaceBefore spaceBefore18 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints33 = new A.SpacingPoints() { Val = 1000 };

            spaceBefore18.Append(spacingPoints33);
            A.BulletFont bulletFont1 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet1 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties40 = new A.DefaultRunProperties() { FontSize = 2800, Kerning = 1200 };

            A.SolidFill solidFill71 = new A.SolidFill();
            A.SchemeColor schemeColor39 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill71.Append(schemeColor39);
            A.LatinFont latinFont59 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont43 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont59 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties40.Append(solidFill71);
            defaultRunProperties40.Append(latinFont59);
            defaultRunProperties40.Append(eastAsianFont43);
            defaultRunProperties40.Append(complexScriptFont59);

            level1ParagraphProperties8.Append(lineSpacing32);
            level1ParagraphProperties8.Append(spaceBefore18);
            level1ParagraphProperties8.Append(bulletFont1);
            level1ParagraphProperties8.Append(characterBullet1);
            level1ParagraphProperties8.Append(defaultRunProperties40);

            A.Level2ParagraphProperties level2ParagraphProperties3 = new A.Level2ParagraphProperties() { LeftMargin = 685800, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing33 = new A.LineSpacing();
            A.SpacingPercent spacingPercent34 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing33.Append(spacingPercent34);

            A.SpaceBefore spaceBefore19 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints34 = new A.SpacingPoints() { Val = 500 };

            spaceBefore19.Append(spacingPoints34);
            A.BulletFont bulletFont2 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet2 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties41 = new A.DefaultRunProperties() { FontSize = 2400, Kerning = 1200 };

            A.SolidFill solidFill72 = new A.SolidFill();
            A.SchemeColor schemeColor40 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill72.Append(schemeColor40);
            A.LatinFont latinFont60 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont44 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont60 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties41.Append(solidFill72);
            defaultRunProperties41.Append(latinFont60);
            defaultRunProperties41.Append(eastAsianFont44);
            defaultRunProperties41.Append(complexScriptFont60);

            level2ParagraphProperties3.Append(lineSpacing33);
            level2ParagraphProperties3.Append(spaceBefore19);
            level2ParagraphProperties3.Append(bulletFont2);
            level2ParagraphProperties3.Append(characterBullet2);
            level2ParagraphProperties3.Append(defaultRunProperties41);

            A.Level3ParagraphProperties level3ParagraphProperties3 = new A.Level3ParagraphProperties() { LeftMargin = 1143000, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing34 = new A.LineSpacing();
            A.SpacingPercent spacingPercent35 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing34.Append(spacingPercent35);

            A.SpaceBefore spaceBefore20 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints35 = new A.SpacingPoints() { Val = 500 };

            spaceBefore20.Append(spacingPoints35);
            A.BulletFont bulletFont3 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet3 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties42 = new A.DefaultRunProperties() { FontSize = 2000, Kerning = 1200 };

            A.SolidFill solidFill73 = new A.SolidFill();
            A.SchemeColor schemeColor41 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill73.Append(schemeColor41);
            A.LatinFont latinFont61 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont45 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont61 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties42.Append(solidFill73);
            defaultRunProperties42.Append(latinFont61);
            defaultRunProperties42.Append(eastAsianFont45);
            defaultRunProperties42.Append(complexScriptFont61);

            level3ParagraphProperties3.Append(lineSpacing34);
            level3ParagraphProperties3.Append(spaceBefore20);
            level3ParagraphProperties3.Append(bulletFont3);
            level3ParagraphProperties3.Append(characterBullet3);
            level3ParagraphProperties3.Append(defaultRunProperties42);

            A.Level4ParagraphProperties level4ParagraphProperties3 = new A.Level4ParagraphProperties() { LeftMargin = 1600200, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing35 = new A.LineSpacing();
            A.SpacingPercent spacingPercent36 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing35.Append(spacingPercent36);

            A.SpaceBefore spaceBefore21 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints36 = new A.SpacingPoints() { Val = 500 };

            spaceBefore21.Append(spacingPoints36);
            A.BulletFont bulletFont4 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet4 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties43 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill74 = new A.SolidFill();
            A.SchemeColor schemeColor42 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill74.Append(schemeColor42);
            A.LatinFont latinFont62 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont46 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont62 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties43.Append(solidFill74);
            defaultRunProperties43.Append(latinFont62);
            defaultRunProperties43.Append(eastAsianFont46);
            defaultRunProperties43.Append(complexScriptFont62);

            level4ParagraphProperties3.Append(lineSpacing35);
            level4ParagraphProperties3.Append(spaceBefore21);
            level4ParagraphProperties3.Append(bulletFont4);
            level4ParagraphProperties3.Append(characterBullet4);
            level4ParagraphProperties3.Append(defaultRunProperties43);

            A.Level5ParagraphProperties level5ParagraphProperties3 = new A.Level5ParagraphProperties() { LeftMargin = 2057400, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing36 = new A.LineSpacing();
            A.SpacingPercent spacingPercent37 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing36.Append(spacingPercent37);

            A.SpaceBefore spaceBefore22 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints37 = new A.SpacingPoints() { Val = 500 };

            spaceBefore22.Append(spacingPoints37);
            A.BulletFont bulletFont5 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet5 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties44 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill75 = new A.SolidFill();
            A.SchemeColor schemeColor43 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill75.Append(schemeColor43);
            A.LatinFont latinFont63 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont47 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont63 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties44.Append(solidFill75);
            defaultRunProperties44.Append(latinFont63);
            defaultRunProperties44.Append(eastAsianFont47);
            defaultRunProperties44.Append(complexScriptFont63);

            level5ParagraphProperties3.Append(lineSpacing36);
            level5ParagraphProperties3.Append(spaceBefore22);
            level5ParagraphProperties3.Append(bulletFont5);
            level5ParagraphProperties3.Append(characterBullet5);
            level5ParagraphProperties3.Append(defaultRunProperties44);

            A.Level6ParagraphProperties level6ParagraphProperties3 = new A.Level6ParagraphProperties() { LeftMargin = 2514600, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing37 = new A.LineSpacing();
            A.SpacingPercent spacingPercent38 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing37.Append(spacingPercent38);

            A.SpaceBefore spaceBefore23 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints38 = new A.SpacingPoints() { Val = 500 };

            spaceBefore23.Append(spacingPoints38);
            A.BulletFont bulletFont6 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet6 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties45 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill76 = new A.SolidFill();
            A.SchemeColor schemeColor44 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill76.Append(schemeColor44);
            A.LatinFont latinFont64 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont48 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont64 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties45.Append(solidFill76);
            defaultRunProperties45.Append(latinFont64);
            defaultRunProperties45.Append(eastAsianFont48);
            defaultRunProperties45.Append(complexScriptFont64);

            level6ParagraphProperties3.Append(lineSpacing37);
            level6ParagraphProperties3.Append(spaceBefore23);
            level6ParagraphProperties3.Append(bulletFont6);
            level6ParagraphProperties3.Append(characterBullet6);
            level6ParagraphProperties3.Append(defaultRunProperties45);

            A.Level7ParagraphProperties level7ParagraphProperties3 = new A.Level7ParagraphProperties() { LeftMargin = 2971800, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing38 = new A.LineSpacing();
            A.SpacingPercent spacingPercent39 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing38.Append(spacingPercent39);

            A.SpaceBefore spaceBefore24 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints39 = new A.SpacingPoints() { Val = 500 };

            spaceBefore24.Append(spacingPoints39);
            A.BulletFont bulletFont7 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet7 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties46 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill77 = new A.SolidFill();
            A.SchemeColor schemeColor45 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill77.Append(schemeColor45);
            A.LatinFont latinFont65 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont49 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont65 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties46.Append(solidFill77);
            defaultRunProperties46.Append(latinFont65);
            defaultRunProperties46.Append(eastAsianFont49);
            defaultRunProperties46.Append(complexScriptFont65);

            level7ParagraphProperties3.Append(lineSpacing38);
            level7ParagraphProperties3.Append(spaceBefore24);
            level7ParagraphProperties3.Append(bulletFont7);
            level7ParagraphProperties3.Append(characterBullet7);
            level7ParagraphProperties3.Append(defaultRunProperties46);

            A.Level8ParagraphProperties level8ParagraphProperties3 = new A.Level8ParagraphProperties() { LeftMargin = 3429000, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing39 = new A.LineSpacing();
            A.SpacingPercent spacingPercent40 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing39.Append(spacingPercent40);

            A.SpaceBefore spaceBefore25 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints40 = new A.SpacingPoints() { Val = 500 };

            spaceBefore25.Append(spacingPoints40);
            A.BulletFont bulletFont8 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet8 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties47 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill78 = new A.SolidFill();
            A.SchemeColor schemeColor46 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill78.Append(schemeColor46);
            A.LatinFont latinFont66 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont50 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont66 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties47.Append(solidFill78);
            defaultRunProperties47.Append(latinFont66);
            defaultRunProperties47.Append(eastAsianFont50);
            defaultRunProperties47.Append(complexScriptFont66);

            level8ParagraphProperties3.Append(lineSpacing39);
            level8ParagraphProperties3.Append(spaceBefore25);
            level8ParagraphProperties3.Append(bulletFont8);
            level8ParagraphProperties3.Append(characterBullet8);
            level8ParagraphProperties3.Append(defaultRunProperties47);

            A.Level9ParagraphProperties level9ParagraphProperties3 = new A.Level9ParagraphProperties() { LeftMargin = 3886200, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing40 = new A.LineSpacing();
            A.SpacingPercent spacingPercent41 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing40.Append(spacingPercent41);

            A.SpaceBefore spaceBefore26 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints41 = new A.SpacingPoints() { Val = 500 };

            spaceBefore26.Append(spacingPoints41);
            A.BulletFont bulletFont9 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet9 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties48 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill79 = new A.SolidFill();
            A.SchemeColor schemeColor47 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill79.Append(schemeColor47);
            A.LatinFont latinFont67 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont51 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont67 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties48.Append(solidFill79);
            defaultRunProperties48.Append(latinFont67);
            defaultRunProperties48.Append(eastAsianFont51);
            defaultRunProperties48.Append(complexScriptFont67);

            level9ParagraphProperties3.Append(lineSpacing40);
            level9ParagraphProperties3.Append(spaceBefore26);
            level9ParagraphProperties3.Append(bulletFont9);
            level9ParagraphProperties3.Append(characterBullet9);
            level9ParagraphProperties3.Append(defaultRunProperties48);

            bodyStyle1.Append(level1ParagraphProperties8);
            bodyStyle1.Append(level2ParagraphProperties3);
            bodyStyle1.Append(level3ParagraphProperties3);
            bodyStyle1.Append(level4ParagraphProperties3);
            bodyStyle1.Append(level5ParagraphProperties3);
            bodyStyle1.Append(level6ParagraphProperties3);
            bodyStyle1.Append(level7ParagraphProperties3);
            bodyStyle1.Append(level8ParagraphProperties3);
            bodyStyle1.Append(level9ParagraphProperties3);

            OtherStyle otherStyle1 = new OtherStyle();

            A.DefaultParagraphProperties defaultParagraphProperties2 = new A.DefaultParagraphProperties();
            A.DefaultRunProperties defaultRunProperties49 = new A.DefaultRunProperties() { Language = "en-US" };

            defaultParagraphProperties2.Append(defaultRunProperties49);

            A.Level1ParagraphProperties level1ParagraphProperties9 = new A.Level1ParagraphProperties() { LeftMargin = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties50 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill80 = new A.SolidFill();
            A.SchemeColor schemeColor48 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill80.Append(schemeColor48);
            A.LatinFont latinFont68 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont52 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont68 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties50.Append(solidFill80);
            defaultRunProperties50.Append(latinFont68);
            defaultRunProperties50.Append(eastAsianFont52);
            defaultRunProperties50.Append(complexScriptFont68);

            level1ParagraphProperties9.Append(defaultRunProperties50);

            A.Level2ParagraphProperties level2ParagraphProperties4 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties51 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill81 = new A.SolidFill();
            A.SchemeColor schemeColor49 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill81.Append(schemeColor49);
            A.LatinFont latinFont69 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont53 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont69 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties51.Append(solidFill81);
            defaultRunProperties51.Append(latinFont69);
            defaultRunProperties51.Append(eastAsianFont53);
            defaultRunProperties51.Append(complexScriptFont69);

            level2ParagraphProperties4.Append(defaultRunProperties51);

            A.Level3ParagraphProperties level3ParagraphProperties4 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties52 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill82 = new A.SolidFill();
            A.SchemeColor schemeColor50 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill82.Append(schemeColor50);
            A.LatinFont latinFont70 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont54 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont70 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties52.Append(solidFill82);
            defaultRunProperties52.Append(latinFont70);
            defaultRunProperties52.Append(eastAsianFont54);
            defaultRunProperties52.Append(complexScriptFont70);

            level3ParagraphProperties4.Append(defaultRunProperties52);

            A.Level4ParagraphProperties level4ParagraphProperties4 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties53 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill83 = new A.SolidFill();
            A.SchemeColor schemeColor51 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill83.Append(schemeColor51);
            A.LatinFont latinFont71 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont55 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont71 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties53.Append(solidFill83);
            defaultRunProperties53.Append(latinFont71);
            defaultRunProperties53.Append(eastAsianFont55);
            defaultRunProperties53.Append(complexScriptFont71);

            level4ParagraphProperties4.Append(defaultRunProperties53);

            A.Level5ParagraphProperties level5ParagraphProperties4 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties54 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill84 = new A.SolidFill();
            A.SchemeColor schemeColor52 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill84.Append(schemeColor52);
            A.LatinFont latinFont72 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont56 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont72 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties54.Append(solidFill84);
            defaultRunProperties54.Append(latinFont72);
            defaultRunProperties54.Append(eastAsianFont56);
            defaultRunProperties54.Append(complexScriptFont72);

            level5ParagraphProperties4.Append(defaultRunProperties54);

            A.Level6ParagraphProperties level6ParagraphProperties4 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties55 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill85 = new A.SolidFill();
            A.SchemeColor schemeColor53 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill85.Append(schemeColor53);
            A.LatinFont latinFont73 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont57 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont73 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties55.Append(solidFill85);
            defaultRunProperties55.Append(latinFont73);
            defaultRunProperties55.Append(eastAsianFont57);
            defaultRunProperties55.Append(complexScriptFont73);

            level6ParagraphProperties4.Append(defaultRunProperties55);

            A.Level7ParagraphProperties level7ParagraphProperties4 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties56 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill86 = new A.SolidFill();
            A.SchemeColor schemeColor54 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill86.Append(schemeColor54);
            A.LatinFont latinFont74 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont58 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont74 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties56.Append(solidFill86);
            defaultRunProperties56.Append(latinFont74);
            defaultRunProperties56.Append(eastAsianFont58);
            defaultRunProperties56.Append(complexScriptFont74);

            level7ParagraphProperties4.Append(defaultRunProperties56);

            A.Level8ParagraphProperties level8ParagraphProperties4 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties57 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill87 = new A.SolidFill();
            A.SchemeColor schemeColor55 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill87.Append(schemeColor55);
            A.LatinFont latinFont75 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont59 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont75 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties57.Append(solidFill87);
            defaultRunProperties57.Append(latinFont75);
            defaultRunProperties57.Append(eastAsianFont59);
            defaultRunProperties57.Append(complexScriptFont75);

            level8ParagraphProperties4.Append(defaultRunProperties57);

            A.Level9ParagraphProperties level9ParagraphProperties4 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties58 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill88 = new A.SolidFill();
            A.SchemeColor schemeColor56 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill88.Append(schemeColor56);
            A.LatinFont latinFont76 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont60 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont76 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties58.Append(solidFill88);
            defaultRunProperties58.Append(latinFont76);
            defaultRunProperties58.Append(eastAsianFont60);
            defaultRunProperties58.Append(complexScriptFont76);

            level9ParagraphProperties4.Append(defaultRunProperties58);

            otherStyle1.Append(defaultParagraphProperties2);
            otherStyle1.Append(level1ParagraphProperties9);
            otherStyle1.Append(level2ParagraphProperties4);
            otherStyle1.Append(level3ParagraphProperties4);
            otherStyle1.Append(level4ParagraphProperties4);
            otherStyle1.Append(level5ParagraphProperties4);
            otherStyle1.Append(level6ParagraphProperties4);
            otherStyle1.Append(level7ParagraphProperties4);
            otherStyle1.Append(level8ParagraphProperties4);
            otherStyle1.Append(level9ParagraphProperties4);

            textStyles1.Append(titleStyle1);
            textStyles1.Append(bodyStyle1);
            textStyles1.Append(otherStyle1);

            slideMaster1.Append(commonSlideData3);
            slideMaster1.Append(colorMap1);
            slideMaster1.Append(slideLayoutIdList1);
            slideMaster1.Append(textStyles1);

            slideMasterPart1.SlideMaster = slideMaster1;
        }

        // Generates content of slideLayoutPart2.
        private void GenerateSlideLayoutPart2Content(SlideLayoutPart slideLayoutPart2)
        {
            SlideLayout slideLayout2 = new SlideLayout() { Type = SlideLayoutValues.ObjectText, Preserve = true };
            slideLayout2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout2.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData4 = new CommonSlideData() { Name = "Content with Caption" };

            ShapeTree shapeTree4 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties10 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties46 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties10 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties46 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties10.Append(nonVisualDrawingProperties46);
            nonVisualGroupShapeProperties10.Append(nonVisualGroupShapeDrawingProperties10);
            nonVisualGroupShapeProperties10.Append(applicationNonVisualDrawingProperties46);

            GroupShapeProperties groupShapeProperties10 = new GroupShapeProperties();

            A.TransformGroup transformGroup10 = new A.TransformGroup();
            A.Offset offset41 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents41 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset10 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents10 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup10.Append(offset41);
            transformGroup10.Append(extents41);
            transformGroup10.Append(childOffset10);
            transformGroup10.Append(childExtents10);

            groupShapeProperties10.Append(transformGroup10);

            Shape shape28 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties28 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties47 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList43 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension43 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement63 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{20E85DB8-061B-7EBB-EB41-B49320094870}\" />");

            // nonVisualDrawingPropertiesExtension43.Append(openXmlUnknownElement63);

            nonVisualDrawingPropertiesExtensionList43.Append(nonVisualDrawingPropertiesExtension43);

            nonVisualDrawingProperties47.Append(nonVisualDrawingPropertiesExtensionList43);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties28 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks13 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties28.Append(shapeLocks13);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties47 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape13 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties47.Append(placeholderShape13);

            nonVisualShapeProperties28.Append(nonVisualDrawingProperties47);
            nonVisualShapeProperties28.Append(nonVisualShapeDrawingProperties28);
            nonVisualShapeProperties28.Append(applicationNonVisualDrawingProperties47);

            ShapeProperties shapeProperties36 = new ShapeProperties();

            A.Transform2D transform2D31 = new A.Transform2D();
            A.Offset offset42 = new A.Offset() { X = 839788L, Y = 457200L };
            A.Extents extents42 = new A.Extents() { Cx = 3932237L, Cy = 1600200L };

            transform2D31.Append(offset42);
            transform2D31.Append(extents42);

            shapeProperties36.Append(transform2D31);

            TextBody textBody49 = new TextBody();
            A.BodyProperties bodyProperties49 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle49 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties10 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties59 = new A.DefaultRunProperties() { FontSize = 3200 };

            level1ParagraphProperties10.Append(defaultRunProperties59);

            listStyle49.Append(level1ParagraphProperties10);

            A.Paragraph paragraph53 = new A.Paragraph();

            A.Run run35 = new A.Run();
            A.RunProperties runProperties39 = new A.RunProperties() { Language = "en-GB" };
            A.Text text39 = new A.Text();
            text39.Text = "Click to edit Master title style";

            run35.Append(runProperties39);
            run35.Append(text39);
            A.EndParagraphRunProperties endParagraphRunProperties44 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph53.Append(run35);
            paragraph53.Append(endParagraphRunProperties44);

            textBody49.Append(bodyProperties49);
            textBody49.Append(listStyle49);
            textBody49.Append(paragraph53);

            shape28.Append(nonVisualShapeProperties28);
            shape28.Append(shapeProperties36);
            shape28.Append(textBody49);

            Shape shape29 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties29 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties48 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Content Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList44 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension44 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement64 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{60550879-5FC0-2415-98A1-E5AD1C6A9A69}\" />");

            // nonVisualDrawingPropertiesExtension44.Append(openXmlUnknownElement64);

            nonVisualDrawingPropertiesExtensionList44.Append(nonVisualDrawingPropertiesExtension44);

            nonVisualDrawingProperties48.Append(nonVisualDrawingPropertiesExtensionList44);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties29 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks14 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties29.Append(shapeLocks14);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties48 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape14 = new PlaceholderShape() { Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties48.Append(placeholderShape14);

            nonVisualShapeProperties29.Append(nonVisualDrawingProperties48);
            nonVisualShapeProperties29.Append(nonVisualShapeDrawingProperties29);
            nonVisualShapeProperties29.Append(applicationNonVisualDrawingProperties48);

            ShapeProperties shapeProperties37 = new ShapeProperties();

            A.Transform2D transform2D32 = new A.Transform2D();
            A.Offset offset43 = new A.Offset() { X = 5183188L, Y = 987425L };
            A.Extents extents43 = new A.Extents() { Cx = 6172200L, Cy = 4873625L };

            transform2D32.Append(offset43);
            transform2D32.Append(extents43);

            shapeProperties37.Append(transform2D32);

            TextBody textBody50 = new TextBody();
            A.BodyProperties bodyProperties50 = new A.BodyProperties();

            A.ListStyle listStyle50 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties11 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties60 = new A.DefaultRunProperties() { FontSize = 3200 };

            level1ParagraphProperties11.Append(defaultRunProperties60);

            A.Level2ParagraphProperties level2ParagraphProperties5 = new A.Level2ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties61 = new A.DefaultRunProperties() { FontSize = 2800 };

            level2ParagraphProperties5.Append(defaultRunProperties61);

            A.Level3ParagraphProperties level3ParagraphProperties5 = new A.Level3ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties62 = new A.DefaultRunProperties() { FontSize = 2400 };

            level3ParagraphProperties5.Append(defaultRunProperties62);

            A.Level4ParagraphProperties level4ParagraphProperties5 = new A.Level4ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties63 = new A.DefaultRunProperties() { FontSize = 2000 };

            level4ParagraphProperties5.Append(defaultRunProperties63);

            A.Level5ParagraphProperties level5ParagraphProperties5 = new A.Level5ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties64 = new A.DefaultRunProperties() { FontSize = 2000 };

            level5ParagraphProperties5.Append(defaultRunProperties64);

            A.Level6ParagraphProperties level6ParagraphProperties5 = new A.Level6ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties65 = new A.DefaultRunProperties() { FontSize = 2000 };

            level6ParagraphProperties5.Append(defaultRunProperties65);

            A.Level7ParagraphProperties level7ParagraphProperties5 = new A.Level7ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties66 = new A.DefaultRunProperties() { FontSize = 2000 };

            level7ParagraphProperties5.Append(defaultRunProperties66);

            A.Level8ParagraphProperties level8ParagraphProperties5 = new A.Level8ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties67 = new A.DefaultRunProperties() { FontSize = 2000 };

            level8ParagraphProperties5.Append(defaultRunProperties67);

            A.Level9ParagraphProperties level9ParagraphProperties5 = new A.Level9ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties68 = new A.DefaultRunProperties() { FontSize = 2000 };

            level9ParagraphProperties5.Append(defaultRunProperties68);

            listStyle50.Append(level1ParagraphProperties11);
            listStyle50.Append(level2ParagraphProperties5);
            listStyle50.Append(level3ParagraphProperties5);
            listStyle50.Append(level4ParagraphProperties5);
            listStyle50.Append(level5ParagraphProperties5);
            listStyle50.Append(level6ParagraphProperties5);
            listStyle50.Append(level7ParagraphProperties5);
            listStyle50.Append(level8ParagraphProperties5);
            listStyle50.Append(level9ParagraphProperties5);

            A.Paragraph paragraph54 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties38 = new A.ParagraphProperties() { Level = 0 };

            A.Run run36 = new A.Run();
            A.RunProperties runProperties40 = new A.RunProperties() { Language = "en-GB" };
            A.Text text40 = new A.Text();
            text40.Text = "Click to edit Master text styles";

            run36.Append(runProperties40);
            run36.Append(text40);

            paragraph54.Append(paragraphProperties38);
            paragraph54.Append(run36);

            A.Paragraph paragraph55 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties39 = new A.ParagraphProperties() { Level = 1 };

            A.Run run37 = new A.Run();
            A.RunProperties runProperties41 = new A.RunProperties() { Language = "en-GB" };
            A.Text text41 = new A.Text();
            text41.Text = "Second level";

            run37.Append(runProperties41);
            run37.Append(text41);

            paragraph55.Append(paragraphProperties39);
            paragraph55.Append(run37);

            A.Paragraph paragraph56 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties40 = new A.ParagraphProperties() { Level = 2 };

            A.Run run38 = new A.Run();
            A.RunProperties runProperties42 = new A.RunProperties() { Language = "en-GB" };
            A.Text text42 = new A.Text();
            text42.Text = "Third level";

            run38.Append(runProperties42);
            run38.Append(text42);

            paragraph56.Append(paragraphProperties40);
            paragraph56.Append(run38);

            A.Paragraph paragraph57 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties41 = new A.ParagraphProperties() { Level = 3 };

            A.Run run39 = new A.Run();
            A.RunProperties runProperties43 = new A.RunProperties() { Language = "en-GB" };
            A.Text text43 = new A.Text();
            text43.Text = "Fourth level";

            run39.Append(runProperties43);
            run39.Append(text43);

            paragraph57.Append(paragraphProperties41);
            paragraph57.Append(run39);

            A.Paragraph paragraph58 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties42 = new A.ParagraphProperties() { Level = 4 };

            A.Run run40 = new A.Run();
            A.RunProperties runProperties44 = new A.RunProperties() { Language = "en-GB" };
            A.Text text44 = new A.Text();
            text44.Text = "Fifth level";

            run40.Append(runProperties44);
            run40.Append(text44);
            A.EndParagraphRunProperties endParagraphRunProperties45 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph58.Append(paragraphProperties42);
            paragraph58.Append(run40);
            paragraph58.Append(endParagraphRunProperties45);

            textBody50.Append(bodyProperties50);
            textBody50.Append(listStyle50);
            textBody50.Append(paragraph54);
            textBody50.Append(paragraph55);
            textBody50.Append(paragraph56);
            textBody50.Append(paragraph57);
            textBody50.Append(paragraph58);

            shape29.Append(nonVisualShapeProperties29);
            shape29.Append(shapeProperties37);
            shape29.Append(textBody50);

            Shape shape30 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties30 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties49 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Text Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList45 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension45 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement65 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{51D187C2-74A9-BCCB-477E-A7910DE87120}\" />");

            // nonVisualDrawingPropertiesExtension45.Append(openXmlUnknownElement65);

            nonVisualDrawingPropertiesExtensionList45.Append(nonVisualDrawingPropertiesExtension45);

            nonVisualDrawingProperties49.Append(nonVisualDrawingPropertiesExtensionList45);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties30 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks15 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties30.Append(shapeLocks15);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties49 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape15 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties49.Append(placeholderShape15);

            nonVisualShapeProperties30.Append(nonVisualDrawingProperties49);
            nonVisualShapeProperties30.Append(nonVisualShapeDrawingProperties30);
            nonVisualShapeProperties30.Append(applicationNonVisualDrawingProperties49);

            ShapeProperties shapeProperties38 = new ShapeProperties();

            A.Transform2D transform2D33 = new A.Transform2D();
            A.Offset offset44 = new A.Offset() { X = 839788L, Y = 2057400L };
            A.Extents extents44 = new A.Extents() { Cx = 3932237L, Cy = 3811588L };

            transform2D33.Append(offset44);
            transform2D33.Append(extents44);

            shapeProperties38.Append(transform2D33);

            TextBody textBody51 = new TextBody();
            A.BodyProperties bodyProperties51 = new A.BodyProperties();

            A.ListStyle listStyle51 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties12 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet25 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties69 = new A.DefaultRunProperties() { FontSize = 1600 };

            level1ParagraphProperties12.Append(noBullet25);
            level1ParagraphProperties12.Append(defaultRunProperties69);

            A.Level2ParagraphProperties level2ParagraphProperties6 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet26 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties70 = new A.DefaultRunProperties() { FontSize = 1400 };

            level2ParagraphProperties6.Append(noBullet26);
            level2ParagraphProperties6.Append(defaultRunProperties70);

            A.Level3ParagraphProperties level3ParagraphProperties6 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet27 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties71 = new A.DefaultRunProperties() { FontSize = 1200 };

            level3ParagraphProperties6.Append(noBullet27);
            level3ParagraphProperties6.Append(defaultRunProperties71);

            A.Level4ParagraphProperties level4ParagraphProperties6 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet28 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties72 = new A.DefaultRunProperties() { FontSize = 1000 };

            level4ParagraphProperties6.Append(noBullet28);
            level4ParagraphProperties6.Append(defaultRunProperties72);

            A.Level5ParagraphProperties level5ParagraphProperties6 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet29 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties73 = new A.DefaultRunProperties() { FontSize = 1000 };

            level5ParagraphProperties6.Append(noBullet29);
            level5ParagraphProperties6.Append(defaultRunProperties73);

            A.Level6ParagraphProperties level6ParagraphProperties6 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet30 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties74 = new A.DefaultRunProperties() { FontSize = 1000 };

            level6ParagraphProperties6.Append(noBullet30);
            level6ParagraphProperties6.Append(defaultRunProperties74);

            A.Level7ParagraphProperties level7ParagraphProperties6 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet31 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties75 = new A.DefaultRunProperties() { FontSize = 1000 };

            level7ParagraphProperties6.Append(noBullet31);
            level7ParagraphProperties6.Append(defaultRunProperties75);

            A.Level8ParagraphProperties level8ParagraphProperties6 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet32 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties76 = new A.DefaultRunProperties() { FontSize = 1000 };

            level8ParagraphProperties6.Append(noBullet32);
            level8ParagraphProperties6.Append(defaultRunProperties76);

            A.Level9ParagraphProperties level9ParagraphProperties6 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet33 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties77 = new A.DefaultRunProperties() { FontSize = 1000 };

            level9ParagraphProperties6.Append(noBullet33);
            level9ParagraphProperties6.Append(defaultRunProperties77);

            listStyle51.Append(level1ParagraphProperties12);
            listStyle51.Append(level2ParagraphProperties6);
            listStyle51.Append(level3ParagraphProperties6);
            listStyle51.Append(level4ParagraphProperties6);
            listStyle51.Append(level5ParagraphProperties6);
            listStyle51.Append(level6ParagraphProperties6);
            listStyle51.Append(level7ParagraphProperties6);
            listStyle51.Append(level8ParagraphProperties6);
            listStyle51.Append(level9ParagraphProperties6);

            A.Paragraph paragraph59 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties43 = new A.ParagraphProperties() { Level = 0 };

            A.Run run41 = new A.Run();
            A.RunProperties runProperties45 = new A.RunProperties() { Language = "en-GB" };
            A.Text text45 = new A.Text();
            text45.Text = "Click to edit Master text styles";

            run41.Append(runProperties45);
            run41.Append(text45);

            paragraph59.Append(paragraphProperties43);
            paragraph59.Append(run41);

            textBody51.Append(bodyProperties51);
            textBody51.Append(listStyle51);
            textBody51.Append(paragraph59);

            shape30.Append(nonVisualShapeProperties30);
            shape30.Append(shapeProperties38);
            shape30.Append(textBody51);

            Shape shape31 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties31 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties50 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Date Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList46 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension46 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement66 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{1E9BAC41-91CD-0D16-C041-855E0E040896}\" />");

            // nonVisualDrawingPropertiesExtension46.Append(openXmlUnknownElement66);

            nonVisualDrawingPropertiesExtensionList46.Append(nonVisualDrawingPropertiesExtension46);

            nonVisualDrawingProperties50.Append(nonVisualDrawingPropertiesExtensionList46);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties31 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks16 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties31.Append(shapeLocks16);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties50 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape16 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties50.Append(placeholderShape16);

            nonVisualShapeProperties31.Append(nonVisualDrawingProperties50);
            nonVisualShapeProperties31.Append(nonVisualShapeDrawingProperties31);
            nonVisualShapeProperties31.Append(applicationNonVisualDrawingProperties50);
            ShapeProperties shapeProperties39 = new ShapeProperties();

            TextBody textBody52 = new TextBody();
            A.BodyProperties bodyProperties52 = new A.BodyProperties();
            A.ListStyle listStyle52 = new A.ListStyle();

            A.Paragraph paragraph60 = new A.Paragraph();

            A.Field field5 = new A.Field() { Id = "{C453C8C6-607B-FF49-A93F-867C282599C0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties46 = new A.RunProperties() { Language = "en-US" };
            runProperties46.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text46 = new A.Text();
            text46.Text = "5/3/2024";

            field5.Append(runProperties46);
            field5.Append(text46);
            A.EndParagraphRunProperties endParagraphRunProperties46 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph60.Append(field5);
            paragraph60.Append(endParagraphRunProperties46);

            textBody52.Append(bodyProperties52);
            textBody52.Append(listStyle52);
            textBody52.Append(paragraph60);

            shape31.Append(nonVisualShapeProperties31);
            shape31.Append(shapeProperties39);
            shape31.Append(textBody52);

            Shape shape32 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties32 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties51 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Footer Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList47 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension47 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement67 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{0DE746CC-56D2-716B-330F-8090E22F9518}\" />");

            // nonVisualDrawingPropertiesExtension47.Append(openXmlUnknownElement67);

            nonVisualDrawingPropertiesExtensionList47.Append(nonVisualDrawingPropertiesExtension47);

            nonVisualDrawingProperties51.Append(nonVisualDrawingPropertiesExtensionList47);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties32 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks17 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties32.Append(shapeLocks17);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties51 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape17 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties51.Append(placeholderShape17);

            nonVisualShapeProperties32.Append(nonVisualDrawingProperties51);
            nonVisualShapeProperties32.Append(nonVisualShapeDrawingProperties32);
            nonVisualShapeProperties32.Append(applicationNonVisualDrawingProperties51);
            ShapeProperties shapeProperties40 = new ShapeProperties();

            TextBody textBody53 = new TextBody();
            A.BodyProperties bodyProperties53 = new A.BodyProperties();
            A.ListStyle listStyle53 = new A.ListStyle();

            A.Paragraph paragraph61 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties47 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph61.Append(endParagraphRunProperties47);

            textBody53.Append(bodyProperties53);
            textBody53.Append(listStyle53);
            textBody53.Append(paragraph61);

            shape32.Append(nonVisualShapeProperties32);
            shape32.Append(shapeProperties40);
            shape32.Append(textBody53);

            Shape shape33 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties33 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties52 = new NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Slide Number Placeholder 6" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList48 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension48 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement68 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{830DBDDF-2E07-9B57-53DB-8C393B607078}\" />");

            // nonVisualDrawingPropertiesExtension48.Append(openXmlUnknownElement68);

            nonVisualDrawingPropertiesExtensionList48.Append(nonVisualDrawingPropertiesExtension48);

            nonVisualDrawingProperties52.Append(nonVisualDrawingPropertiesExtensionList48);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties33 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks18 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties33.Append(shapeLocks18);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties52 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape18 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties52.Append(placeholderShape18);

            nonVisualShapeProperties33.Append(nonVisualDrawingProperties52);
            nonVisualShapeProperties33.Append(nonVisualShapeDrawingProperties33);
            nonVisualShapeProperties33.Append(applicationNonVisualDrawingProperties52);
            ShapeProperties shapeProperties41 = new ShapeProperties();

            TextBody textBody54 = new TextBody();
            A.BodyProperties bodyProperties54 = new A.BodyProperties();
            A.ListStyle listStyle54 = new A.ListStyle();

            A.Paragraph paragraph62 = new A.Paragraph();

            A.Field field6 = new A.Field() { Id = "{402C1264-1A6A-2D47-B539-28043A88F412}", Type = "slidenum" };

            A.RunProperties runProperties47 = new A.RunProperties() { Language = "en-US" };
            runProperties47.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text47 = new A.Text();
            text47.Text = "‹#›";

            field6.Append(runProperties47);
            field6.Append(text47);
            A.EndParagraphRunProperties endParagraphRunProperties48 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph62.Append(field6);
            paragraph62.Append(endParagraphRunProperties48);

            textBody54.Append(bodyProperties54);
            textBody54.Append(listStyle54);
            textBody54.Append(paragraph62);

            shape33.Append(nonVisualShapeProperties33);
            shape33.Append(shapeProperties41);
            shape33.Append(textBody54);

            shapeTree4.Append(nonVisualGroupShapeProperties10);
            shapeTree4.Append(groupShapeProperties10);
            shapeTree4.Append(shape28);
            shapeTree4.Append(shape29);
            shapeTree4.Append(shape30);
            shapeTree4.Append(shape31);
            shapeTree4.Append(shape32);
            shapeTree4.Append(shape33);

            CommonSlideDataExtensionList commonSlideDataExtensionList4 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension4 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId4 = new P14.CreationId() { Val = (UInt32Value)3212331070U };
            creationId4.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension4.Append(creationId4);

            commonSlideDataExtensionList4.Append(commonSlideDataExtension4);

            commonSlideData4.Append(shapeTree4);
            commonSlideData4.Append(commonSlideDataExtensionList4);

            ColorMapOverride colorMapOverride3 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping3 = new A.MasterColorMapping();

            colorMapOverride3.Append(masterColorMapping3);

            slideLayout2.Append(commonSlideData4);
            slideLayout2.Append(colorMapOverride3);

            slideLayoutPart2.SlideLayout = slideLayout2;
        }

        // Generates content of slideLayoutPart3.
        private void GenerateSlideLayoutPart3Content(SlideLayoutPart slideLayoutPart3)
        {
            SlideLayout slideLayout3 = new SlideLayout() { Type = SlideLayoutValues.SectionHeader, Preserve = true };
            slideLayout3.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout3.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData5 = new CommonSlideData() { Name = "Section Header" };

            ShapeTree shapeTree5 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties11 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties53 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties11 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties53 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties11.Append(nonVisualDrawingProperties53);
            nonVisualGroupShapeProperties11.Append(nonVisualGroupShapeDrawingProperties11);
            nonVisualGroupShapeProperties11.Append(applicationNonVisualDrawingProperties53);

            GroupShapeProperties groupShapeProperties11 = new GroupShapeProperties();

            A.TransformGroup transformGroup11 = new A.TransformGroup();
            A.Offset offset45 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents45 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset11 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents11 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup11.Append(offset45);
            transformGroup11.Append(extents45);
            transformGroup11.Append(childOffset11);
            transformGroup11.Append(childExtents11);

            groupShapeProperties11.Append(transformGroup11);

            Shape shape34 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties34 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties54 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList49 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension49 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement69 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{F43ECC81-88F5-9C6D-69EB-793ADB3A252D}\" />");

            // nonVisualDrawingPropertiesExtension49.Append(openXmlUnknownElement69);

            nonVisualDrawingPropertiesExtensionList49.Append(nonVisualDrawingPropertiesExtension49);

            nonVisualDrawingProperties54.Append(nonVisualDrawingPropertiesExtensionList49);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties34 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks19 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties34.Append(shapeLocks19);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties54 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape19 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties54.Append(placeholderShape19);

            nonVisualShapeProperties34.Append(nonVisualDrawingProperties54);
            nonVisualShapeProperties34.Append(nonVisualShapeDrawingProperties34);
            nonVisualShapeProperties34.Append(applicationNonVisualDrawingProperties54);

            ShapeProperties shapeProperties42 = new ShapeProperties();

            A.Transform2D transform2D34 = new A.Transform2D();
            A.Offset offset46 = new A.Offset() { X = 831850L, Y = 1709738L };
            A.Extents extents46 = new A.Extents() { Cx = 10515600L, Cy = 2852737L };

            transform2D34.Append(offset46);
            transform2D34.Append(extents46);

            shapeProperties42.Append(transform2D34);

            TextBody textBody55 = new TextBody();
            A.BodyProperties bodyProperties55 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle55 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties13 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties78 = new A.DefaultRunProperties() { FontSize = 6000 };

            level1ParagraphProperties13.Append(defaultRunProperties78);

            listStyle55.Append(level1ParagraphProperties13);

            A.Paragraph paragraph63 = new A.Paragraph();

            A.Run run42 = new A.Run();
            A.RunProperties runProperties48 = new A.RunProperties() { Language = "en-GB" };
            A.Text text48 = new A.Text();
            text48.Text = "Click to edit Master title style";

            run42.Append(runProperties48);
            run42.Append(text48);
            A.EndParagraphRunProperties endParagraphRunProperties49 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph63.Append(run42);
            paragraph63.Append(endParagraphRunProperties49);

            textBody55.Append(bodyProperties55);
            textBody55.Append(listStyle55);
            textBody55.Append(paragraph63);

            shape34.Append(nonVisualShapeProperties34);
            shape34.Append(shapeProperties42);
            shape34.Append(textBody55);

            Shape shape35 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties35 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties55 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Text Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList50 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension50 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement70 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{1024183D-3429-1DC4-B72C-60E150233AB4}\" />");

            // nonVisualDrawingPropertiesExtension50.Append(openXmlUnknownElement70);

            nonVisualDrawingPropertiesExtensionList50.Append(nonVisualDrawingPropertiesExtension50);

            nonVisualDrawingProperties55.Append(nonVisualDrawingPropertiesExtensionList50);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties35 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks20 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties35.Append(shapeLocks20);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties55 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape20 = new PlaceholderShape() { Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties55.Append(placeholderShape20);

            nonVisualShapeProperties35.Append(nonVisualDrawingProperties55);
            nonVisualShapeProperties35.Append(nonVisualShapeDrawingProperties35);
            nonVisualShapeProperties35.Append(applicationNonVisualDrawingProperties55);

            ShapeProperties shapeProperties43 = new ShapeProperties();

            A.Transform2D transform2D35 = new A.Transform2D();
            A.Offset offset47 = new A.Offset() { X = 831850L, Y = 4589463L };
            A.Extents extents47 = new A.Extents() { Cx = 10515600L, Cy = 1500187L };

            transform2D35.Append(offset47);
            transform2D35.Append(extents47);

            shapeProperties43.Append(transform2D35);

            TextBody textBody56 = new TextBody();
            A.BodyProperties bodyProperties56 = new A.BodyProperties();

            A.ListStyle listStyle56 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties14 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet34 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties79 = new A.DefaultRunProperties() { FontSize = 2400 };

            A.SolidFill solidFill89 = new A.SolidFill();

            A.SchemeColor schemeColor57 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint4 = new A.Tint() { Val = 82000 };

            schemeColor57.Append(tint4);

            solidFill89.Append(schemeColor57);

            defaultRunProperties79.Append(solidFill89);

            level1ParagraphProperties14.Append(noBullet34);
            level1ParagraphProperties14.Append(defaultRunProperties79);

            A.Level2ParagraphProperties level2ParagraphProperties7 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet35 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties80 = new A.DefaultRunProperties() { FontSize = 2000 };

            A.SolidFill solidFill90 = new A.SolidFill();

            A.SchemeColor schemeColor58 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint5 = new A.Tint() { Val = 82000 };

            schemeColor58.Append(tint5);

            solidFill90.Append(schemeColor58);

            defaultRunProperties80.Append(solidFill90);

            level2ParagraphProperties7.Append(noBullet35);
            level2ParagraphProperties7.Append(defaultRunProperties80);

            A.Level3ParagraphProperties level3ParagraphProperties7 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet36 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties81 = new A.DefaultRunProperties() { FontSize = 1800 };

            A.SolidFill solidFill91 = new A.SolidFill();

            A.SchemeColor schemeColor59 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint6 = new A.Tint() { Val = 82000 };

            schemeColor59.Append(tint6);

            solidFill91.Append(schemeColor59);

            defaultRunProperties81.Append(solidFill91);

            level3ParagraphProperties7.Append(noBullet36);
            level3ParagraphProperties7.Append(defaultRunProperties81);

            A.Level4ParagraphProperties level4ParagraphProperties7 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet37 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties82 = new A.DefaultRunProperties() { FontSize = 1600 };

            A.SolidFill solidFill92 = new A.SolidFill();

            A.SchemeColor schemeColor60 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint7 = new A.Tint() { Val = 82000 };

            schemeColor60.Append(tint7);

            solidFill92.Append(schemeColor60);

            defaultRunProperties82.Append(solidFill92);

            level4ParagraphProperties7.Append(noBullet37);
            level4ParagraphProperties7.Append(defaultRunProperties82);

            A.Level5ParagraphProperties level5ParagraphProperties7 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet38 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties83 = new A.DefaultRunProperties() { FontSize = 1600 };

            A.SolidFill solidFill93 = new A.SolidFill();

            A.SchemeColor schemeColor61 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint8 = new A.Tint() { Val = 82000 };

            schemeColor61.Append(tint8);

            solidFill93.Append(schemeColor61);

            defaultRunProperties83.Append(solidFill93);

            level5ParagraphProperties7.Append(noBullet38);
            level5ParagraphProperties7.Append(defaultRunProperties83);

            A.Level6ParagraphProperties level6ParagraphProperties7 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet39 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties84 = new A.DefaultRunProperties() { FontSize = 1600 };

            A.SolidFill solidFill94 = new A.SolidFill();

            A.SchemeColor schemeColor62 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint9 = new A.Tint() { Val = 82000 };

            schemeColor62.Append(tint9);

            solidFill94.Append(schemeColor62);

            defaultRunProperties84.Append(solidFill94);

            level6ParagraphProperties7.Append(noBullet39);
            level6ParagraphProperties7.Append(defaultRunProperties84);

            A.Level7ParagraphProperties level7ParagraphProperties7 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet40 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties85 = new A.DefaultRunProperties() { FontSize = 1600 };

            A.SolidFill solidFill95 = new A.SolidFill();

            A.SchemeColor schemeColor63 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint10 = new A.Tint() { Val = 82000 };

            schemeColor63.Append(tint10);

            solidFill95.Append(schemeColor63);

            defaultRunProperties85.Append(solidFill95);

            level7ParagraphProperties7.Append(noBullet40);
            level7ParagraphProperties7.Append(defaultRunProperties85);

            A.Level8ParagraphProperties level8ParagraphProperties7 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet41 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties86 = new A.DefaultRunProperties() { FontSize = 1600 };

            A.SolidFill solidFill96 = new A.SolidFill();

            A.SchemeColor schemeColor64 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint11 = new A.Tint() { Val = 82000 };

            schemeColor64.Append(tint11);

            solidFill96.Append(schemeColor64);

            defaultRunProperties86.Append(solidFill96);

            level8ParagraphProperties7.Append(noBullet41);
            level8ParagraphProperties7.Append(defaultRunProperties86);

            A.Level9ParagraphProperties level9ParagraphProperties7 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet42 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties87 = new A.DefaultRunProperties() { FontSize = 1600 };

            A.SolidFill solidFill97 = new A.SolidFill();

            A.SchemeColor schemeColor65 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint12 = new A.Tint() { Val = 82000 };

            schemeColor65.Append(tint12);

            solidFill97.Append(schemeColor65);

            defaultRunProperties87.Append(solidFill97);

            level9ParagraphProperties7.Append(noBullet42);
            level9ParagraphProperties7.Append(defaultRunProperties87);

            listStyle56.Append(level1ParagraphProperties14);
            listStyle56.Append(level2ParagraphProperties7);
            listStyle56.Append(level3ParagraphProperties7);
            listStyle56.Append(level4ParagraphProperties7);
            listStyle56.Append(level5ParagraphProperties7);
            listStyle56.Append(level6ParagraphProperties7);
            listStyle56.Append(level7ParagraphProperties7);
            listStyle56.Append(level8ParagraphProperties7);
            listStyle56.Append(level9ParagraphProperties7);

            A.Paragraph paragraph64 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties44 = new A.ParagraphProperties() { Level = 0 };

            A.Run run43 = new A.Run();
            A.RunProperties runProperties49 = new A.RunProperties() { Language = "en-GB" };
            A.Text text49 = new A.Text();
            text49.Text = "Click to edit Master text styles";

            run43.Append(runProperties49);
            run43.Append(text49);

            paragraph64.Append(paragraphProperties44);
            paragraph64.Append(run43);

            textBody56.Append(bodyProperties56);
            textBody56.Append(listStyle56);
            textBody56.Append(paragraph64);

            shape35.Append(nonVisualShapeProperties35);
            shape35.Append(shapeProperties43);
            shape35.Append(textBody56);

            Shape shape36 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties36 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties56 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList51 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension51 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement71 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{687EA223-A7D3-63C2-9B62-D3B5CEDC3126}\" />");

            // nonVisualDrawingPropertiesExtension51.Append(openXmlUnknownElement71);

            nonVisualDrawingPropertiesExtensionList51.Append(nonVisualDrawingPropertiesExtension51);

            nonVisualDrawingProperties56.Append(nonVisualDrawingPropertiesExtensionList51);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties36 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks21 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties36.Append(shapeLocks21);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties56 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape21 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties56.Append(placeholderShape21);

            nonVisualShapeProperties36.Append(nonVisualDrawingProperties56);
            nonVisualShapeProperties36.Append(nonVisualShapeDrawingProperties36);
            nonVisualShapeProperties36.Append(applicationNonVisualDrawingProperties56);
            ShapeProperties shapeProperties44 = new ShapeProperties();

            TextBody textBody57 = new TextBody();
            A.BodyProperties bodyProperties57 = new A.BodyProperties();
            A.ListStyle listStyle57 = new A.ListStyle();

            A.Paragraph paragraph65 = new A.Paragraph();

            A.Field field7 = new A.Field() { Id = "{C453C8C6-607B-FF49-A93F-867C282599C0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties50 = new A.RunProperties() { Language = "en-US" };
            runProperties50.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text50 = new A.Text();
            text50.Text = "5/3/2024";

            field7.Append(runProperties50);
            field7.Append(text50);
            A.EndParagraphRunProperties endParagraphRunProperties50 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph65.Append(field7);
            paragraph65.Append(endParagraphRunProperties50);

            textBody57.Append(bodyProperties57);
            textBody57.Append(listStyle57);
            textBody57.Append(paragraph65);

            shape36.Append(nonVisualShapeProperties36);
            shape36.Append(shapeProperties44);
            shape36.Append(textBody57);

            Shape shape37 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties37 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties57 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList52 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension52 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement72 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{E5C159AA-14AF-849F-C907-E47EB2B5BF5A}\" />");

            // nonVisualDrawingPropertiesExtension52.Append(openXmlUnknownElement72);

            nonVisualDrawingPropertiesExtensionList52.Append(nonVisualDrawingPropertiesExtension52);

            nonVisualDrawingProperties57.Append(nonVisualDrawingPropertiesExtensionList52);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties37 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks22 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties37.Append(shapeLocks22);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties57 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape22 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties57.Append(placeholderShape22);

            nonVisualShapeProperties37.Append(nonVisualDrawingProperties57);
            nonVisualShapeProperties37.Append(nonVisualShapeDrawingProperties37);
            nonVisualShapeProperties37.Append(applicationNonVisualDrawingProperties57);
            ShapeProperties shapeProperties45 = new ShapeProperties();

            TextBody textBody58 = new TextBody();
            A.BodyProperties bodyProperties58 = new A.BodyProperties();
            A.ListStyle listStyle58 = new A.ListStyle();

            A.Paragraph paragraph66 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties51 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph66.Append(endParagraphRunProperties51);

            textBody58.Append(bodyProperties58);
            textBody58.Append(listStyle58);
            textBody58.Append(paragraph66);

            shape37.Append(nonVisualShapeProperties37);
            shape37.Append(shapeProperties45);
            shape37.Append(textBody58);

            Shape shape38 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties38 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties58 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList53 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension53 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement73 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{6EAC116F-DAE1-3034-22C9-98DFA2702DB0}\" />");

            // nonVisualDrawingPropertiesExtension53.Append(openXmlUnknownElement73);

            nonVisualDrawingPropertiesExtensionList53.Append(nonVisualDrawingPropertiesExtension53);

            nonVisualDrawingProperties58.Append(nonVisualDrawingPropertiesExtensionList53);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties38 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks23 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties38.Append(shapeLocks23);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties58 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape23 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties58.Append(placeholderShape23);

            nonVisualShapeProperties38.Append(nonVisualDrawingProperties58);
            nonVisualShapeProperties38.Append(nonVisualShapeDrawingProperties38);
            nonVisualShapeProperties38.Append(applicationNonVisualDrawingProperties58);
            ShapeProperties shapeProperties46 = new ShapeProperties();

            TextBody textBody59 = new TextBody();
            A.BodyProperties bodyProperties59 = new A.BodyProperties();
            A.ListStyle listStyle59 = new A.ListStyle();

            A.Paragraph paragraph67 = new A.Paragraph();

            A.Field field8 = new A.Field() { Id = "{402C1264-1A6A-2D47-B539-28043A88F412}", Type = "slidenum" };

            A.RunProperties runProperties51 = new A.RunProperties() { Language = "en-US" };
            runProperties51.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text51 = new A.Text();
            text51.Text = "‹#›";

            field8.Append(runProperties51);
            field8.Append(text51);
            A.EndParagraphRunProperties endParagraphRunProperties52 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph67.Append(field8);
            paragraph67.Append(endParagraphRunProperties52);

            textBody59.Append(bodyProperties59);
            textBody59.Append(listStyle59);
            textBody59.Append(paragraph67);

            shape38.Append(nonVisualShapeProperties38);
            shape38.Append(shapeProperties46);
            shape38.Append(textBody59);

            shapeTree5.Append(nonVisualGroupShapeProperties11);
            shapeTree5.Append(groupShapeProperties11);
            shapeTree5.Append(shape34);
            shapeTree5.Append(shape35);
            shapeTree5.Append(shape36);
            shapeTree5.Append(shape37);
            shapeTree5.Append(shape38);

            CommonSlideDataExtensionList commonSlideDataExtensionList5 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension5 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId5 = new P14.CreationId() { Val = (UInt32Value)220367687U };
            creationId5.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension5.Append(creationId5);

            commonSlideDataExtensionList5.Append(commonSlideDataExtension5);

            commonSlideData5.Append(shapeTree5);
            commonSlideData5.Append(commonSlideDataExtensionList5);

            ColorMapOverride colorMapOverride4 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping4 = new A.MasterColorMapping();

            colorMapOverride4.Append(masterColorMapping4);

            slideLayout3.Append(commonSlideData5);
            slideLayout3.Append(colorMapOverride4);

            slideLayoutPart3.SlideLayout = slideLayout3;
        }

        // Generates content of slideLayoutPart4.
        private void GenerateSlideLayoutPart4Content(SlideLayoutPart slideLayoutPart4)
        {
            SlideLayout slideLayout4 = new SlideLayout() { Type = SlideLayoutValues.Blank, Preserve = true };
            slideLayout4.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout4.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout4.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData6 = new CommonSlideData() { Name = "Blank" };

            ShapeTree shapeTree6 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties12 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties59 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties12 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties59 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties12.Append(nonVisualDrawingProperties59);
            nonVisualGroupShapeProperties12.Append(nonVisualGroupShapeDrawingProperties12);
            nonVisualGroupShapeProperties12.Append(applicationNonVisualDrawingProperties59);

            GroupShapeProperties groupShapeProperties12 = new GroupShapeProperties();

            A.TransformGroup transformGroup12 = new A.TransformGroup();
            A.Offset offset48 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents48 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset12 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents12 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup12.Append(offset48);
            transformGroup12.Append(extents48);
            transformGroup12.Append(childOffset12);
            transformGroup12.Append(childExtents12);

            groupShapeProperties12.Append(transformGroup12);

            Shape shape39 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties39 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties60 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Date Placeholder 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList54 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension54 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement74 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{94596E9D-A840-39F0-A920-FD8A12022D67}\" />");

            // nonVisualDrawingPropertiesExtension54.Append(openXmlUnknownElement74);

            nonVisualDrawingPropertiesExtensionList54.Append(nonVisualDrawingPropertiesExtension54);

            nonVisualDrawingProperties60.Append(nonVisualDrawingPropertiesExtensionList54);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties39 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks24 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties39.Append(shapeLocks24);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties60 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape24 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties60.Append(placeholderShape24);

            nonVisualShapeProperties39.Append(nonVisualDrawingProperties60);
            nonVisualShapeProperties39.Append(nonVisualShapeDrawingProperties39);
            nonVisualShapeProperties39.Append(applicationNonVisualDrawingProperties60);
            ShapeProperties shapeProperties47 = new ShapeProperties();

            TextBody textBody60 = new TextBody();
            A.BodyProperties bodyProperties60 = new A.BodyProperties();
            A.ListStyle listStyle60 = new A.ListStyle();

            A.Paragraph paragraph68 = new A.Paragraph();

            A.Field field9 = new A.Field() { Id = "{C453C8C6-607B-FF49-A93F-867C282599C0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties52 = new A.RunProperties() { Language = "en-US" };
            runProperties52.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text52 = new A.Text();
            text52.Text = "5/3/2024";

            field9.Append(runProperties52);
            field9.Append(text52);
            A.EndParagraphRunProperties endParagraphRunProperties53 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph68.Append(field9);
            paragraph68.Append(endParagraphRunProperties53);

            textBody60.Append(bodyProperties60);
            textBody60.Append(listStyle60);
            textBody60.Append(paragraph68);

            shape39.Append(nonVisualShapeProperties39);
            shape39.Append(shapeProperties47);
            shape39.Append(textBody60);

            Shape shape40 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties40 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties61 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Footer Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList55 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension55 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement75 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{985F9389-C88B-7A86-0AFB-EC52D7EFF927}\" />");

            // nonVisualDrawingPropertiesExtension55.Append(openXmlUnknownElement75);

            nonVisualDrawingPropertiesExtensionList55.Append(nonVisualDrawingPropertiesExtension55);

            nonVisualDrawingProperties61.Append(nonVisualDrawingPropertiesExtensionList55);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties40 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks25 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties40.Append(shapeLocks25);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties61 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape25 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties61.Append(placeholderShape25);

            nonVisualShapeProperties40.Append(nonVisualDrawingProperties61);
            nonVisualShapeProperties40.Append(nonVisualShapeDrawingProperties40);
            nonVisualShapeProperties40.Append(applicationNonVisualDrawingProperties61);
            ShapeProperties shapeProperties48 = new ShapeProperties();

            TextBody textBody61 = new TextBody();
            A.BodyProperties bodyProperties61 = new A.BodyProperties();
            A.ListStyle listStyle61 = new A.ListStyle();

            A.Paragraph paragraph69 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties54 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph69.Append(endParagraphRunProperties54);

            textBody61.Append(bodyProperties61);
            textBody61.Append(listStyle61);
            textBody61.Append(paragraph69);

            shape40.Append(nonVisualShapeProperties40);
            shape40.Append(shapeProperties48);
            shape40.Append(textBody61);

            Shape shape41 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties41 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties62 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Slide Number Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList56 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension56 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement76 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{C537D283-CF1D-7C29-940C-6C5BCC634848}\" />");

            // nonVisualDrawingPropertiesExtension56.Append(openXmlUnknownElement76);

            nonVisualDrawingPropertiesExtensionList56.Append(nonVisualDrawingPropertiesExtension56);

            nonVisualDrawingProperties62.Append(nonVisualDrawingPropertiesExtensionList56);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties41 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks26 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties41.Append(shapeLocks26);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties62 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape26 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties62.Append(placeholderShape26);

            nonVisualShapeProperties41.Append(nonVisualDrawingProperties62);
            nonVisualShapeProperties41.Append(nonVisualShapeDrawingProperties41);
            nonVisualShapeProperties41.Append(applicationNonVisualDrawingProperties62);
            ShapeProperties shapeProperties49 = new ShapeProperties();

            TextBody textBody62 = new TextBody();
            A.BodyProperties bodyProperties62 = new A.BodyProperties();
            A.ListStyle listStyle62 = new A.ListStyle();

            A.Paragraph paragraph70 = new A.Paragraph();

            A.Field field10 = new A.Field() { Id = "{402C1264-1A6A-2D47-B539-28043A88F412}", Type = "slidenum" };

            A.RunProperties runProperties53 = new A.RunProperties() { Language = "en-US" };
            runProperties53.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text53 = new A.Text();
            text53.Text = "‹#›";

            field10.Append(runProperties53);
            field10.Append(text53);
            A.EndParagraphRunProperties endParagraphRunProperties55 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph70.Append(field10);
            paragraph70.Append(endParagraphRunProperties55);

            textBody62.Append(bodyProperties62);
            textBody62.Append(listStyle62);
            textBody62.Append(paragraph70);

            shape41.Append(nonVisualShapeProperties41);
            shape41.Append(shapeProperties49);
            shape41.Append(textBody62);

            shapeTree6.Append(nonVisualGroupShapeProperties12);
            shapeTree6.Append(groupShapeProperties12);
            shapeTree6.Append(shape39);
            shapeTree6.Append(shape40);
            shapeTree6.Append(shape41);

            CommonSlideDataExtensionList commonSlideDataExtensionList6 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension6 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId6 = new P14.CreationId() { Val = (UInt32Value)839272568U };
            creationId6.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension6.Append(creationId6);

            commonSlideDataExtensionList6.Append(commonSlideDataExtension6);

            commonSlideData6.Append(shapeTree6);
            commonSlideData6.Append(commonSlideDataExtensionList6);

            ColorMapOverride colorMapOverride5 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping5 = new A.MasterColorMapping();

            colorMapOverride5.Append(masterColorMapping5);

            slideLayout4.Append(commonSlideData6);
            slideLayout4.Append(colorMapOverride5);

            slideLayoutPart4.SlideLayout = slideLayout4;
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex27 = new A.RgbColorModelHex() { Val = "0E2841" };

            dark2Color1.Append(rgbColorModelHex27);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex28 = new A.RgbColorModelHex() { Val = "E8E8E8" };

            light2Color1.Append(rgbColorModelHex28);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex29 = new A.RgbColorModelHex() { Val = "156082" };

            accent1Color1.Append(rgbColorModelHex29);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex30 = new A.RgbColorModelHex() { Val = "E97132" };

            accent2Color1.Append(rgbColorModelHex30);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex31 = new A.RgbColorModelHex() { Val = "196B24" };

            accent3Color1.Append(rgbColorModelHex31);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex32 = new A.RgbColorModelHex() { Val = "0F9ED5" };

            accent4Color1.Append(rgbColorModelHex32);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex33 = new A.RgbColorModelHex() { Val = "A02B93" };

            accent5Color1.Append(rgbColorModelHex33);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex34 = new A.RgbColorModelHex() { Val = "4EA72E" };

            accent6Color1.Append(rgbColorModelHex34);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex35 = new A.RgbColorModelHex() { Val = "467886" };

            hyperlink1.Append(rgbColorModelHex35);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex36 = new A.RgbColorModelHex() { Val = "96607D" };

            followedHyperlinkColor1.Append(rgbColorModelHex36);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont77 = new A.LatinFont() { Typeface = "Aptos Display", Panose = "02110004020202020204" };
            A.EastAsianFont eastAsianFont61 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont77 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "游ゴシック Light" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "等线 Light" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Angsana New" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Armn", Typeface = "Arial" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Bugi", Typeface = "Leelawadee UI" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Bopo", Typeface = "Microsoft JhengHei" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Java", Typeface = "Javanese Text" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Lisu", Typeface = "Segoe UI" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Mymr", Typeface = "Myanmar Text" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Nkoo", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Olck", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Osma", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Phag", Typeface = "Phagspa" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Syrn", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Syrj", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Syre", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Sora", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Tale", Typeface = "Microsoft Tai Le" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Talu", Typeface = "Microsoft New Tai Lue" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tfng", Typeface = "Ebrima" };

            majorFont1.Append(latinFont77);
            majorFont1.Append(eastAsianFont61);
            majorFont1.Append(complexScriptFont77);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);
            majorFont1.Append(supplementalFont31);
            majorFont1.Append(supplementalFont32);
            majorFont1.Append(supplementalFont33);
            majorFont1.Append(supplementalFont34);
            majorFont1.Append(supplementalFont35);
            majorFont1.Append(supplementalFont36);
            majorFont1.Append(supplementalFont37);
            majorFont1.Append(supplementalFont38);
            majorFont1.Append(supplementalFont39);
            majorFont1.Append(supplementalFont40);
            majorFont1.Append(supplementalFont41);
            majorFont1.Append(supplementalFont42);
            majorFont1.Append(supplementalFont43);
            majorFont1.Append(supplementalFont44);
            majorFont1.Append(supplementalFont45);
            majorFont1.Append(supplementalFont46);
            majorFont1.Append(supplementalFont47);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont78 = new A.LatinFont() { Typeface = "Aptos", Panose = "02110004020202020204" };
            A.EastAsianFont eastAsianFont62 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont78 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Jpan", Typeface = "游ゴシック" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Hans", Typeface = "等线" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Thai", Typeface = "Cordia New" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont61 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont62 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont63 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont64 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont65 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont66 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont67 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont68 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont69 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont70 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont71 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont72 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont73 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont74 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont75 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont76 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont77 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };
            A.SupplementalFont supplementalFont78 = new A.SupplementalFont() { Script = "Armn", Typeface = "Arial" };
            A.SupplementalFont supplementalFont79 = new A.SupplementalFont() { Script = "Bugi", Typeface = "Leelawadee UI" };
            A.SupplementalFont supplementalFont80 = new A.SupplementalFont() { Script = "Bopo", Typeface = "Microsoft JhengHei" };
            A.SupplementalFont supplementalFont81 = new A.SupplementalFont() { Script = "Java", Typeface = "Javanese Text" };
            A.SupplementalFont supplementalFont82 = new A.SupplementalFont() { Script = "Lisu", Typeface = "Segoe UI" };
            A.SupplementalFont supplementalFont83 = new A.SupplementalFont() { Script = "Mymr", Typeface = "Myanmar Text" };
            A.SupplementalFont supplementalFont84 = new A.SupplementalFont() { Script = "Nkoo", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont85 = new A.SupplementalFont() { Script = "Olck", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont86 = new A.SupplementalFont() { Script = "Osma", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont87 = new A.SupplementalFont() { Script = "Phag", Typeface = "Phagspa" };
            A.SupplementalFont supplementalFont88 = new A.SupplementalFont() { Script = "Syrn", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont89 = new A.SupplementalFont() { Script = "Syrj", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont90 = new A.SupplementalFont() { Script = "Syre", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont91 = new A.SupplementalFont() { Script = "Sora", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont92 = new A.SupplementalFont() { Script = "Tale", Typeface = "Microsoft Tai Le" };
            A.SupplementalFont supplementalFont93 = new A.SupplementalFont() { Script = "Talu", Typeface = "Microsoft New Tai Lue" };
            A.SupplementalFont supplementalFont94 = new A.SupplementalFont() { Script = "Tfng", Typeface = "Ebrima" };

            minorFont1.Append(latinFont78);
            minorFont1.Append(eastAsianFont62);
            minorFont1.Append(complexScriptFont78);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);
            minorFont1.Append(supplementalFont61);
            minorFont1.Append(supplementalFont62);
            minorFont1.Append(supplementalFont63);
            minorFont1.Append(supplementalFont64);
            minorFont1.Append(supplementalFont65);
            minorFont1.Append(supplementalFont66);
            minorFont1.Append(supplementalFont67);
            minorFont1.Append(supplementalFont68);
            minorFont1.Append(supplementalFont69);
            minorFont1.Append(supplementalFont70);
            minorFont1.Append(supplementalFont71);
            minorFont1.Append(supplementalFont72);
            minorFont1.Append(supplementalFont73);
            minorFont1.Append(supplementalFont74);
            minorFont1.Append(supplementalFont75);
            minorFont1.Append(supplementalFont76);
            minorFont1.Append(supplementalFont77);
            minorFont1.Append(supplementalFont78);
            minorFont1.Append(supplementalFont79);
            minorFont1.Append(supplementalFont80);
            minorFont1.Append(supplementalFont81);
            minorFont1.Append(supplementalFont82);
            minorFont1.Append(supplementalFont83);
            minorFont1.Append(supplementalFont84);
            minorFont1.Append(supplementalFont85);
            minorFont1.Append(supplementalFont86);
            minorFont1.Append(supplementalFont87);
            minorFont1.Append(supplementalFont88);
            minorFont1.Append(supplementalFont89);
            minorFont1.Append(supplementalFont90);
            minorFont1.Append(supplementalFont91);
            minorFont1.Append(supplementalFont92);
            minorFont1.Append(supplementalFont93);
            minorFont1.Append(supplementalFont94);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill98 = new A.SolidFill();
            A.SchemeColor schemeColor66 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill98.Append(schemeColor66);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor67 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 105000 };
            A.Tint tint13 = new A.Tint() { Val = 67000 };

            schemeColor67.Append(luminanceModulation4);
            schemeColor67.Append(saturationModulation1);
            schemeColor67.Append(tint13);

            gradientStop1.Append(schemeColor67);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor68 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 103000 };
            A.Tint tint14 = new A.Tint() { Val = 73000 };

            schemeColor68.Append(luminanceModulation5);
            schemeColor68.Append(saturationModulation2);
            schemeColor68.Append(tint14);

            gradientStop2.Append(schemeColor68);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor69 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 109000 };
            A.Tint tint15 = new A.Tint() { Val = 81000 };

            schemeColor69.Append(luminanceModulation6);
            schemeColor69.Append(saturationModulation3);
            schemeColor69.Append(tint15);

            gradientStop3.Append(schemeColor69);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor70 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 103000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 102000 };
            A.Tint tint16 = new A.Tint() { Val = 94000 };

            schemeColor70.Append(saturationModulation4);
            schemeColor70.Append(luminanceModulation7);
            schemeColor70.Append(tint16);

            gradientStop4.Append(schemeColor70);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor71 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 110000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 100000 };
            A.Shade shade1 = new A.Shade() { Val = 100000 };

            schemeColor71.Append(saturationModulation5);
            schemeColor71.Append(luminanceModulation8);
            schemeColor71.Append(shade1);

            gradientStop5.Append(schemeColor71);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor72 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation9 = new A.LuminanceModulation() { Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 120000 };
            A.Shade shade2 = new A.Shade() { Val = 78000 };

            schemeColor72.Append(luminanceModulation9);
            schemeColor72.Append(saturationModulation6);
            schemeColor72.Append(shade2);

            gradientStop6.Append(schemeColor72);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill98);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline30 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill99 = new A.SolidFill();
            A.SchemeColor schemeColor73 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill99.Append(schemeColor73);
            A.PresetDash presetDash82 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter() { Limit = 800000 };

            outline30.Append(solidFill99);
            outline30.Append(presetDash82);
            outline30.Append(miter1);

            A.Outline outline31 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill100 = new A.SolidFill();
            A.SchemeColor schemeColor74 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill100.Append(schemeColor74);
            A.PresetDash presetDash83 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter() { Limit = 800000 };

            outline31.Append(solidFill100);
            outline31.Append(presetDash83);
            outline31.Append(miter2);

            A.Outline outline32 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill101 = new A.SolidFill();
            A.SchemeColor schemeColor75 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill101.Append(schemeColor75);
            A.PresetDash presetDash84 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter() { Limit = 800000 };

            outline32.Append(solidFill101);
            outline32.Append(presetDash84);
            outline32.Append(miter3);

            lineStyleList1.Append(outline30);
            lineStyleList1.Append(outline31);
            lineStyleList1.Append(outline32);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();
            A.EffectList effectList23 = new A.EffectList();

            effectStyle1.Append(effectList23);

            A.EffectStyle effectStyle2 = new A.EffectStyle();
            A.EffectList effectList24 = new A.EffectList();

            effectStyle2.Append(effectList24);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList25 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex37 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha16 = new A.Alpha() { Val = 63000 };

            rgbColorModelHex37.Append(alpha16);

            outerShadow2.Append(rgbColorModelHex37);

            effectList25.Append(outerShadow2);

            effectStyle3.Append(effectList25);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill102 = new A.SolidFill();
            A.SchemeColor schemeColor76 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill102.Append(schemeColor76);

            A.SolidFill solidFill103 = new A.SolidFill();

            A.SchemeColor schemeColor77 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint17 = new A.Tint() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

            schemeColor77.Append(tint17);
            schemeColor77.Append(saturationModulation7);

            solidFill103.Append(schemeColor77);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor78 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint18 = new A.Tint() { Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 150000 };
            A.Shade shade3 = new A.Shade() { Val = 98000 };
            A.LuminanceModulation luminanceModulation10 = new A.LuminanceModulation() { Val = 102000 };

            schemeColor78.Append(tint18);
            schemeColor78.Append(saturationModulation8);
            schemeColor78.Append(shade3);
            schemeColor78.Append(luminanceModulation10);

            gradientStop7.Append(schemeColor78);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor79 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint19 = new A.Tint() { Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 130000 };
            A.Shade shade4 = new A.Shade() { Val = 90000 };
            A.LuminanceModulation luminanceModulation11 = new A.LuminanceModulation() { Val = 103000 };

            schemeColor79.Append(tint19);
            schemeColor79.Append(saturationModulation9);
            schemeColor79.Append(shade4);
            schemeColor79.Append(luminanceModulation11);

            gradientStop8.Append(schemeColor79);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor80 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade5 = new A.Shade() { Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 120000 };

            schemeColor80.Append(shade5);
            schemeColor80.Append(saturationModulation10);

            gradientStop9.Append(schemeColor80);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList1.Append(solidFill102);
            backgroundFillStyleList1.Append(solidFill103);
            backgroundFillStyleList1.Append(gradientFill3);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);

            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();

            A.LineDefault lineDefault1 = new A.LineDefault();
            A.ShapeProperties shapeProperties50 = new A.ShapeProperties();
            A.BodyProperties bodyProperties63 = new A.BodyProperties();
            A.ListStyle listStyle63 = new A.ListStyle();

            A.ShapeStyle shapeStyle1 = new A.ShapeStyle();

            A.LineReference lineReference1 = new A.LineReference() { Index = (UInt32Value)2U };
            A.SchemeColor schemeColor81 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            lineReference1.Append(schemeColor81);

            A.FillReference fillReference1 = new A.FillReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor82 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference1.Append(schemeColor82);

            A.EffectReference effectReference1 = new A.EffectReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor83 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference1.Append(schemeColor83);

            A.FontReference fontReference1 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor84 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference1.Append(schemeColor84);

            shapeStyle1.Append(lineReference1);
            shapeStyle1.Append(fillReference1);
            shapeStyle1.Append(effectReference1);
            shapeStyle1.Append(fontReference1);

            lineDefault1.Append(shapeProperties50);
            lineDefault1.Append(bodyProperties63);
            lineDefault1.Append(listStyle63);
            lineDefault1.Append(shapeStyle1);

            objectDefaults1.Append(lineDefault1);
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            A.OfficeStyleSheetExtensionList officeStyleSheetExtensionList1 = new A.OfficeStyleSheetExtensionList();

            A.OfficeStyleSheetExtension officeStyleSheetExtension1 = new A.OfficeStyleSheetExtension() { Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

            Thm15.ThemeFamily themeFamily1 = new Thm15.ThemeFamily() { Name = "Office Theme", Id = "{2E142A2C-CD16-42D6-873A-C26D2A0506FA}", Vid = "{1BDDFF52-6CD6-40A5-AB3C-68EB2F1E4D0A}" };
            themeFamily1.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

            officeStyleSheetExtension1.Append(themeFamily1);

            officeStyleSheetExtensionList1.Append(officeStyleSheetExtension1);

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);
            theme1.Append(officeStyleSheetExtensionList1);

            themePart1.Theme = theme1;
        }

        // Generates content of slideLayoutPart5.
        private void GenerateSlideLayoutPart5Content(SlideLayoutPart slideLayoutPart5)
        {
            SlideLayout slideLayout5 = new SlideLayout() { Type = SlideLayoutValues.Object, Preserve = true };
            slideLayout5.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout5.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout5.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData7 = new CommonSlideData() { Name = "Title and Content" };

            ShapeTree shapeTree7 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties13 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties63 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties13 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties63 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties13.Append(nonVisualDrawingProperties63);
            nonVisualGroupShapeProperties13.Append(nonVisualGroupShapeDrawingProperties13);
            nonVisualGroupShapeProperties13.Append(applicationNonVisualDrawingProperties63);

            GroupShapeProperties groupShapeProperties13 = new GroupShapeProperties();

            A.TransformGroup transformGroup13 = new A.TransformGroup();
            A.Offset offset49 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents49 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset13 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents13 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup13.Append(offset49);
            transformGroup13.Append(extents49);
            transformGroup13.Append(childOffset13);
            transformGroup13.Append(childExtents13);

            groupShapeProperties13.Append(transformGroup13);

            Shape shape42 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties42 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties64 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList57 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension57 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement77 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{A6A0301D-C2B5-988E-B772-9C55574EF9C3}\" />");

            // nonVisualDrawingPropertiesExtension57.Append(openXmlUnknownElement77);

            nonVisualDrawingPropertiesExtensionList57.Append(nonVisualDrawingPropertiesExtension57);

            nonVisualDrawingProperties64.Append(nonVisualDrawingPropertiesExtensionList57);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties42 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks27 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties42.Append(shapeLocks27);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties64 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape27 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties64.Append(placeholderShape27);

            nonVisualShapeProperties42.Append(nonVisualDrawingProperties64);
            nonVisualShapeProperties42.Append(nonVisualShapeDrawingProperties42);
            nonVisualShapeProperties42.Append(applicationNonVisualDrawingProperties64);
            ShapeProperties shapeProperties51 = new ShapeProperties();

            TextBody textBody63 = new TextBody();
            A.BodyProperties bodyProperties64 = new A.BodyProperties();
            A.ListStyle listStyle64 = new A.ListStyle();

            A.Paragraph paragraph71 = new A.Paragraph();

            A.Run run44 = new A.Run();
            A.RunProperties runProperties54 = new A.RunProperties() { Language = "en-GB" };
            A.Text text54 = new A.Text();
            text54.Text = "Click to edit Master title style";

            run44.Append(runProperties54);
            run44.Append(text54);
            A.EndParagraphRunProperties endParagraphRunProperties56 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph71.Append(run44);
            paragraph71.Append(endParagraphRunProperties56);

            textBody63.Append(bodyProperties64);
            textBody63.Append(listStyle64);
            textBody63.Append(paragraph71);

            shape42.Append(nonVisualShapeProperties42);
            shape42.Append(shapeProperties51);
            shape42.Append(textBody63);

            Shape shape43 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties43 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties65 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Content Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList58 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension58 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement78 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{774BF6B6-A17B-8ECE-A42B-1C9959F1471B}\" />");

            // nonVisualDrawingPropertiesExtension58.Append(openXmlUnknownElement78);

            nonVisualDrawingPropertiesExtensionList58.Append(nonVisualDrawingPropertiesExtension58);

            nonVisualDrawingProperties65.Append(nonVisualDrawingPropertiesExtensionList58);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties43 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks28 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties43.Append(shapeLocks28);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties65 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape28 = new PlaceholderShape() { Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties65.Append(placeholderShape28);

            nonVisualShapeProperties43.Append(nonVisualDrawingProperties65);
            nonVisualShapeProperties43.Append(nonVisualShapeDrawingProperties43);
            nonVisualShapeProperties43.Append(applicationNonVisualDrawingProperties65);
            ShapeProperties shapeProperties52 = new ShapeProperties();

            TextBody textBody64 = new TextBody();
            A.BodyProperties bodyProperties65 = new A.BodyProperties();
            A.ListStyle listStyle65 = new A.ListStyle();

            A.Paragraph paragraph72 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties45 = new A.ParagraphProperties() { Level = 0 };

            A.Run run45 = new A.Run();
            A.RunProperties runProperties55 = new A.RunProperties() { Language = "en-GB" };
            A.Text text55 = new A.Text();
            text55.Text = "Click to edit Master text styles";

            run45.Append(runProperties55);
            run45.Append(text55);

            paragraph72.Append(paragraphProperties45);
            paragraph72.Append(run45);

            A.Paragraph paragraph73 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties46 = new A.ParagraphProperties() { Level = 1 };

            A.Run run46 = new A.Run();
            A.RunProperties runProperties56 = new A.RunProperties() { Language = "en-GB" };
            A.Text text56 = new A.Text();
            text56.Text = "Second level";

            run46.Append(runProperties56);
            run46.Append(text56);

            paragraph73.Append(paragraphProperties46);
            paragraph73.Append(run46);

            A.Paragraph paragraph74 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties47 = new A.ParagraphProperties() { Level = 2 };

            A.Run run47 = new A.Run();
            A.RunProperties runProperties57 = new A.RunProperties() { Language = "en-GB" };
            A.Text text57 = new A.Text();
            text57.Text = "Third level";

            run47.Append(runProperties57);
            run47.Append(text57);

            paragraph74.Append(paragraphProperties47);
            paragraph74.Append(run47);

            A.Paragraph paragraph75 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties48 = new A.ParagraphProperties() { Level = 3 };

            A.Run run48 = new A.Run();
            A.RunProperties runProperties58 = new A.RunProperties() { Language = "en-GB" };
            A.Text text58 = new A.Text();
            text58.Text = "Fourth level";

            run48.Append(runProperties58);
            run48.Append(text58);

            paragraph75.Append(paragraphProperties48);
            paragraph75.Append(run48);

            A.Paragraph paragraph76 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties49 = new A.ParagraphProperties() { Level = 4 };

            A.Run run49 = new A.Run();
            A.RunProperties runProperties59 = new A.RunProperties() { Language = "en-GB" };
            A.Text text59 = new A.Text();
            text59.Text = "Fifth level";

            run49.Append(runProperties59);
            run49.Append(text59);
            A.EndParagraphRunProperties endParagraphRunProperties57 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph76.Append(paragraphProperties49);
            paragraph76.Append(run49);
            paragraph76.Append(endParagraphRunProperties57);

            textBody64.Append(bodyProperties65);
            textBody64.Append(listStyle65);
            textBody64.Append(paragraph72);
            textBody64.Append(paragraph73);
            textBody64.Append(paragraph74);
            textBody64.Append(paragraph75);
            textBody64.Append(paragraph76);

            shape43.Append(nonVisualShapeProperties43);
            shape43.Append(shapeProperties52);
            shape43.Append(textBody64);

            Shape shape44 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties44 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties66 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList59 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension59 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement79 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{4F739CC7-90EA-F93B-3177-8B858A2AB2B7}\" />");

            // nonVisualDrawingPropertiesExtension59.Append(openXmlUnknownElement79);

            nonVisualDrawingPropertiesExtensionList59.Append(nonVisualDrawingPropertiesExtension59);

            nonVisualDrawingProperties66.Append(nonVisualDrawingPropertiesExtensionList59);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties44 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks29 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties44.Append(shapeLocks29);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties66 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape29 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties66.Append(placeholderShape29);

            nonVisualShapeProperties44.Append(nonVisualDrawingProperties66);
            nonVisualShapeProperties44.Append(nonVisualShapeDrawingProperties44);
            nonVisualShapeProperties44.Append(applicationNonVisualDrawingProperties66);
            ShapeProperties shapeProperties53 = new ShapeProperties();

            TextBody textBody65 = new TextBody();
            A.BodyProperties bodyProperties66 = new A.BodyProperties();
            A.ListStyle listStyle66 = new A.ListStyle();

            A.Paragraph paragraph77 = new A.Paragraph();

            A.Field field11 = new A.Field() { Id = "{C453C8C6-607B-FF49-A93F-867C282599C0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties60 = new A.RunProperties() { Language = "en-US" };
            runProperties60.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text60 = new A.Text();
            text60.Text = "5/3/2024";

            field11.Append(runProperties60);
            field11.Append(text60);
            A.EndParagraphRunProperties endParagraphRunProperties58 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph77.Append(field11);
            paragraph77.Append(endParagraphRunProperties58);

            textBody65.Append(bodyProperties66);
            textBody65.Append(listStyle66);
            textBody65.Append(paragraph77);

            shape44.Append(nonVisualShapeProperties44);
            shape44.Append(shapeProperties53);
            shape44.Append(textBody65);

            Shape shape45 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties45 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties67 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList60 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension60 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement80 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{9F923CF6-4339-7214-64FE-05254CE154A5}\" />");

            // nonVisualDrawingPropertiesExtension60.Append(openXmlUnknownElement80);

            nonVisualDrawingPropertiesExtensionList60.Append(nonVisualDrawingPropertiesExtension60);

            nonVisualDrawingProperties67.Append(nonVisualDrawingPropertiesExtensionList60);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties45 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks30 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties45.Append(shapeLocks30);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties67 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape30 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties67.Append(placeholderShape30);

            nonVisualShapeProperties45.Append(nonVisualDrawingProperties67);
            nonVisualShapeProperties45.Append(nonVisualShapeDrawingProperties45);
            nonVisualShapeProperties45.Append(applicationNonVisualDrawingProperties67);
            ShapeProperties shapeProperties54 = new ShapeProperties();

            TextBody textBody66 = new TextBody();
            A.BodyProperties bodyProperties67 = new A.BodyProperties();
            A.ListStyle listStyle67 = new A.ListStyle();

            A.Paragraph paragraph78 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties59 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph78.Append(endParagraphRunProperties59);

            textBody66.Append(bodyProperties67);
            textBody66.Append(listStyle67);
            textBody66.Append(paragraph78);

            shape45.Append(nonVisualShapeProperties45);
            shape45.Append(shapeProperties54);
            shape45.Append(textBody66);

            Shape shape46 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties46 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties68 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList61 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension61 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement81 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{C4CEBDB1-364D-B0A4-1601-CEB6EBD102D1}\" />");

            // nonVisualDrawingPropertiesExtension61.Append(openXmlUnknownElement81);

            nonVisualDrawingPropertiesExtensionList61.Append(nonVisualDrawingPropertiesExtension61);

            nonVisualDrawingProperties68.Append(nonVisualDrawingPropertiesExtensionList61);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties46 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks31 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties46.Append(shapeLocks31);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties68 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape31 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties68.Append(placeholderShape31);

            nonVisualShapeProperties46.Append(nonVisualDrawingProperties68);
            nonVisualShapeProperties46.Append(nonVisualShapeDrawingProperties46);
            nonVisualShapeProperties46.Append(applicationNonVisualDrawingProperties68);
            ShapeProperties shapeProperties55 = new ShapeProperties();

            TextBody textBody67 = new TextBody();
            A.BodyProperties bodyProperties68 = new A.BodyProperties();
            A.ListStyle listStyle68 = new A.ListStyle();

            A.Paragraph paragraph79 = new A.Paragraph();

            A.Field field12 = new A.Field() { Id = "{402C1264-1A6A-2D47-B539-28043A88F412}", Type = "slidenum" };

            A.RunProperties runProperties61 = new A.RunProperties() { Language = "en-US" };
            runProperties61.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text61 = new A.Text();
            text61.Text = "‹#›";

            field12.Append(runProperties61);
            field12.Append(text61);
            A.EndParagraphRunProperties endParagraphRunProperties60 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph79.Append(field12);
            paragraph79.Append(endParagraphRunProperties60);

            textBody67.Append(bodyProperties68);
            textBody67.Append(listStyle68);
            textBody67.Append(paragraph79);

            shape46.Append(nonVisualShapeProperties46);
            shape46.Append(shapeProperties55);
            shape46.Append(textBody67);

            shapeTree7.Append(nonVisualGroupShapeProperties13);
            shapeTree7.Append(groupShapeProperties13);
            shapeTree7.Append(shape42);
            shapeTree7.Append(shape43);
            shapeTree7.Append(shape44);
            shapeTree7.Append(shape45);
            shapeTree7.Append(shape46);

            CommonSlideDataExtensionList commonSlideDataExtensionList7 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension7 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId7 = new P14.CreationId() { Val = (UInt32Value)2814172873U };
            creationId7.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension7.Append(creationId7);

            commonSlideDataExtensionList7.Append(commonSlideDataExtension7);

            commonSlideData7.Append(shapeTree7);
            commonSlideData7.Append(commonSlideDataExtensionList7);

            ColorMapOverride colorMapOverride6 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping6 = new A.MasterColorMapping();

            colorMapOverride6.Append(masterColorMapping6);

            slideLayout5.Append(commonSlideData7);
            slideLayout5.Append(colorMapOverride6);

            slideLayoutPart5.SlideLayout = slideLayout5;
        }

        // Generates content of slideLayoutPart6.
        private void GenerateSlideLayoutPart6Content(SlideLayoutPart slideLayoutPart6)
        {
            SlideLayout slideLayout6 = new SlideLayout() { Type = SlideLayoutValues.TitleOnly, Preserve = true };
            slideLayout6.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout6.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout6.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData8 = new CommonSlideData() { Name = "Title Only" };

            ShapeTree shapeTree8 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties14 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties69 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties14 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties69 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties14.Append(nonVisualDrawingProperties69);
            nonVisualGroupShapeProperties14.Append(nonVisualGroupShapeDrawingProperties14);
            nonVisualGroupShapeProperties14.Append(applicationNonVisualDrawingProperties69);

            GroupShapeProperties groupShapeProperties14 = new GroupShapeProperties();

            A.TransformGroup transformGroup14 = new A.TransformGroup();
            A.Offset offset50 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents50 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset14 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents14 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup14.Append(offset50);
            transformGroup14.Append(extents50);
            transformGroup14.Append(childOffset14);
            transformGroup14.Append(childExtents14);

            groupShapeProperties14.Append(transformGroup14);

            Shape shape47 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties47 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties70 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList62 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension62 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement82 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{5D655540-61C5-D422-EC26-2F16766C986B}\" />");

            // nonVisualDrawingPropertiesExtension62.Append(openXmlUnknownElement82);

            nonVisualDrawingPropertiesExtensionList62.Append(nonVisualDrawingPropertiesExtension62);

            nonVisualDrawingProperties70.Append(nonVisualDrawingPropertiesExtensionList62);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties47 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks32 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties47.Append(shapeLocks32);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties70 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape32 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties70.Append(placeholderShape32);

            nonVisualShapeProperties47.Append(nonVisualDrawingProperties70);
            nonVisualShapeProperties47.Append(nonVisualShapeDrawingProperties47);
            nonVisualShapeProperties47.Append(applicationNonVisualDrawingProperties70);
            ShapeProperties shapeProperties56 = new ShapeProperties();

            TextBody textBody68 = new TextBody();
            A.BodyProperties bodyProperties69 = new A.BodyProperties();
            A.ListStyle listStyle69 = new A.ListStyle();

            A.Paragraph paragraph80 = new A.Paragraph();

            A.Run run50 = new A.Run();
            A.RunProperties runProperties62 = new A.RunProperties() { Language = "en-GB" };
            A.Text text62 = new A.Text();
            text62.Text = "Click to edit Master title style";

            run50.Append(runProperties62);
            run50.Append(text62);
            A.EndParagraphRunProperties endParagraphRunProperties61 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph80.Append(run50);
            paragraph80.Append(endParagraphRunProperties61);

            textBody68.Append(bodyProperties69);
            textBody68.Append(listStyle69);
            textBody68.Append(paragraph80);

            shape47.Append(nonVisualShapeProperties47);
            shape47.Append(shapeProperties56);
            shape47.Append(textBody68);

            Shape shape48 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties48 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties71 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Date Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList63 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension63 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement83 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{25E7510F-412D-E358-3F12-D56A11005344}\" />");

            // nonVisualDrawingPropertiesExtension63.Append(openXmlUnknownElement83);

            nonVisualDrawingPropertiesExtensionList63.Append(nonVisualDrawingPropertiesExtension63);

            nonVisualDrawingProperties71.Append(nonVisualDrawingPropertiesExtensionList63);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties48 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks33 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties48.Append(shapeLocks33);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties71 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape33 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties71.Append(placeholderShape33);

            nonVisualShapeProperties48.Append(nonVisualDrawingProperties71);
            nonVisualShapeProperties48.Append(nonVisualShapeDrawingProperties48);
            nonVisualShapeProperties48.Append(applicationNonVisualDrawingProperties71);
            ShapeProperties shapeProperties57 = new ShapeProperties();

            TextBody textBody69 = new TextBody();
            A.BodyProperties bodyProperties70 = new A.BodyProperties();
            A.ListStyle listStyle70 = new A.ListStyle();

            A.Paragraph paragraph81 = new A.Paragraph();

            A.Field field13 = new A.Field() { Id = "{C453C8C6-607B-FF49-A93F-867C282599C0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties63 = new A.RunProperties() { Language = "en-US" };
            runProperties63.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text63 = new A.Text();
            text63.Text = "5/3/2024";

            field13.Append(runProperties63);
            field13.Append(text63);
            A.EndParagraphRunProperties endParagraphRunProperties62 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph81.Append(field13);
            paragraph81.Append(endParagraphRunProperties62);

            textBody69.Append(bodyProperties70);
            textBody69.Append(listStyle70);
            textBody69.Append(paragraph81);

            shape48.Append(nonVisualShapeProperties48);
            shape48.Append(shapeProperties57);
            shape48.Append(textBody69);

            Shape shape49 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties49 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties72 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Footer Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList64 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension64 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement84 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{63E2A94E-DB7A-9653-8D11-04B56ADB710F}\" />");

            // nonVisualDrawingPropertiesExtension64.Append(openXmlUnknownElement84);

            nonVisualDrawingPropertiesExtensionList64.Append(nonVisualDrawingPropertiesExtension64);

            nonVisualDrawingProperties72.Append(nonVisualDrawingPropertiesExtensionList64);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties49 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks34 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties49.Append(shapeLocks34);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties72 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape34 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties72.Append(placeholderShape34);

            nonVisualShapeProperties49.Append(nonVisualDrawingProperties72);
            nonVisualShapeProperties49.Append(nonVisualShapeDrawingProperties49);
            nonVisualShapeProperties49.Append(applicationNonVisualDrawingProperties72);
            ShapeProperties shapeProperties58 = new ShapeProperties();

            TextBody textBody70 = new TextBody();
            A.BodyProperties bodyProperties71 = new A.BodyProperties();
            A.ListStyle listStyle71 = new A.ListStyle();

            A.Paragraph paragraph82 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties63 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph82.Append(endParagraphRunProperties63);

            textBody70.Append(bodyProperties71);
            textBody70.Append(listStyle71);
            textBody70.Append(paragraph82);

            shape49.Append(nonVisualShapeProperties49);
            shape49.Append(shapeProperties58);
            shape49.Append(textBody70);

            Shape shape50 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties50 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties73 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Slide Number Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList65 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension65 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement85 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{AC2FC352-4E24-BE6B-9775-E94427FBFD2D}\" />");

            // nonVisualDrawingPropertiesExtension65.Append(openXmlUnknownElement85);

            nonVisualDrawingPropertiesExtensionList65.Append(nonVisualDrawingPropertiesExtension65);

            nonVisualDrawingProperties73.Append(nonVisualDrawingPropertiesExtensionList65);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties50 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks35 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties50.Append(shapeLocks35);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties73 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape35 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties73.Append(placeholderShape35);

            nonVisualShapeProperties50.Append(nonVisualDrawingProperties73);
            nonVisualShapeProperties50.Append(nonVisualShapeDrawingProperties50);
            nonVisualShapeProperties50.Append(applicationNonVisualDrawingProperties73);
            ShapeProperties shapeProperties59 = new ShapeProperties();

            TextBody textBody71 = new TextBody();
            A.BodyProperties bodyProperties72 = new A.BodyProperties();
            A.ListStyle listStyle72 = new A.ListStyle();

            A.Paragraph paragraph83 = new A.Paragraph();

            A.Field field14 = new A.Field() { Id = "{402C1264-1A6A-2D47-B539-28043A88F412}", Type = "slidenum" };

            A.RunProperties runProperties64 = new A.RunProperties() { Language = "en-US" };
            runProperties64.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text64 = new A.Text();
            text64.Text = "‹#›";

            field14.Append(runProperties64);
            field14.Append(text64);
            A.EndParagraphRunProperties endParagraphRunProperties64 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph83.Append(field14);
            paragraph83.Append(endParagraphRunProperties64);

            textBody71.Append(bodyProperties72);
            textBody71.Append(listStyle72);
            textBody71.Append(paragraph83);

            shape50.Append(nonVisualShapeProperties50);
            shape50.Append(shapeProperties59);
            shape50.Append(textBody71);

            shapeTree8.Append(nonVisualGroupShapeProperties14);
            shapeTree8.Append(groupShapeProperties14);
            shapeTree8.Append(shape47);
            shapeTree8.Append(shape48);
            shapeTree8.Append(shape49);
            shapeTree8.Append(shape50);

            CommonSlideDataExtensionList commonSlideDataExtensionList8 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension8 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId8 = new P14.CreationId() { Val = (UInt32Value)1549284485U };
            creationId8.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension8.Append(creationId8);

            commonSlideDataExtensionList8.Append(commonSlideDataExtension8);

            commonSlideData8.Append(shapeTree8);
            commonSlideData8.Append(commonSlideDataExtensionList8);

            ColorMapOverride colorMapOverride7 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping7 = new A.MasterColorMapping();

            colorMapOverride7.Append(masterColorMapping7);

            slideLayout6.Append(commonSlideData8);
            slideLayout6.Append(colorMapOverride7);

            slideLayoutPart6.SlideLayout = slideLayout6;
        }

        // Generates content of slideLayoutPart7.
        private void GenerateSlideLayoutPart7Content(SlideLayoutPart slideLayoutPart7)
        {
            SlideLayout slideLayout7 = new SlideLayout() { Type = SlideLayoutValues.VerticalTitleAndText, Preserve = true };
            slideLayout7.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout7.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout7.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData9 = new CommonSlideData() { Name = "Vertical Title and Text" };

            ShapeTree shapeTree9 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties15 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties74 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties15 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties74 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties15.Append(nonVisualDrawingProperties74);
            nonVisualGroupShapeProperties15.Append(nonVisualGroupShapeDrawingProperties15);
            nonVisualGroupShapeProperties15.Append(applicationNonVisualDrawingProperties74);

            GroupShapeProperties groupShapeProperties15 = new GroupShapeProperties();

            A.TransformGroup transformGroup15 = new A.TransformGroup();
            A.Offset offset51 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents51 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset15 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents15 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup15.Append(offset51);
            transformGroup15.Append(extents51);
            transformGroup15.Append(childOffset15);
            transformGroup15.Append(childExtents15);

            groupShapeProperties15.Append(transformGroup15);

            Shape shape51 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties51 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties75 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Vertical Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList66 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension66 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement86 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{30BFFE16-C34C-C16F-5392-C8CD12919754}\" />");

            // nonVisualDrawingPropertiesExtension66.Append(openXmlUnknownElement86);

            nonVisualDrawingPropertiesExtensionList66.Append(nonVisualDrawingPropertiesExtension66);

            nonVisualDrawingProperties75.Append(nonVisualDrawingPropertiesExtensionList66);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties51 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks36 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties51.Append(shapeLocks36);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties75 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape36 = new PlaceholderShape() { Type = PlaceholderValues.Title, Orientation = DirectionValues.Vertical };

            applicationNonVisualDrawingProperties75.Append(placeholderShape36);

            nonVisualShapeProperties51.Append(nonVisualDrawingProperties75);
            nonVisualShapeProperties51.Append(nonVisualShapeDrawingProperties51);
            nonVisualShapeProperties51.Append(applicationNonVisualDrawingProperties75);

            ShapeProperties shapeProperties60 = new ShapeProperties();

            A.Transform2D transform2D36 = new A.Transform2D();
            A.Offset offset52 = new A.Offset() { X = 8724900L, Y = 365125L };
            A.Extents extents52 = new A.Extents() { Cx = 2628900L, Cy = 5811838L };

            transform2D36.Append(offset52);
            transform2D36.Append(extents52);

            shapeProperties60.Append(transform2D36);

            TextBody textBody72 = new TextBody();
            A.BodyProperties bodyProperties73 = new A.BodyProperties() { Vertical = A.TextVerticalValues.EastAsianVetical };
            A.ListStyle listStyle73 = new A.ListStyle();

            A.Paragraph paragraph84 = new A.Paragraph();

            A.Run run51 = new A.Run();
            A.RunProperties runProperties65 = new A.RunProperties() { Language = "en-GB" };
            A.Text text65 = new A.Text();
            text65.Text = "Click to edit Master title style";

            run51.Append(runProperties65);
            run51.Append(text65);
            A.EndParagraphRunProperties endParagraphRunProperties65 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph84.Append(run51);
            paragraph84.Append(endParagraphRunProperties65);

            textBody72.Append(bodyProperties73);
            textBody72.Append(listStyle73);
            textBody72.Append(paragraph84);

            shape51.Append(nonVisualShapeProperties51);
            shape51.Append(shapeProperties60);
            shape51.Append(textBody72);

            Shape shape52 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties52 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties76 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Vertical Text Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList67 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension67 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement87 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{DCE89AA1-2CA3-538C-FEC3-3E1EF2A88A8E}\" />");

            // nonVisualDrawingPropertiesExtension67.Append(openXmlUnknownElement87);

            nonVisualDrawingPropertiesExtensionList67.Append(nonVisualDrawingPropertiesExtension67);

            nonVisualDrawingProperties76.Append(nonVisualDrawingPropertiesExtensionList67);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties52 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks37 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties52.Append(shapeLocks37);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties76 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape37 = new PlaceholderShape() { Type = PlaceholderValues.Body, Orientation = DirectionValues.Vertical, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties76.Append(placeholderShape37);

            nonVisualShapeProperties52.Append(nonVisualDrawingProperties76);
            nonVisualShapeProperties52.Append(nonVisualShapeDrawingProperties52);
            nonVisualShapeProperties52.Append(applicationNonVisualDrawingProperties76);

            ShapeProperties shapeProperties61 = new ShapeProperties();

            A.Transform2D transform2D37 = new A.Transform2D();
            A.Offset offset53 = new A.Offset() { X = 838200L, Y = 365125L };
            A.Extents extents53 = new A.Extents() { Cx = 7734300L, Cy = 5811838L };

            transform2D37.Append(offset53);
            transform2D37.Append(extents53);

            shapeProperties61.Append(transform2D37);

            TextBody textBody73 = new TextBody();
            A.BodyProperties bodyProperties74 = new A.BodyProperties() { Vertical = A.TextVerticalValues.EastAsianVetical };
            A.ListStyle listStyle74 = new A.ListStyle();

            A.Paragraph paragraph85 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties50 = new A.ParagraphProperties() { Level = 0 };

            A.Run run52 = new A.Run();
            A.RunProperties runProperties66 = new A.RunProperties() { Language = "en-GB" };
            A.Text text66 = new A.Text();
            text66.Text = "Click to edit Master text styles";

            run52.Append(runProperties66);
            run52.Append(text66);

            paragraph85.Append(paragraphProperties50);
            paragraph85.Append(run52);

            A.Paragraph paragraph86 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties51 = new A.ParagraphProperties() { Level = 1 };

            A.Run run53 = new A.Run();
            A.RunProperties runProperties67 = new A.RunProperties() { Language = "en-GB" };
            A.Text text67 = new A.Text();
            text67.Text = "Second level";

            run53.Append(runProperties67);
            run53.Append(text67);

            paragraph86.Append(paragraphProperties51);
            paragraph86.Append(run53);

            A.Paragraph paragraph87 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties52 = new A.ParagraphProperties() { Level = 2 };

            A.Run run54 = new A.Run();
            A.RunProperties runProperties68 = new A.RunProperties() { Language = "en-GB" };
            A.Text text68 = new A.Text();
            text68.Text = "Third level";

            run54.Append(runProperties68);
            run54.Append(text68);

            paragraph87.Append(paragraphProperties52);
            paragraph87.Append(run54);

            A.Paragraph paragraph88 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties53 = new A.ParagraphProperties() { Level = 3 };

            A.Run run55 = new A.Run();
            A.RunProperties runProperties69 = new A.RunProperties() { Language = "en-GB" };
            A.Text text69 = new A.Text();
            text69.Text = "Fourth level";

            run55.Append(runProperties69);
            run55.Append(text69);

            paragraph88.Append(paragraphProperties53);
            paragraph88.Append(run55);

            A.Paragraph paragraph89 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties54 = new A.ParagraphProperties() { Level = 4 };

            A.Run run56 = new A.Run();
            A.RunProperties runProperties70 = new A.RunProperties() { Language = "en-GB" };
            A.Text text70 = new A.Text();
            text70.Text = "Fifth level";

            run56.Append(runProperties70);
            run56.Append(text70);
            A.EndParagraphRunProperties endParagraphRunProperties66 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph89.Append(paragraphProperties54);
            paragraph89.Append(run56);
            paragraph89.Append(endParagraphRunProperties66);

            textBody73.Append(bodyProperties74);
            textBody73.Append(listStyle74);
            textBody73.Append(paragraph85);
            textBody73.Append(paragraph86);
            textBody73.Append(paragraph87);
            textBody73.Append(paragraph88);
            textBody73.Append(paragraph89);

            shape52.Append(nonVisualShapeProperties52);
            shape52.Append(shapeProperties61);
            shape52.Append(textBody73);

            Shape shape53 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties53 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties77 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList68 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension68 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement88 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{011EE19B-B164-7B5B-BF06-E00FA259E2EB}\" />");

            // nonVisualDrawingPropertiesExtension68.Append(openXmlUnknownElement88);

            nonVisualDrawingPropertiesExtensionList68.Append(nonVisualDrawingPropertiesExtension68);

            nonVisualDrawingProperties77.Append(nonVisualDrawingPropertiesExtensionList68);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties53 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks38 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties53.Append(shapeLocks38);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties77 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape38 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties77.Append(placeholderShape38);

            nonVisualShapeProperties53.Append(nonVisualDrawingProperties77);
            nonVisualShapeProperties53.Append(nonVisualShapeDrawingProperties53);
            nonVisualShapeProperties53.Append(applicationNonVisualDrawingProperties77);
            ShapeProperties shapeProperties62 = new ShapeProperties();

            TextBody textBody74 = new TextBody();
            A.BodyProperties bodyProperties75 = new A.BodyProperties();
            A.ListStyle listStyle75 = new A.ListStyle();

            A.Paragraph paragraph90 = new A.Paragraph();

            A.Field field15 = new A.Field() { Id = "{C453C8C6-607B-FF49-A93F-867C282599C0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties71 = new A.RunProperties() { Language = "en-US" };
            runProperties71.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text71 = new A.Text();
            text71.Text = "5/3/2024";

            field15.Append(runProperties71);
            field15.Append(text71);
            A.EndParagraphRunProperties endParagraphRunProperties67 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph90.Append(field15);
            paragraph90.Append(endParagraphRunProperties67);

            textBody74.Append(bodyProperties75);
            textBody74.Append(listStyle75);
            textBody74.Append(paragraph90);

            shape53.Append(nonVisualShapeProperties53);
            shape53.Append(shapeProperties62);
            shape53.Append(textBody74);

            Shape shape54 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties54 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties78 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList69 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension69 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement89 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{91E9DA0E-C57C-70E6-C065-C40C989E70DE}\" />");

            // nonVisualDrawingPropertiesExtension69.Append(openXmlUnknownElement89);

            nonVisualDrawingPropertiesExtensionList69.Append(nonVisualDrawingPropertiesExtension69);

            nonVisualDrawingProperties78.Append(nonVisualDrawingPropertiesExtensionList69);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties54 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks39 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties54.Append(shapeLocks39);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties78 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape39 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties78.Append(placeholderShape39);

            nonVisualShapeProperties54.Append(nonVisualDrawingProperties78);
            nonVisualShapeProperties54.Append(nonVisualShapeDrawingProperties54);
            nonVisualShapeProperties54.Append(applicationNonVisualDrawingProperties78);
            ShapeProperties shapeProperties63 = new ShapeProperties();

            TextBody textBody75 = new TextBody();
            A.BodyProperties bodyProperties76 = new A.BodyProperties();
            A.ListStyle listStyle76 = new A.ListStyle();

            A.Paragraph paragraph91 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties68 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph91.Append(endParagraphRunProperties68);

            textBody75.Append(bodyProperties76);
            textBody75.Append(listStyle76);
            textBody75.Append(paragraph91);

            shape54.Append(nonVisualShapeProperties54);
            shape54.Append(shapeProperties63);
            shape54.Append(textBody75);

            Shape shape55 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties55 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties79 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList70 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension70 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement90 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{66B3D7A7-3E0A-1344-518E-43BEBB8486D4}\" />");

            // nonVisualDrawingPropertiesExtension70.Append(openXmlUnknownElement90);

            nonVisualDrawingPropertiesExtensionList70.Append(nonVisualDrawingPropertiesExtension70);

            nonVisualDrawingProperties79.Append(nonVisualDrawingPropertiesExtensionList70);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties55 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks40 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties55.Append(shapeLocks40);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties79 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape40 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties79.Append(placeholderShape40);

            nonVisualShapeProperties55.Append(nonVisualDrawingProperties79);
            nonVisualShapeProperties55.Append(nonVisualShapeDrawingProperties55);
            nonVisualShapeProperties55.Append(applicationNonVisualDrawingProperties79);
            ShapeProperties shapeProperties64 = new ShapeProperties();

            TextBody textBody76 = new TextBody();
            A.BodyProperties bodyProperties77 = new A.BodyProperties();
            A.ListStyle listStyle77 = new A.ListStyle();

            A.Paragraph paragraph92 = new A.Paragraph();

            A.Field field16 = new A.Field() { Id = "{402C1264-1A6A-2D47-B539-28043A88F412}", Type = "slidenum" };

            A.RunProperties runProperties72 = new A.RunProperties() { Language = "en-US" };
            runProperties72.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text72 = new A.Text();
            text72.Text = "‹#›";

            field16.Append(runProperties72);
            field16.Append(text72);
            A.EndParagraphRunProperties endParagraphRunProperties69 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph92.Append(field16);
            paragraph92.Append(endParagraphRunProperties69);

            textBody76.Append(bodyProperties77);
            textBody76.Append(listStyle77);
            textBody76.Append(paragraph92);

            shape55.Append(nonVisualShapeProperties55);
            shape55.Append(shapeProperties64);
            shape55.Append(textBody76);

            shapeTree9.Append(nonVisualGroupShapeProperties15);
            shapeTree9.Append(groupShapeProperties15);
            shapeTree9.Append(shape51);
            shapeTree9.Append(shape52);
            shapeTree9.Append(shape53);
            shapeTree9.Append(shape54);
            shapeTree9.Append(shape55);

            CommonSlideDataExtensionList commonSlideDataExtensionList9 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension9 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId9 = new P14.CreationId() { Val = (UInt32Value)3581977006U };
            creationId9.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension9.Append(creationId9);

            commonSlideDataExtensionList9.Append(commonSlideDataExtension9);

            commonSlideData9.Append(shapeTree9);
            commonSlideData9.Append(commonSlideDataExtensionList9);

            ColorMapOverride colorMapOverride8 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping8 = new A.MasterColorMapping();

            colorMapOverride8.Append(masterColorMapping8);

            slideLayout7.Append(commonSlideData9);
            slideLayout7.Append(colorMapOverride8);

            slideLayoutPart7.SlideLayout = slideLayout7;
        }

        // Generates content of slideLayoutPart8.
        private void GenerateSlideLayoutPart8Content(SlideLayoutPart slideLayoutPart8)
        {
            SlideLayout slideLayout8 = new SlideLayout() { Type = SlideLayoutValues.TwoTextAndTwoObjects, Preserve = true };
            slideLayout8.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout8.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout8.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData10 = new CommonSlideData() { Name = "Comparison" };

            ShapeTree shapeTree10 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties16 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties80 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties16 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties80 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties16.Append(nonVisualDrawingProperties80);
            nonVisualGroupShapeProperties16.Append(nonVisualGroupShapeDrawingProperties16);
            nonVisualGroupShapeProperties16.Append(applicationNonVisualDrawingProperties80);

            GroupShapeProperties groupShapeProperties16 = new GroupShapeProperties();

            A.TransformGroup transformGroup16 = new A.TransformGroup();
            A.Offset offset54 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents54 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset16 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents16 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup16.Append(offset54);
            transformGroup16.Append(extents54);
            transformGroup16.Append(childOffset16);
            transformGroup16.Append(childExtents16);

            groupShapeProperties16.Append(transformGroup16);

            Shape shape56 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties56 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties81 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList71 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension71 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement91 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{4285DEE6-90C0-B7C2-FEAE-DC9881417269}\" />");

            // nonVisualDrawingPropertiesExtension71.Append(openXmlUnknownElement91);

            nonVisualDrawingPropertiesExtensionList71.Append(nonVisualDrawingPropertiesExtension71);

            nonVisualDrawingProperties81.Append(nonVisualDrawingPropertiesExtensionList71);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties56 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks41 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties56.Append(shapeLocks41);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties81 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape41 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties81.Append(placeholderShape41);

            nonVisualShapeProperties56.Append(nonVisualDrawingProperties81);
            nonVisualShapeProperties56.Append(nonVisualShapeDrawingProperties56);
            nonVisualShapeProperties56.Append(applicationNonVisualDrawingProperties81);

            ShapeProperties shapeProperties65 = new ShapeProperties();

            A.Transform2D transform2D38 = new A.Transform2D();
            A.Offset offset55 = new A.Offset() { X = 839788L, Y = 365125L };
            A.Extents extents55 = new A.Extents() { Cx = 10515600L, Cy = 1325563L };

            transform2D38.Append(offset55);
            transform2D38.Append(extents55);

            shapeProperties65.Append(transform2D38);

            TextBody textBody77 = new TextBody();
            A.BodyProperties bodyProperties78 = new A.BodyProperties();
            A.ListStyle listStyle78 = new A.ListStyle();

            A.Paragraph paragraph93 = new A.Paragraph();

            A.Run run57 = new A.Run();
            A.RunProperties runProperties73 = new A.RunProperties() { Language = "en-GB" };
            A.Text text73 = new A.Text();
            text73.Text = "Click to edit Master title style";

            run57.Append(runProperties73);
            run57.Append(text73);
            A.EndParagraphRunProperties endParagraphRunProperties70 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph93.Append(run57);
            paragraph93.Append(endParagraphRunProperties70);

            textBody77.Append(bodyProperties78);
            textBody77.Append(listStyle78);
            textBody77.Append(paragraph93);

            shape56.Append(nonVisualShapeProperties56);
            shape56.Append(shapeProperties65);
            shape56.Append(textBody77);

            Shape shape57 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties57 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties82 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Text Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList72 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension72 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement92 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{68D72E56-10D3-8D48-B9C7-4CF0FB45D369}\" />");

            // nonVisualDrawingPropertiesExtension72.Append(openXmlUnknownElement92);

            nonVisualDrawingPropertiesExtensionList72.Append(nonVisualDrawingPropertiesExtension72);

            nonVisualDrawingProperties82.Append(nonVisualDrawingPropertiesExtensionList72);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties57 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks42 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties57.Append(shapeLocks42);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties82 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape42 = new PlaceholderShape() { Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties82.Append(placeholderShape42);

            nonVisualShapeProperties57.Append(nonVisualDrawingProperties82);
            nonVisualShapeProperties57.Append(nonVisualShapeDrawingProperties57);
            nonVisualShapeProperties57.Append(applicationNonVisualDrawingProperties82);

            ShapeProperties shapeProperties66 = new ShapeProperties();

            A.Transform2D transform2D39 = new A.Transform2D();
            A.Offset offset56 = new A.Offset() { X = 839788L, Y = 1681163L };
            A.Extents extents56 = new A.Extents() { Cx = 5157787L, Cy = 823912L };

            transform2D39.Append(offset56);
            transform2D39.Append(extents56);

            shapeProperties66.Append(transform2D39);

            TextBody textBody78 = new TextBody();
            A.BodyProperties bodyProperties79 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle79 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties15 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet43 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties88 = new A.DefaultRunProperties() { FontSize = 2400, Bold = true };

            level1ParagraphProperties15.Append(noBullet43);
            level1ParagraphProperties15.Append(defaultRunProperties88);

            A.Level2ParagraphProperties level2ParagraphProperties8 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet44 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties89 = new A.DefaultRunProperties() { FontSize = 2000, Bold = true };

            level2ParagraphProperties8.Append(noBullet44);
            level2ParagraphProperties8.Append(defaultRunProperties89);

            A.Level3ParagraphProperties level3ParagraphProperties8 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet45 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties90 = new A.DefaultRunProperties() { FontSize = 1800, Bold = true };

            level3ParagraphProperties8.Append(noBullet45);
            level3ParagraphProperties8.Append(defaultRunProperties90);

            A.Level4ParagraphProperties level4ParagraphProperties8 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet46 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties91 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level4ParagraphProperties8.Append(noBullet46);
            level4ParagraphProperties8.Append(defaultRunProperties91);

            A.Level5ParagraphProperties level5ParagraphProperties8 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet47 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties92 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level5ParagraphProperties8.Append(noBullet47);
            level5ParagraphProperties8.Append(defaultRunProperties92);

            A.Level6ParagraphProperties level6ParagraphProperties8 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet48 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties93 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level6ParagraphProperties8.Append(noBullet48);
            level6ParagraphProperties8.Append(defaultRunProperties93);

            A.Level7ParagraphProperties level7ParagraphProperties8 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet49 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties94 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level7ParagraphProperties8.Append(noBullet49);
            level7ParagraphProperties8.Append(defaultRunProperties94);

            A.Level8ParagraphProperties level8ParagraphProperties8 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet50 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties95 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level8ParagraphProperties8.Append(noBullet50);
            level8ParagraphProperties8.Append(defaultRunProperties95);

            A.Level9ParagraphProperties level9ParagraphProperties8 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet51 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties96 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level9ParagraphProperties8.Append(noBullet51);
            level9ParagraphProperties8.Append(defaultRunProperties96);

            listStyle79.Append(level1ParagraphProperties15);
            listStyle79.Append(level2ParagraphProperties8);
            listStyle79.Append(level3ParagraphProperties8);
            listStyle79.Append(level4ParagraphProperties8);
            listStyle79.Append(level5ParagraphProperties8);
            listStyle79.Append(level6ParagraphProperties8);
            listStyle79.Append(level7ParagraphProperties8);
            listStyle79.Append(level8ParagraphProperties8);
            listStyle79.Append(level9ParagraphProperties8);

            A.Paragraph paragraph94 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties55 = new A.ParagraphProperties() { Level = 0 };

            A.Run run58 = new A.Run();
            A.RunProperties runProperties74 = new A.RunProperties() { Language = "en-GB" };
            A.Text text74 = new A.Text();
            text74.Text = "Click to edit Master text styles";

            run58.Append(runProperties74);
            run58.Append(text74);

            paragraph94.Append(paragraphProperties55);
            paragraph94.Append(run58);

            textBody78.Append(bodyProperties79);
            textBody78.Append(listStyle79);
            textBody78.Append(paragraph94);

            shape57.Append(nonVisualShapeProperties57);
            shape57.Append(shapeProperties66);
            shape57.Append(textBody78);

            Shape shape58 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties58 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties83 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Content Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList73 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension73 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement93 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{AA730891-13A1-B3EE-33C6-64B2E7965C9D}\" />");

            // nonVisualDrawingPropertiesExtension73.Append(openXmlUnknownElement93);

            nonVisualDrawingPropertiesExtensionList73.Append(nonVisualDrawingPropertiesExtension73);

            nonVisualDrawingProperties83.Append(nonVisualDrawingPropertiesExtensionList73);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties58 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks43 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties58.Append(shapeLocks43);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties83 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape43 = new PlaceholderShape() { Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties83.Append(placeholderShape43);

            nonVisualShapeProperties58.Append(nonVisualDrawingProperties83);
            nonVisualShapeProperties58.Append(nonVisualShapeDrawingProperties58);
            nonVisualShapeProperties58.Append(applicationNonVisualDrawingProperties83);

            ShapeProperties shapeProperties67 = new ShapeProperties();

            A.Transform2D transform2D40 = new A.Transform2D();
            A.Offset offset57 = new A.Offset() { X = 839788L, Y = 2505075L };
            A.Extents extents57 = new A.Extents() { Cx = 5157787L, Cy = 3684588L };

            transform2D40.Append(offset57);
            transform2D40.Append(extents57);

            shapeProperties67.Append(transform2D40);

            TextBody textBody79 = new TextBody();
            A.BodyProperties bodyProperties80 = new A.BodyProperties();
            A.ListStyle listStyle80 = new A.ListStyle();

            A.Paragraph paragraph95 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties56 = new A.ParagraphProperties() { Level = 0 };

            A.Run run59 = new A.Run();
            A.RunProperties runProperties75 = new A.RunProperties() { Language = "en-GB" };
            A.Text text75 = new A.Text();
            text75.Text = "Click to edit Master text styles";

            run59.Append(runProperties75);
            run59.Append(text75);

            paragraph95.Append(paragraphProperties56);
            paragraph95.Append(run59);

            A.Paragraph paragraph96 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties57 = new A.ParagraphProperties() { Level = 1 };

            A.Run run60 = new A.Run();
            A.RunProperties runProperties76 = new A.RunProperties() { Language = "en-GB" };
            A.Text text76 = new A.Text();
            text76.Text = "Second level";

            run60.Append(runProperties76);
            run60.Append(text76);

            paragraph96.Append(paragraphProperties57);
            paragraph96.Append(run60);

            A.Paragraph paragraph97 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties58 = new A.ParagraphProperties() { Level = 2 };

            A.Run run61 = new A.Run();
            A.RunProperties runProperties77 = new A.RunProperties() { Language = "en-GB" };
            A.Text text77 = new A.Text();
            text77.Text = "Third level";

            run61.Append(runProperties77);
            run61.Append(text77);

            paragraph97.Append(paragraphProperties58);
            paragraph97.Append(run61);

            A.Paragraph paragraph98 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties59 = new A.ParagraphProperties() { Level = 3 };

            A.Run run62 = new A.Run();
            A.RunProperties runProperties78 = new A.RunProperties() { Language = "en-GB" };
            A.Text text78 = new A.Text();
            text78.Text = "Fourth level";

            run62.Append(runProperties78);
            run62.Append(text78);

            paragraph98.Append(paragraphProperties59);
            paragraph98.Append(run62);

            A.Paragraph paragraph99 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties60 = new A.ParagraphProperties() { Level = 4 };

            A.Run run63 = new A.Run();
            A.RunProperties runProperties79 = new A.RunProperties() { Language = "en-GB" };
            A.Text text79 = new A.Text();
            text79.Text = "Fifth level";

            run63.Append(runProperties79);
            run63.Append(text79);
            A.EndParagraphRunProperties endParagraphRunProperties71 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph99.Append(paragraphProperties60);
            paragraph99.Append(run63);
            paragraph99.Append(endParagraphRunProperties71);

            textBody79.Append(bodyProperties80);
            textBody79.Append(listStyle80);
            textBody79.Append(paragraph95);
            textBody79.Append(paragraph96);
            textBody79.Append(paragraph97);
            textBody79.Append(paragraph98);
            textBody79.Append(paragraph99);

            shape58.Append(nonVisualShapeProperties58);
            shape58.Append(shapeProperties67);
            shape58.Append(textBody79);

            Shape shape59 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties59 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties84 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Text Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList74 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension74 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement94 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{3121263C-6EDA-7FD1-ABE9-E6F72646397D}\" />");

            // nonVisualDrawingPropertiesExtension74.Append(openXmlUnknownElement94);

            nonVisualDrawingPropertiesExtensionList74.Append(nonVisualDrawingPropertiesExtension74);

            nonVisualDrawingProperties84.Append(nonVisualDrawingPropertiesExtensionList74);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties59 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks44 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties59.Append(shapeLocks44);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties84 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape44 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)3U };

            applicationNonVisualDrawingProperties84.Append(placeholderShape44);

            nonVisualShapeProperties59.Append(nonVisualDrawingProperties84);
            nonVisualShapeProperties59.Append(nonVisualShapeDrawingProperties59);
            nonVisualShapeProperties59.Append(applicationNonVisualDrawingProperties84);

            ShapeProperties shapeProperties68 = new ShapeProperties();

            A.Transform2D transform2D41 = new A.Transform2D();
            A.Offset offset58 = new A.Offset() { X = 6172200L, Y = 1681163L };
            A.Extents extents58 = new A.Extents() { Cx = 5183188L, Cy = 823912L };

            transform2D41.Append(offset58);
            transform2D41.Append(extents58);

            shapeProperties68.Append(transform2D41);

            TextBody textBody80 = new TextBody();
            A.BodyProperties bodyProperties81 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle81 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties16 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet52 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties97 = new A.DefaultRunProperties() { FontSize = 2400, Bold = true };

            level1ParagraphProperties16.Append(noBullet52);
            level1ParagraphProperties16.Append(defaultRunProperties97);

            A.Level2ParagraphProperties level2ParagraphProperties9 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet53 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties98 = new A.DefaultRunProperties() { FontSize = 2000, Bold = true };

            level2ParagraphProperties9.Append(noBullet53);
            level2ParagraphProperties9.Append(defaultRunProperties98);

            A.Level3ParagraphProperties level3ParagraphProperties9 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet54 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties99 = new A.DefaultRunProperties() { FontSize = 1800, Bold = true };

            level3ParagraphProperties9.Append(noBullet54);
            level3ParagraphProperties9.Append(defaultRunProperties99);

            A.Level4ParagraphProperties level4ParagraphProperties9 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet55 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties100 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level4ParagraphProperties9.Append(noBullet55);
            level4ParagraphProperties9.Append(defaultRunProperties100);

            A.Level5ParagraphProperties level5ParagraphProperties9 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet56 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties101 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level5ParagraphProperties9.Append(noBullet56);
            level5ParagraphProperties9.Append(defaultRunProperties101);

            A.Level6ParagraphProperties level6ParagraphProperties9 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet57 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties102 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level6ParagraphProperties9.Append(noBullet57);
            level6ParagraphProperties9.Append(defaultRunProperties102);

            A.Level7ParagraphProperties level7ParagraphProperties9 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet58 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties103 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level7ParagraphProperties9.Append(noBullet58);
            level7ParagraphProperties9.Append(defaultRunProperties103);

            A.Level8ParagraphProperties level8ParagraphProperties9 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet59 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties104 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level8ParagraphProperties9.Append(noBullet59);
            level8ParagraphProperties9.Append(defaultRunProperties104);

            A.Level9ParagraphProperties level9ParagraphProperties9 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet60 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties105 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level9ParagraphProperties9.Append(noBullet60);
            level9ParagraphProperties9.Append(defaultRunProperties105);

            listStyle81.Append(level1ParagraphProperties16);
            listStyle81.Append(level2ParagraphProperties9);
            listStyle81.Append(level3ParagraphProperties9);
            listStyle81.Append(level4ParagraphProperties9);
            listStyle81.Append(level5ParagraphProperties9);
            listStyle81.Append(level6ParagraphProperties9);
            listStyle81.Append(level7ParagraphProperties9);
            listStyle81.Append(level8ParagraphProperties9);
            listStyle81.Append(level9ParagraphProperties9);

            A.Paragraph paragraph100 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties61 = new A.ParagraphProperties() { Level = 0 };

            A.Run run64 = new A.Run();
            A.RunProperties runProperties80 = new A.RunProperties() { Language = "en-GB" };
            A.Text text80 = new A.Text();
            text80.Text = "Click to edit Master text styles";

            run64.Append(runProperties80);
            run64.Append(text80);

            paragraph100.Append(paragraphProperties61);
            paragraph100.Append(run64);

            textBody80.Append(bodyProperties81);
            textBody80.Append(listStyle81);
            textBody80.Append(paragraph100);

            shape59.Append(nonVisualShapeProperties59);
            shape59.Append(shapeProperties68);
            shape59.Append(textBody80);

            Shape shape60 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties60 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties85 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Content Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList75 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension75 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement95 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{0E13E803-9700-065C-EB0F-BCE3D19BC1D1}\" />");

            // nonVisualDrawingPropertiesExtension75.Append(openXmlUnknownElement95);

            nonVisualDrawingPropertiesExtensionList75.Append(nonVisualDrawingPropertiesExtension75);

            nonVisualDrawingProperties85.Append(nonVisualDrawingPropertiesExtensionList75);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties60 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks45 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties60.Append(shapeLocks45);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties85 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape45 = new PlaceholderShape() { Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)4U };

            applicationNonVisualDrawingProperties85.Append(placeholderShape45);

            nonVisualShapeProperties60.Append(nonVisualDrawingProperties85);
            nonVisualShapeProperties60.Append(nonVisualShapeDrawingProperties60);
            nonVisualShapeProperties60.Append(applicationNonVisualDrawingProperties85);

            ShapeProperties shapeProperties69 = new ShapeProperties();

            A.Transform2D transform2D42 = new A.Transform2D();
            A.Offset offset59 = new A.Offset() { X = 6172200L, Y = 2505075L };
            A.Extents extents59 = new A.Extents() { Cx = 5183188L, Cy = 3684588L };

            transform2D42.Append(offset59);
            transform2D42.Append(extents59);

            shapeProperties69.Append(transform2D42);

            TextBody textBody81 = new TextBody();
            A.BodyProperties bodyProperties82 = new A.BodyProperties();
            A.ListStyle listStyle82 = new A.ListStyle();

            A.Paragraph paragraph101 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties62 = new A.ParagraphProperties() { Level = 0 };

            A.Run run65 = new A.Run();
            A.RunProperties runProperties81 = new A.RunProperties() { Language = "en-GB" };
            A.Text text81 = new A.Text();
            text81.Text = "Click to edit Master text styles";

            run65.Append(runProperties81);
            run65.Append(text81);

            paragraph101.Append(paragraphProperties62);
            paragraph101.Append(run65);

            A.Paragraph paragraph102 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties63 = new A.ParagraphProperties() { Level = 1 };

            A.Run run66 = new A.Run();
            A.RunProperties runProperties82 = new A.RunProperties() { Language = "en-GB" };
            A.Text text82 = new A.Text();
            text82.Text = "Second level";

            run66.Append(runProperties82);
            run66.Append(text82);

            paragraph102.Append(paragraphProperties63);
            paragraph102.Append(run66);

            A.Paragraph paragraph103 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties64 = new A.ParagraphProperties() { Level = 2 };

            A.Run run67 = new A.Run();
            A.RunProperties runProperties83 = new A.RunProperties() { Language = "en-GB" };
            A.Text text83 = new A.Text();
            text83.Text = "Third level";

            run67.Append(runProperties83);
            run67.Append(text83);

            paragraph103.Append(paragraphProperties64);
            paragraph103.Append(run67);

            A.Paragraph paragraph104 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties65 = new A.ParagraphProperties() { Level = 3 };

            A.Run run68 = new A.Run();
            A.RunProperties runProperties84 = new A.RunProperties() { Language = "en-GB" };
            A.Text text84 = new A.Text();
            text84.Text = "Fourth level";

            run68.Append(runProperties84);
            run68.Append(text84);

            paragraph104.Append(paragraphProperties65);
            paragraph104.Append(run68);

            A.Paragraph paragraph105 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties66 = new A.ParagraphProperties() { Level = 4 };

            A.Run run69 = new A.Run();
            A.RunProperties runProperties85 = new A.RunProperties() { Language = "en-GB" };
            A.Text text85 = new A.Text();
            text85.Text = "Fifth level";

            run69.Append(runProperties85);
            run69.Append(text85);
            A.EndParagraphRunProperties endParagraphRunProperties72 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph105.Append(paragraphProperties66);
            paragraph105.Append(run69);
            paragraph105.Append(endParagraphRunProperties72);

            textBody81.Append(bodyProperties82);
            textBody81.Append(listStyle82);
            textBody81.Append(paragraph101);
            textBody81.Append(paragraph102);
            textBody81.Append(paragraph103);
            textBody81.Append(paragraph104);
            textBody81.Append(paragraph105);

            shape60.Append(nonVisualShapeProperties60);
            shape60.Append(shapeProperties69);
            shape60.Append(textBody81);

            Shape shape61 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties61 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties86 = new NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Date Placeholder 6" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList76 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension76 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement96 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{C6182576-5F99-F491-A555-E50378601773}\" />");

            // nonVisualDrawingPropertiesExtension76.Append(openXmlUnknownElement96);

            nonVisualDrawingPropertiesExtensionList76.Append(nonVisualDrawingPropertiesExtension76);

            nonVisualDrawingProperties86.Append(nonVisualDrawingPropertiesExtensionList76);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties61 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks46 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties61.Append(shapeLocks46);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties86 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape46 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties86.Append(placeholderShape46);

            nonVisualShapeProperties61.Append(nonVisualDrawingProperties86);
            nonVisualShapeProperties61.Append(nonVisualShapeDrawingProperties61);
            nonVisualShapeProperties61.Append(applicationNonVisualDrawingProperties86);
            ShapeProperties shapeProperties70 = new ShapeProperties();

            TextBody textBody82 = new TextBody();
            A.BodyProperties bodyProperties83 = new A.BodyProperties();
            A.ListStyle listStyle83 = new A.ListStyle();

            A.Paragraph paragraph106 = new A.Paragraph();

            A.Field field17 = new A.Field() { Id = "{C453C8C6-607B-FF49-A93F-867C282599C0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties86 = new A.RunProperties() { Language = "en-US" };
            runProperties86.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text86 = new A.Text();
            text86.Text = "5/3/2024";

            field17.Append(runProperties86);
            field17.Append(text86);
            A.EndParagraphRunProperties endParagraphRunProperties73 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph106.Append(field17);
            paragraph106.Append(endParagraphRunProperties73);

            textBody82.Append(bodyProperties83);
            textBody82.Append(listStyle83);
            textBody82.Append(paragraph106);

            shape61.Append(nonVisualShapeProperties61);
            shape61.Append(shapeProperties70);
            shape61.Append(textBody82);

            Shape shape62 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties62 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties87 = new NonVisualDrawingProperties() { Id = (UInt32Value)8U, Name = "Footer Placeholder 7" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList77 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension77 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement97 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{7563219F-397D-90F8-F392-0467B2FB79E6}\" />");

            // nonVisualDrawingPropertiesExtension77.Append(openXmlUnknownElement97);

            nonVisualDrawingPropertiesExtensionList77.Append(nonVisualDrawingPropertiesExtension77);

            nonVisualDrawingProperties87.Append(nonVisualDrawingPropertiesExtensionList77);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties62 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks47 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties62.Append(shapeLocks47);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties87 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape47 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties87.Append(placeholderShape47);

            nonVisualShapeProperties62.Append(nonVisualDrawingProperties87);
            nonVisualShapeProperties62.Append(nonVisualShapeDrawingProperties62);
            nonVisualShapeProperties62.Append(applicationNonVisualDrawingProperties87);
            ShapeProperties shapeProperties71 = new ShapeProperties();

            TextBody textBody83 = new TextBody();
            A.BodyProperties bodyProperties84 = new A.BodyProperties();
            A.ListStyle listStyle84 = new A.ListStyle();

            A.Paragraph paragraph107 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties74 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph107.Append(endParagraphRunProperties74);

            textBody83.Append(bodyProperties84);
            textBody83.Append(listStyle84);
            textBody83.Append(paragraph107);

            shape62.Append(nonVisualShapeProperties62);
            shape62.Append(shapeProperties71);
            shape62.Append(textBody83);

            Shape shape63 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties63 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties88 = new NonVisualDrawingProperties() { Id = (UInt32Value)9U, Name = "Slide Number Placeholder 8" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList78 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension78 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement98 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{BDBA2A99-3015-7601-BF15-66BFF870DA91}\" />");

            // nonVisualDrawingPropertiesExtension78.Append(openXmlUnknownElement98);

            nonVisualDrawingPropertiesExtensionList78.Append(nonVisualDrawingPropertiesExtension78);

            nonVisualDrawingProperties88.Append(nonVisualDrawingPropertiesExtensionList78);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties63 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks48 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties63.Append(shapeLocks48);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties88 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape48 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties88.Append(placeholderShape48);

            nonVisualShapeProperties63.Append(nonVisualDrawingProperties88);
            nonVisualShapeProperties63.Append(nonVisualShapeDrawingProperties63);
            nonVisualShapeProperties63.Append(applicationNonVisualDrawingProperties88);
            ShapeProperties shapeProperties72 = new ShapeProperties();

            TextBody textBody84 = new TextBody();
            A.BodyProperties bodyProperties85 = new A.BodyProperties();
            A.ListStyle listStyle85 = new A.ListStyle();

            A.Paragraph paragraph108 = new A.Paragraph();

            A.Field field18 = new A.Field() { Id = "{402C1264-1A6A-2D47-B539-28043A88F412}", Type = "slidenum" };

            A.RunProperties runProperties87 = new A.RunProperties() { Language = "en-US" };
            runProperties87.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text87 = new A.Text();
            text87.Text = "‹#›";

            field18.Append(runProperties87);
            field18.Append(text87);
            A.EndParagraphRunProperties endParagraphRunProperties75 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph108.Append(field18);
            paragraph108.Append(endParagraphRunProperties75);

            textBody84.Append(bodyProperties85);
            textBody84.Append(listStyle85);
            textBody84.Append(paragraph108);

            shape63.Append(nonVisualShapeProperties63);
            shape63.Append(shapeProperties72);
            shape63.Append(textBody84);

            shapeTree10.Append(nonVisualGroupShapeProperties16);
            shapeTree10.Append(groupShapeProperties16);
            shapeTree10.Append(shape56);
            shapeTree10.Append(shape57);
            shapeTree10.Append(shape58);
            shapeTree10.Append(shape59);
            shapeTree10.Append(shape60);
            shapeTree10.Append(shape61);
            shapeTree10.Append(shape62);
            shapeTree10.Append(shape63);

            CommonSlideDataExtensionList commonSlideDataExtensionList10 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension10 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId10 = new P14.CreationId() { Val = (UInt32Value)2372890705U };
            creationId10.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension10.Append(creationId10);

            commonSlideDataExtensionList10.Append(commonSlideDataExtension10);

            commonSlideData10.Append(shapeTree10);
            commonSlideData10.Append(commonSlideDataExtensionList10);

            ColorMapOverride colorMapOverride9 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping9 = new A.MasterColorMapping();

            colorMapOverride9.Append(masterColorMapping9);

            slideLayout8.Append(commonSlideData10);
            slideLayout8.Append(colorMapOverride9);

            slideLayoutPart8.SlideLayout = slideLayout8;
        }

        // Generates content of slideLayoutPart9.
        private void GenerateSlideLayoutPart9Content(SlideLayoutPart slideLayoutPart9)
        {
            SlideLayout slideLayout9 = new SlideLayout() { Type = SlideLayoutValues.VerticalText, Preserve = true };
            slideLayout9.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout9.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout9.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData11 = new CommonSlideData() { Name = "Title and Vertical Text" };

            ShapeTree shapeTree11 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties17 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties89 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties17 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties89 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties17.Append(nonVisualDrawingProperties89);
            nonVisualGroupShapeProperties17.Append(nonVisualGroupShapeDrawingProperties17);
            nonVisualGroupShapeProperties17.Append(applicationNonVisualDrawingProperties89);

            GroupShapeProperties groupShapeProperties17 = new GroupShapeProperties();

            A.TransformGroup transformGroup17 = new A.TransformGroup();
            A.Offset offset60 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents60 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset17 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents17 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup17.Append(offset60);
            transformGroup17.Append(extents60);
            transformGroup17.Append(childOffset17);
            transformGroup17.Append(childExtents17);

            groupShapeProperties17.Append(transformGroup17);

            Shape shape64 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties64 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties90 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList79 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension79 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement99 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{C87C6705-2FD8-45CD-A0D6-ECE47E8C63C4}\" />");

            // nonVisualDrawingPropertiesExtension79.Append(openXmlUnknownElement99);

            nonVisualDrawingPropertiesExtensionList79.Append(nonVisualDrawingPropertiesExtension79);

            nonVisualDrawingProperties90.Append(nonVisualDrawingPropertiesExtensionList79);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties64 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks49 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties64.Append(shapeLocks49);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties90 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape49 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties90.Append(placeholderShape49);

            nonVisualShapeProperties64.Append(nonVisualDrawingProperties90);
            nonVisualShapeProperties64.Append(nonVisualShapeDrawingProperties64);
            nonVisualShapeProperties64.Append(applicationNonVisualDrawingProperties90);
            ShapeProperties shapeProperties73 = new ShapeProperties();

            TextBody textBody85 = new TextBody();
            A.BodyProperties bodyProperties86 = new A.BodyProperties();
            A.ListStyle listStyle86 = new A.ListStyle();

            A.Paragraph paragraph109 = new A.Paragraph();

            A.Run run70 = new A.Run();
            A.RunProperties runProperties88 = new A.RunProperties() { Language = "en-GB" };
            A.Text text88 = new A.Text();
            text88.Text = "Click to edit Master title style";

            run70.Append(runProperties88);
            run70.Append(text88);
            A.EndParagraphRunProperties endParagraphRunProperties76 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph109.Append(run70);
            paragraph109.Append(endParagraphRunProperties76);

            textBody85.Append(bodyProperties86);
            textBody85.Append(listStyle86);
            textBody85.Append(paragraph109);

            shape64.Append(nonVisualShapeProperties64);
            shape64.Append(shapeProperties73);
            shape64.Append(textBody85);

            Shape shape65 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties65 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties91 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Vertical Text Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList80 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension80 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement100 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{8179E5F4-2046-C56C-D4FA-8027BE0A6ADB}\" />");

            // nonVisualDrawingPropertiesExtension80.Append(openXmlUnknownElement100);

            nonVisualDrawingPropertiesExtensionList80.Append(nonVisualDrawingPropertiesExtension80);

            nonVisualDrawingProperties91.Append(nonVisualDrawingPropertiesExtensionList80);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties65 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks50 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties65.Append(shapeLocks50);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties91 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape50 = new PlaceholderShape() { Type = PlaceholderValues.Body, Orientation = DirectionValues.Vertical, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties91.Append(placeholderShape50);

            nonVisualShapeProperties65.Append(nonVisualDrawingProperties91);
            nonVisualShapeProperties65.Append(nonVisualShapeDrawingProperties65);
            nonVisualShapeProperties65.Append(applicationNonVisualDrawingProperties91);
            ShapeProperties shapeProperties74 = new ShapeProperties();

            TextBody textBody86 = new TextBody();
            A.BodyProperties bodyProperties87 = new A.BodyProperties() { Vertical = A.TextVerticalValues.EastAsianVetical };
            A.ListStyle listStyle87 = new A.ListStyle();

            A.Paragraph paragraph110 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties67 = new A.ParagraphProperties() { Level = 0 };

            A.Run run71 = new A.Run();
            A.RunProperties runProperties89 = new A.RunProperties() { Language = "en-GB" };
            A.Text text89 = new A.Text();
            text89.Text = "Click to edit Master text styles";

            run71.Append(runProperties89);
            run71.Append(text89);

            paragraph110.Append(paragraphProperties67);
            paragraph110.Append(run71);

            A.Paragraph paragraph111 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties68 = new A.ParagraphProperties() { Level = 1 };

            A.Run run72 = new A.Run();
            A.RunProperties runProperties90 = new A.RunProperties() { Language = "en-GB" };
            A.Text text90 = new A.Text();
            text90.Text = "Second level";

            run72.Append(runProperties90);
            run72.Append(text90);

            paragraph111.Append(paragraphProperties68);
            paragraph111.Append(run72);

            A.Paragraph paragraph112 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties69 = new A.ParagraphProperties() { Level = 2 };

            A.Run run73 = new A.Run();
            A.RunProperties runProperties91 = new A.RunProperties() { Language = "en-GB" };
            A.Text text91 = new A.Text();
            text91.Text = "Third level";

            run73.Append(runProperties91);
            run73.Append(text91);

            paragraph112.Append(paragraphProperties69);
            paragraph112.Append(run73);

            A.Paragraph paragraph113 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties70 = new A.ParagraphProperties() { Level = 3 };

            A.Run run74 = new A.Run();
            A.RunProperties runProperties92 = new A.RunProperties() { Language = "en-GB" };
            A.Text text92 = new A.Text();
            text92.Text = "Fourth level";

            run74.Append(runProperties92);
            run74.Append(text92);

            paragraph113.Append(paragraphProperties70);
            paragraph113.Append(run74);

            A.Paragraph paragraph114 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties71 = new A.ParagraphProperties() { Level = 4 };

            A.Run run75 = new A.Run();
            A.RunProperties runProperties93 = new A.RunProperties() { Language = "en-GB" };
            A.Text text93 = new A.Text();
            text93.Text = "Fifth level";

            run75.Append(runProperties93);
            run75.Append(text93);
            A.EndParagraphRunProperties endParagraphRunProperties77 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph114.Append(paragraphProperties71);
            paragraph114.Append(run75);
            paragraph114.Append(endParagraphRunProperties77);

            textBody86.Append(bodyProperties87);
            textBody86.Append(listStyle87);
            textBody86.Append(paragraph110);
            textBody86.Append(paragraph111);
            textBody86.Append(paragraph112);
            textBody86.Append(paragraph113);
            textBody86.Append(paragraph114);

            shape65.Append(nonVisualShapeProperties65);
            shape65.Append(shapeProperties74);
            shape65.Append(textBody86);

            Shape shape66 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties66 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties92 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList81 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension81 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement101 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{7A744EC2-06A0-778B-1F1C-7A83D83324F0}\" />");

            // nonVisualDrawingPropertiesExtension81.Append(openXmlUnknownElement101);

            nonVisualDrawingPropertiesExtensionList81.Append(nonVisualDrawingPropertiesExtension81);

            nonVisualDrawingProperties92.Append(nonVisualDrawingPropertiesExtensionList81);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties66 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks51 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties66.Append(shapeLocks51);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties92 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape51 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties92.Append(placeholderShape51);

            nonVisualShapeProperties66.Append(nonVisualDrawingProperties92);
            nonVisualShapeProperties66.Append(nonVisualShapeDrawingProperties66);
            nonVisualShapeProperties66.Append(applicationNonVisualDrawingProperties92);
            ShapeProperties shapeProperties75 = new ShapeProperties();

            TextBody textBody87 = new TextBody();
            A.BodyProperties bodyProperties88 = new A.BodyProperties();
            A.ListStyle listStyle88 = new A.ListStyle();

            A.Paragraph paragraph115 = new A.Paragraph();

            A.Field field19 = new A.Field() { Id = "{C453C8C6-607B-FF49-A93F-867C282599C0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties94 = new A.RunProperties() { Language = "en-US" };
            runProperties94.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text94 = new A.Text();
            text94.Text = "5/3/2024";

            field19.Append(runProperties94);
            field19.Append(text94);
            A.EndParagraphRunProperties endParagraphRunProperties78 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph115.Append(field19);
            paragraph115.Append(endParagraphRunProperties78);

            textBody87.Append(bodyProperties88);
            textBody87.Append(listStyle88);
            textBody87.Append(paragraph115);

            shape66.Append(nonVisualShapeProperties66);
            shape66.Append(shapeProperties75);
            shape66.Append(textBody87);

            Shape shape67 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties67 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties93 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList82 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension82 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement102 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{464D367A-8325-3C86-DF1A-E0B0A0A6242A}\" />");

            // nonVisualDrawingPropertiesExtension82.Append(openXmlUnknownElement102);

            nonVisualDrawingPropertiesExtensionList82.Append(nonVisualDrawingPropertiesExtension82);

            nonVisualDrawingProperties93.Append(nonVisualDrawingPropertiesExtensionList82);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties67 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks52 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties67.Append(shapeLocks52);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties93 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape52 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties93.Append(placeholderShape52);

            nonVisualShapeProperties67.Append(nonVisualDrawingProperties93);
            nonVisualShapeProperties67.Append(nonVisualShapeDrawingProperties67);
            nonVisualShapeProperties67.Append(applicationNonVisualDrawingProperties93);
            ShapeProperties shapeProperties76 = new ShapeProperties();

            TextBody textBody88 = new TextBody();
            A.BodyProperties bodyProperties89 = new A.BodyProperties();
            A.ListStyle listStyle89 = new A.ListStyle();

            A.Paragraph paragraph116 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties79 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph116.Append(endParagraphRunProperties79);

            textBody88.Append(bodyProperties89);
            textBody88.Append(listStyle89);
            textBody88.Append(paragraph116);

            shape67.Append(nonVisualShapeProperties67);
            shape67.Append(shapeProperties76);
            shape67.Append(textBody88);

            Shape shape68 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties68 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties94 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList83 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension83 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement103 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{CE459884-73BB-9CD0-5533-F01C170E8520}\" />");

            // nonVisualDrawingPropertiesExtension83.Append(openXmlUnknownElement103);

            nonVisualDrawingPropertiesExtensionList83.Append(nonVisualDrawingPropertiesExtension83);

            nonVisualDrawingProperties94.Append(nonVisualDrawingPropertiesExtensionList83);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties68 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks53 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties68.Append(shapeLocks53);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties94 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape53 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties94.Append(placeholderShape53);

            nonVisualShapeProperties68.Append(nonVisualDrawingProperties94);
            nonVisualShapeProperties68.Append(nonVisualShapeDrawingProperties68);
            nonVisualShapeProperties68.Append(applicationNonVisualDrawingProperties94);
            ShapeProperties shapeProperties77 = new ShapeProperties();

            TextBody textBody89 = new TextBody();
            A.BodyProperties bodyProperties90 = new A.BodyProperties();
            A.ListStyle listStyle90 = new A.ListStyle();

            A.Paragraph paragraph117 = new A.Paragraph();

            A.Field field20 = new A.Field() { Id = "{402C1264-1A6A-2D47-B539-28043A88F412}", Type = "slidenum" };

            A.RunProperties runProperties95 = new A.RunProperties() { Language = "en-US" };
            runProperties95.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text95 = new A.Text();
            text95.Text = "‹#›";

            field20.Append(runProperties95);
            field20.Append(text95);
            A.EndParagraphRunProperties endParagraphRunProperties80 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph117.Append(field20);
            paragraph117.Append(endParagraphRunProperties80);

            textBody89.Append(bodyProperties90);
            textBody89.Append(listStyle90);
            textBody89.Append(paragraph117);

            shape68.Append(nonVisualShapeProperties68);
            shape68.Append(shapeProperties77);
            shape68.Append(textBody89);

            shapeTree11.Append(nonVisualGroupShapeProperties17);
            shapeTree11.Append(groupShapeProperties17);
            shapeTree11.Append(shape64);
            shapeTree11.Append(shape65);
            shapeTree11.Append(shape66);
            shapeTree11.Append(shape67);
            shapeTree11.Append(shape68);

            CommonSlideDataExtensionList commonSlideDataExtensionList11 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension11 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId11 = new P14.CreationId() { Val = (UInt32Value)3629721967U };
            creationId11.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension11.Append(creationId11);

            commonSlideDataExtensionList11.Append(commonSlideDataExtension11);

            commonSlideData11.Append(shapeTree11);
            commonSlideData11.Append(commonSlideDataExtensionList11);

            ColorMapOverride colorMapOverride10 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping10 = new A.MasterColorMapping();

            colorMapOverride10.Append(masterColorMapping10);

            slideLayout9.Append(commonSlideData11);
            slideLayout9.Append(colorMapOverride10);

            slideLayoutPart9.SlideLayout = slideLayout9;
        }

        // Generates content of slideLayoutPart10.
        private void GenerateSlideLayoutPart10Content(SlideLayoutPart slideLayoutPart10)
        {
            SlideLayout slideLayout10 = new SlideLayout() { Type = SlideLayoutValues.TwoObjects, Preserve = true };
            slideLayout10.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout10.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout10.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData12 = new CommonSlideData() { Name = "Two Content" };

            ShapeTree shapeTree12 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties18 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties95 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties18 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties95 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties18.Append(nonVisualDrawingProperties95);
            nonVisualGroupShapeProperties18.Append(nonVisualGroupShapeDrawingProperties18);
            nonVisualGroupShapeProperties18.Append(applicationNonVisualDrawingProperties95);

            GroupShapeProperties groupShapeProperties18 = new GroupShapeProperties();

            A.TransformGroup transformGroup18 = new A.TransformGroup();
            A.Offset offset61 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents61 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset18 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents18 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup18.Append(offset61);
            transformGroup18.Append(extents61);
            transformGroup18.Append(childOffset18);
            transformGroup18.Append(childExtents18);

            groupShapeProperties18.Append(transformGroup18);

            Shape shape69 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties69 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties96 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList84 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension84 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement104 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{44E88A92-7A78-455F-9E80-A8452842A9DD}\" />");

            // nonVisualDrawingPropertiesExtension84.Append(openXmlUnknownElement104);

            nonVisualDrawingPropertiesExtensionList84.Append(nonVisualDrawingPropertiesExtension84);

            nonVisualDrawingProperties96.Append(nonVisualDrawingPropertiesExtensionList84);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties69 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks54 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties69.Append(shapeLocks54);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties96 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape54 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties96.Append(placeholderShape54);

            nonVisualShapeProperties69.Append(nonVisualDrawingProperties96);
            nonVisualShapeProperties69.Append(nonVisualShapeDrawingProperties69);
            nonVisualShapeProperties69.Append(applicationNonVisualDrawingProperties96);
            ShapeProperties shapeProperties78 = new ShapeProperties();

            TextBody textBody90 = new TextBody();
            A.BodyProperties bodyProperties91 = new A.BodyProperties();
            A.ListStyle listStyle91 = new A.ListStyle();

            A.Paragraph paragraph118 = new A.Paragraph();

            A.Run run76 = new A.Run();
            A.RunProperties runProperties96 = new A.RunProperties() { Language = "en-GB" };
            A.Text text96 = new A.Text();
            text96.Text = "Click to edit Master title style";

            run76.Append(runProperties96);
            run76.Append(text96);
            A.EndParagraphRunProperties endParagraphRunProperties81 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph118.Append(run76);
            paragraph118.Append(endParagraphRunProperties81);

            textBody90.Append(bodyProperties91);
            textBody90.Append(listStyle91);
            textBody90.Append(paragraph118);

            shape69.Append(nonVisualShapeProperties69);
            shape69.Append(shapeProperties78);
            shape69.Append(textBody90);

            Shape shape70 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties70 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties97 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Content Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList85 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension85 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement105 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{3BC1C7DE-9CBB-DAD1-0942-64ED8F99D659}\" />");

            // nonVisualDrawingPropertiesExtension85.Append(openXmlUnknownElement105);

            nonVisualDrawingPropertiesExtensionList85.Append(nonVisualDrawingPropertiesExtension85);

            nonVisualDrawingProperties97.Append(nonVisualDrawingPropertiesExtensionList85);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties70 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks55 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties70.Append(shapeLocks55);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties97 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape55 = new PlaceholderShape() { Size = PlaceholderSizeValues.Half, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties97.Append(placeholderShape55);

            nonVisualShapeProperties70.Append(nonVisualDrawingProperties97);
            nonVisualShapeProperties70.Append(nonVisualShapeDrawingProperties70);
            nonVisualShapeProperties70.Append(applicationNonVisualDrawingProperties97);

            ShapeProperties shapeProperties79 = new ShapeProperties();

            A.Transform2D transform2D43 = new A.Transform2D();
            A.Offset offset62 = new A.Offset() { X = 838200L, Y = 1825625L };
            A.Extents extents62 = new A.Extents() { Cx = 5181600L, Cy = 4351338L };

            transform2D43.Append(offset62);
            transform2D43.Append(extents62);

            shapeProperties79.Append(transform2D43);

            TextBody textBody91 = new TextBody();
            A.BodyProperties bodyProperties92 = new A.BodyProperties();
            A.ListStyle listStyle92 = new A.ListStyle();

            A.Paragraph paragraph119 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties72 = new A.ParagraphProperties() { Level = 0 };

            A.Run run77 = new A.Run();
            A.RunProperties runProperties97 = new A.RunProperties() { Language = "en-GB" };
            A.Text text97 = new A.Text();
            text97.Text = "Click to edit Master text styles";

            run77.Append(runProperties97);
            run77.Append(text97);

            paragraph119.Append(paragraphProperties72);
            paragraph119.Append(run77);

            A.Paragraph paragraph120 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties73 = new A.ParagraphProperties() { Level = 1 };

            A.Run run78 = new A.Run();
            A.RunProperties runProperties98 = new A.RunProperties() { Language = "en-GB" };
            A.Text text98 = new A.Text();
            text98.Text = "Second level";

            run78.Append(runProperties98);
            run78.Append(text98);

            paragraph120.Append(paragraphProperties73);
            paragraph120.Append(run78);

            A.Paragraph paragraph121 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties74 = new A.ParagraphProperties() { Level = 2 };

            A.Run run79 = new A.Run();
            A.RunProperties runProperties99 = new A.RunProperties() { Language = "en-GB" };
            A.Text text99 = new A.Text();
            text99.Text = "Third level";

            run79.Append(runProperties99);
            run79.Append(text99);

            paragraph121.Append(paragraphProperties74);
            paragraph121.Append(run79);

            A.Paragraph paragraph122 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties75 = new A.ParagraphProperties() { Level = 3 };

            A.Run run80 = new A.Run();
            A.RunProperties runProperties100 = new A.RunProperties() { Language = "en-GB" };
            A.Text text100 = new A.Text();
            text100.Text = "Fourth level";

            run80.Append(runProperties100);
            run80.Append(text100);

            paragraph122.Append(paragraphProperties75);
            paragraph122.Append(run80);

            A.Paragraph paragraph123 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties76 = new A.ParagraphProperties() { Level = 4 };

            A.Run run81 = new A.Run();
            A.RunProperties runProperties101 = new A.RunProperties() { Language = "en-GB" };
            A.Text text101 = new A.Text();
            text101.Text = "Fifth level";

            run81.Append(runProperties101);
            run81.Append(text101);
            A.EndParagraphRunProperties endParagraphRunProperties82 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph123.Append(paragraphProperties76);
            paragraph123.Append(run81);
            paragraph123.Append(endParagraphRunProperties82);

            textBody91.Append(bodyProperties92);
            textBody91.Append(listStyle92);
            textBody91.Append(paragraph119);
            textBody91.Append(paragraph120);
            textBody91.Append(paragraph121);
            textBody91.Append(paragraph122);
            textBody91.Append(paragraph123);

            shape70.Append(nonVisualShapeProperties70);
            shape70.Append(shapeProperties79);
            shape70.Append(textBody91);

            Shape shape71 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties71 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties98 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Content Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList86 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension86 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement106 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{D0357539-A53F-BDF1-2405-C97E0047074E}\" />");

            // nonVisualDrawingPropertiesExtension86.Append(openXmlUnknownElement106);

            nonVisualDrawingPropertiesExtensionList86.Append(nonVisualDrawingPropertiesExtension86);

            nonVisualDrawingProperties98.Append(nonVisualDrawingPropertiesExtensionList86);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties71 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks56 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties71.Append(shapeLocks56);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties98 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape56 = new PlaceholderShape() { Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties98.Append(placeholderShape56);

            nonVisualShapeProperties71.Append(nonVisualDrawingProperties98);
            nonVisualShapeProperties71.Append(nonVisualShapeDrawingProperties71);
            nonVisualShapeProperties71.Append(applicationNonVisualDrawingProperties98);

            ShapeProperties shapeProperties80 = new ShapeProperties();

            A.Transform2D transform2D44 = new A.Transform2D();
            A.Offset offset63 = new A.Offset() { X = 6172200L, Y = 1825625L };
            A.Extents extents63 = new A.Extents() { Cx = 5181600L, Cy = 4351338L };

            transform2D44.Append(offset63);
            transform2D44.Append(extents63);

            shapeProperties80.Append(transform2D44);

            TextBody textBody92 = new TextBody();
            A.BodyProperties bodyProperties93 = new A.BodyProperties();
            A.ListStyle listStyle93 = new A.ListStyle();

            A.Paragraph paragraph124 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties77 = new A.ParagraphProperties() { Level = 0 };

            A.Run run82 = new A.Run();
            A.RunProperties runProperties102 = new A.RunProperties() { Language = "en-GB" };
            A.Text text102 = new A.Text();
            text102.Text = "Click to edit Master text styles";

            run82.Append(runProperties102);
            run82.Append(text102);

            paragraph124.Append(paragraphProperties77);
            paragraph124.Append(run82);

            A.Paragraph paragraph125 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties78 = new A.ParagraphProperties() { Level = 1 };

            A.Run run83 = new A.Run();
            A.RunProperties runProperties103 = new A.RunProperties() { Language = "en-GB" };
            A.Text text103 = new A.Text();
            text103.Text = "Second level";

            run83.Append(runProperties103);
            run83.Append(text103);

            paragraph125.Append(paragraphProperties78);
            paragraph125.Append(run83);

            A.Paragraph paragraph126 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties79 = new A.ParagraphProperties() { Level = 2 };

            A.Run run84 = new A.Run();
            A.RunProperties runProperties104 = new A.RunProperties() { Language = "en-GB" };
            A.Text text104 = new A.Text();
            text104.Text = "Third level";

            run84.Append(runProperties104);
            run84.Append(text104);

            paragraph126.Append(paragraphProperties79);
            paragraph126.Append(run84);

            A.Paragraph paragraph127 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties80 = new A.ParagraphProperties() { Level = 3 };

            A.Run run85 = new A.Run();
            A.RunProperties runProperties105 = new A.RunProperties() { Language = "en-GB" };
            A.Text text105 = new A.Text();
            text105.Text = "Fourth level";

            run85.Append(runProperties105);
            run85.Append(text105);

            paragraph127.Append(paragraphProperties80);
            paragraph127.Append(run85);

            A.Paragraph paragraph128 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties81 = new A.ParagraphProperties() { Level = 4 };

            A.Run run86 = new A.Run();
            A.RunProperties runProperties106 = new A.RunProperties() { Language = "en-GB" };
            A.Text text106 = new A.Text();
            text106.Text = "Fifth level";

            run86.Append(runProperties106);
            run86.Append(text106);
            A.EndParagraphRunProperties endParagraphRunProperties83 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph128.Append(paragraphProperties81);
            paragraph128.Append(run86);
            paragraph128.Append(endParagraphRunProperties83);

            textBody92.Append(bodyProperties93);
            textBody92.Append(listStyle93);
            textBody92.Append(paragraph124);
            textBody92.Append(paragraph125);
            textBody92.Append(paragraph126);
            textBody92.Append(paragraph127);
            textBody92.Append(paragraph128);

            shape71.Append(nonVisualShapeProperties71);
            shape71.Append(shapeProperties80);
            shape71.Append(textBody92);

            Shape shape72 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties72 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties99 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Date Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList87 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension87 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement107 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{5189AE70-A178-119F-0A61-A0D319B113AB}\" />");

            // nonVisualDrawingPropertiesExtension87.Append(openXmlUnknownElement107);

            nonVisualDrawingPropertiesExtensionList87.Append(nonVisualDrawingPropertiesExtension87);

            nonVisualDrawingProperties99.Append(nonVisualDrawingPropertiesExtensionList87);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties72 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks57 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties72.Append(shapeLocks57);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties99 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape57 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties99.Append(placeholderShape57);

            nonVisualShapeProperties72.Append(nonVisualDrawingProperties99);
            nonVisualShapeProperties72.Append(nonVisualShapeDrawingProperties72);
            nonVisualShapeProperties72.Append(applicationNonVisualDrawingProperties99);
            ShapeProperties shapeProperties81 = new ShapeProperties();

            TextBody textBody93 = new TextBody();
            A.BodyProperties bodyProperties94 = new A.BodyProperties();
            A.ListStyle listStyle94 = new A.ListStyle();

            A.Paragraph paragraph129 = new A.Paragraph();

            A.Field field21 = new A.Field() { Id = "{C453C8C6-607B-FF49-A93F-867C282599C0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties107 = new A.RunProperties() { Language = "en-US" };
            runProperties107.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text107 = new A.Text();
            text107.Text = "5/3/2024";

            field21.Append(runProperties107);
            field21.Append(text107);
            A.EndParagraphRunProperties endParagraphRunProperties84 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph129.Append(field21);
            paragraph129.Append(endParagraphRunProperties84);

            textBody93.Append(bodyProperties94);
            textBody93.Append(listStyle94);
            textBody93.Append(paragraph129);

            shape72.Append(nonVisualShapeProperties72);
            shape72.Append(shapeProperties81);
            shape72.Append(textBody93);

            Shape shape73 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties73 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties100 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Footer Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList88 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension88 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement108 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{FED0CE86-19BA-0FFD-6BCE-1FF315FFF64F}\" />");

            // nonVisualDrawingPropertiesExtension88.Append(openXmlUnknownElement108);

            nonVisualDrawingPropertiesExtensionList88.Append(nonVisualDrawingPropertiesExtension88);

            nonVisualDrawingProperties100.Append(nonVisualDrawingPropertiesExtensionList88);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties73 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks58 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties73.Append(shapeLocks58);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties100 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape58 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties100.Append(placeholderShape58);

            nonVisualShapeProperties73.Append(nonVisualDrawingProperties100);
            nonVisualShapeProperties73.Append(nonVisualShapeDrawingProperties73);
            nonVisualShapeProperties73.Append(applicationNonVisualDrawingProperties100);
            ShapeProperties shapeProperties82 = new ShapeProperties();

            TextBody textBody94 = new TextBody();
            A.BodyProperties bodyProperties95 = new A.BodyProperties();
            A.ListStyle listStyle95 = new A.ListStyle();

            A.Paragraph paragraph130 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties85 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph130.Append(endParagraphRunProperties85);

            textBody94.Append(bodyProperties95);
            textBody94.Append(listStyle95);
            textBody94.Append(paragraph130);

            shape73.Append(nonVisualShapeProperties73);
            shape73.Append(shapeProperties82);
            shape73.Append(textBody94);

            Shape shape74 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties74 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties101 = new NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Slide Number Placeholder 6" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList89 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension89 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement109 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{4150BABB-CD69-E8C3-819B-4103D2A39BDF}\" />");

            // nonVisualDrawingPropertiesExtension89.Append(openXmlUnknownElement109);

            nonVisualDrawingPropertiesExtensionList89.Append(nonVisualDrawingPropertiesExtension89);

            nonVisualDrawingProperties101.Append(nonVisualDrawingPropertiesExtensionList89);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties74 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks59 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties74.Append(shapeLocks59);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties101 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape59 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties101.Append(placeholderShape59);

            nonVisualShapeProperties74.Append(nonVisualDrawingProperties101);
            nonVisualShapeProperties74.Append(nonVisualShapeDrawingProperties74);
            nonVisualShapeProperties74.Append(applicationNonVisualDrawingProperties101);
            ShapeProperties shapeProperties83 = new ShapeProperties();

            TextBody textBody95 = new TextBody();
            A.BodyProperties bodyProperties96 = new A.BodyProperties();
            A.ListStyle listStyle96 = new A.ListStyle();

            A.Paragraph paragraph131 = new A.Paragraph();

            A.Field field22 = new A.Field() { Id = "{402C1264-1A6A-2D47-B539-28043A88F412}", Type = "slidenum" };

            A.RunProperties runProperties108 = new A.RunProperties() { Language = "en-US" };
            runProperties108.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text108 = new A.Text();
            text108.Text = "‹#›";

            field22.Append(runProperties108);
            field22.Append(text108);
            A.EndParagraphRunProperties endParagraphRunProperties86 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph131.Append(field22);
            paragraph131.Append(endParagraphRunProperties86);

            textBody95.Append(bodyProperties96);
            textBody95.Append(listStyle96);
            textBody95.Append(paragraph131);

            shape74.Append(nonVisualShapeProperties74);
            shape74.Append(shapeProperties83);
            shape74.Append(textBody95);

            shapeTree12.Append(nonVisualGroupShapeProperties18);
            shapeTree12.Append(groupShapeProperties18);
            shapeTree12.Append(shape69);
            shapeTree12.Append(shape70);
            shapeTree12.Append(shape71);
            shapeTree12.Append(shape72);
            shapeTree12.Append(shape73);
            shapeTree12.Append(shape74);

            CommonSlideDataExtensionList commonSlideDataExtensionList12 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension12 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId12 = new P14.CreationId() { Val = (UInt32Value)487346656U };
            creationId12.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension12.Append(creationId12);

            commonSlideDataExtensionList12.Append(commonSlideDataExtension12);

            commonSlideData12.Append(shapeTree12);
            commonSlideData12.Append(commonSlideDataExtensionList12);

            ColorMapOverride colorMapOverride11 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping11 = new A.MasterColorMapping();

            colorMapOverride11.Append(masterColorMapping11);

            slideLayout10.Append(commonSlideData12);
            slideLayout10.Append(colorMapOverride11);

            slideLayoutPart10.SlideLayout = slideLayout10;
        }

        // Generates content of slideLayoutPart11.
        private void GenerateSlideLayoutPart11Content(SlideLayoutPart slideLayoutPart11)
        {
            SlideLayout slideLayout11 = new SlideLayout() { Type = SlideLayoutValues.PictureText, Preserve = true };
            slideLayout11.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout11.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout11.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData13 = new CommonSlideData() { Name = "Picture with Caption" };

            ShapeTree shapeTree13 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties19 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties102 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties19 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties102 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties19.Append(nonVisualDrawingProperties102);
            nonVisualGroupShapeProperties19.Append(nonVisualGroupShapeDrawingProperties19);
            nonVisualGroupShapeProperties19.Append(applicationNonVisualDrawingProperties102);

            GroupShapeProperties groupShapeProperties19 = new GroupShapeProperties();

            A.TransformGroup transformGroup19 = new A.TransformGroup();
            A.Offset offset64 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents64 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset19 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents19 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup19.Append(offset64);
            transformGroup19.Append(extents64);
            transformGroup19.Append(childOffset19);
            transformGroup19.Append(childExtents19);

            groupShapeProperties19.Append(transformGroup19);

            Shape shape75 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties75 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties103 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList90 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension90 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement110 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{601183E0-28E8-8038-5D32-86F9CC7982A9}\" />");

            // nonVisualDrawingPropertiesExtension90.Append(openXmlUnknownElement110);

            nonVisualDrawingPropertiesExtensionList90.Append(nonVisualDrawingPropertiesExtension90);

            nonVisualDrawingProperties103.Append(nonVisualDrawingPropertiesExtensionList90);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties75 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks60 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties75.Append(shapeLocks60);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties103 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape60 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties103.Append(placeholderShape60);

            nonVisualShapeProperties75.Append(nonVisualDrawingProperties103);
            nonVisualShapeProperties75.Append(nonVisualShapeDrawingProperties75);
            nonVisualShapeProperties75.Append(applicationNonVisualDrawingProperties103);

            ShapeProperties shapeProperties84 = new ShapeProperties();

            A.Transform2D transform2D45 = new A.Transform2D();
            A.Offset offset65 = new A.Offset() { X = 839788L, Y = 457200L };
            A.Extents extents65 = new A.Extents() { Cx = 3932237L, Cy = 1600200L };

            transform2D45.Append(offset65);
            transform2D45.Append(extents65);

            shapeProperties84.Append(transform2D45);

            TextBody textBody96 = new TextBody();
            A.BodyProperties bodyProperties97 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle97 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties17 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties106 = new A.DefaultRunProperties() { FontSize = 3200 };

            level1ParagraphProperties17.Append(defaultRunProperties106);

            listStyle97.Append(level1ParagraphProperties17);

            A.Paragraph paragraph132 = new A.Paragraph();

            A.Run run87 = new A.Run();
            A.RunProperties runProperties109 = new A.RunProperties() { Language = "en-GB" };
            A.Text text109 = new A.Text();
            text109.Text = "Click to edit Master title style";

            run87.Append(runProperties109);
            run87.Append(text109);
            A.EndParagraphRunProperties endParagraphRunProperties87 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph132.Append(run87);
            paragraph132.Append(endParagraphRunProperties87);

            textBody96.Append(bodyProperties97);
            textBody96.Append(listStyle97);
            textBody96.Append(paragraph132);

            shape75.Append(nonVisualShapeProperties75);
            shape75.Append(shapeProperties84);
            shape75.Append(textBody96);

            Shape shape76 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties76 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties104 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Picture Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList91 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension91 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement111 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{DC045E55-61E5-942B-5187-C0EB64307CEB}\" />");

            // nonVisualDrawingPropertiesExtension91.Append(openXmlUnknownElement111);

            nonVisualDrawingPropertiesExtensionList91.Append(nonVisualDrawingPropertiesExtension91);

            nonVisualDrawingProperties104.Append(nonVisualDrawingPropertiesExtensionList91);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties76 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks61 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties76.Append(shapeLocks61);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties104 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape61 = new PlaceholderShape() { Type = PlaceholderValues.Picture, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties104.Append(placeholderShape61);

            nonVisualShapeProperties76.Append(nonVisualDrawingProperties104);
            nonVisualShapeProperties76.Append(nonVisualShapeDrawingProperties76);
            nonVisualShapeProperties76.Append(applicationNonVisualDrawingProperties104);

            ShapeProperties shapeProperties85 = new ShapeProperties();

            A.Transform2D transform2D46 = new A.Transform2D();
            A.Offset offset66 = new A.Offset() { X = 5183188L, Y = 987425L };
            A.Extents extents66 = new A.Extents() { Cx = 6172200L, Cy = 4873625L };

            transform2D46.Append(offset66);
            transform2D46.Append(extents66);

            shapeProperties85.Append(transform2D46);

            TextBody textBody97 = new TextBody();
            A.BodyProperties bodyProperties98 = new A.BodyProperties();

            A.ListStyle listStyle98 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties18 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet61 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties107 = new A.DefaultRunProperties() { FontSize = 3200 };

            level1ParagraphProperties18.Append(noBullet61);
            level1ParagraphProperties18.Append(defaultRunProperties107);

            A.Level2ParagraphProperties level2ParagraphProperties10 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet62 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties108 = new A.DefaultRunProperties() { FontSize = 2800 };

            level2ParagraphProperties10.Append(noBullet62);
            level2ParagraphProperties10.Append(defaultRunProperties108);

            A.Level3ParagraphProperties level3ParagraphProperties10 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet63 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties109 = new A.DefaultRunProperties() { FontSize = 2400 };

            level3ParagraphProperties10.Append(noBullet63);
            level3ParagraphProperties10.Append(defaultRunProperties109);

            A.Level4ParagraphProperties level4ParagraphProperties10 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet64 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties110 = new A.DefaultRunProperties() { FontSize = 2000 };

            level4ParagraphProperties10.Append(noBullet64);
            level4ParagraphProperties10.Append(defaultRunProperties110);

            A.Level5ParagraphProperties level5ParagraphProperties10 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet65 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties111 = new A.DefaultRunProperties() { FontSize = 2000 };

            level5ParagraphProperties10.Append(noBullet65);
            level5ParagraphProperties10.Append(defaultRunProperties111);

            A.Level6ParagraphProperties level6ParagraphProperties10 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet66 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties112 = new A.DefaultRunProperties() { FontSize = 2000 };

            level6ParagraphProperties10.Append(noBullet66);
            level6ParagraphProperties10.Append(defaultRunProperties112);

            A.Level7ParagraphProperties level7ParagraphProperties10 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet67 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties113 = new A.DefaultRunProperties() { FontSize = 2000 };

            level7ParagraphProperties10.Append(noBullet67);
            level7ParagraphProperties10.Append(defaultRunProperties113);

            A.Level8ParagraphProperties level8ParagraphProperties10 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet68 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties114 = new A.DefaultRunProperties() { FontSize = 2000 };

            level8ParagraphProperties10.Append(noBullet68);
            level8ParagraphProperties10.Append(defaultRunProperties114);

            A.Level9ParagraphProperties level9ParagraphProperties10 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet69 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties115 = new A.DefaultRunProperties() { FontSize = 2000 };

            level9ParagraphProperties10.Append(noBullet69);
            level9ParagraphProperties10.Append(defaultRunProperties115);

            listStyle98.Append(level1ParagraphProperties18);
            listStyle98.Append(level2ParagraphProperties10);
            listStyle98.Append(level3ParagraphProperties10);
            listStyle98.Append(level4ParagraphProperties10);
            listStyle98.Append(level5ParagraphProperties10);
            listStyle98.Append(level6ParagraphProperties10);
            listStyle98.Append(level7ParagraphProperties10);
            listStyle98.Append(level8ParagraphProperties10);
            listStyle98.Append(level9ParagraphProperties10);

            A.Paragraph paragraph133 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties88 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph133.Append(endParagraphRunProperties88);

            textBody97.Append(bodyProperties98);
            textBody97.Append(listStyle98);
            textBody97.Append(paragraph133);

            shape76.Append(nonVisualShapeProperties76);
            shape76.Append(shapeProperties85);
            shape76.Append(textBody97);

            Shape shape77 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties77 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties105 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Text Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList92 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension92 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement112 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{D70E556F-5A06-A4AA-9D11-5480A55745A9}\" />");

            // nonVisualDrawingPropertiesExtension92.Append(openXmlUnknownElement112);

            nonVisualDrawingPropertiesExtensionList92.Append(nonVisualDrawingPropertiesExtension92);

            nonVisualDrawingProperties105.Append(nonVisualDrawingPropertiesExtensionList92);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties77 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks62 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties77.Append(shapeLocks62);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties105 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape62 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties105.Append(placeholderShape62);

            nonVisualShapeProperties77.Append(nonVisualDrawingProperties105);
            nonVisualShapeProperties77.Append(nonVisualShapeDrawingProperties77);
            nonVisualShapeProperties77.Append(applicationNonVisualDrawingProperties105);

            ShapeProperties shapeProperties86 = new ShapeProperties();

            A.Transform2D transform2D47 = new A.Transform2D();
            A.Offset offset67 = new A.Offset() { X = 839788L, Y = 2057400L };
            A.Extents extents67 = new A.Extents() { Cx = 3932237L, Cy = 3811588L };

            transform2D47.Append(offset67);
            transform2D47.Append(extents67);

            shapeProperties86.Append(transform2D47);

            TextBody textBody98 = new TextBody();
            A.BodyProperties bodyProperties99 = new A.BodyProperties();

            A.ListStyle listStyle99 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties19 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet70 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties116 = new A.DefaultRunProperties() { FontSize = 1600 };

            level1ParagraphProperties19.Append(noBullet70);
            level1ParagraphProperties19.Append(defaultRunProperties116);

            A.Level2ParagraphProperties level2ParagraphProperties11 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet71 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties117 = new A.DefaultRunProperties() { FontSize = 1400 };

            level2ParagraphProperties11.Append(noBullet71);
            level2ParagraphProperties11.Append(defaultRunProperties117);

            A.Level3ParagraphProperties level3ParagraphProperties11 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet72 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties118 = new A.DefaultRunProperties() { FontSize = 1200 };

            level3ParagraphProperties11.Append(noBullet72);
            level3ParagraphProperties11.Append(defaultRunProperties118);

            A.Level4ParagraphProperties level4ParagraphProperties11 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet73 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties119 = new A.DefaultRunProperties() { FontSize = 1000 };

            level4ParagraphProperties11.Append(noBullet73);
            level4ParagraphProperties11.Append(defaultRunProperties119);

            A.Level5ParagraphProperties level5ParagraphProperties11 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet74 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties120 = new A.DefaultRunProperties() { FontSize = 1000 };

            level5ParagraphProperties11.Append(noBullet74);
            level5ParagraphProperties11.Append(defaultRunProperties120);

            A.Level6ParagraphProperties level6ParagraphProperties11 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet75 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties121 = new A.DefaultRunProperties() { FontSize = 1000 };

            level6ParagraphProperties11.Append(noBullet75);
            level6ParagraphProperties11.Append(defaultRunProperties121);

            A.Level7ParagraphProperties level7ParagraphProperties11 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet76 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties122 = new A.DefaultRunProperties() { FontSize = 1000 };

            level7ParagraphProperties11.Append(noBullet76);
            level7ParagraphProperties11.Append(defaultRunProperties122);

            A.Level8ParagraphProperties level8ParagraphProperties11 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet77 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties123 = new A.DefaultRunProperties() { FontSize = 1000 };

            level8ParagraphProperties11.Append(noBullet77);
            level8ParagraphProperties11.Append(defaultRunProperties123);

            A.Level9ParagraphProperties level9ParagraphProperties11 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet78 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties124 = new A.DefaultRunProperties() { FontSize = 1000 };

            level9ParagraphProperties11.Append(noBullet78);
            level9ParagraphProperties11.Append(defaultRunProperties124);

            listStyle99.Append(level1ParagraphProperties19);
            listStyle99.Append(level2ParagraphProperties11);
            listStyle99.Append(level3ParagraphProperties11);
            listStyle99.Append(level4ParagraphProperties11);
            listStyle99.Append(level5ParagraphProperties11);
            listStyle99.Append(level6ParagraphProperties11);
            listStyle99.Append(level7ParagraphProperties11);
            listStyle99.Append(level8ParagraphProperties11);
            listStyle99.Append(level9ParagraphProperties11);

            A.Paragraph paragraph134 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties82 = new A.ParagraphProperties() { Level = 0 };

            A.Run run88 = new A.Run();
            A.RunProperties runProperties110 = new A.RunProperties() { Language = "en-GB" };
            A.Text text110 = new A.Text();
            text110.Text = "Click to edit Master text styles";

            run88.Append(runProperties110);
            run88.Append(text110);

            paragraph134.Append(paragraphProperties82);
            paragraph134.Append(run88);

            textBody98.Append(bodyProperties99);
            textBody98.Append(listStyle99);
            textBody98.Append(paragraph134);

            shape77.Append(nonVisualShapeProperties77);
            shape77.Append(shapeProperties86);
            shape77.Append(textBody98);

            Shape shape78 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties78 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties106 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Date Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList93 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension93 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement113 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{B8C63D13-50D9-677E-9DC7-C27611FD6956}\" />");

            // nonVisualDrawingPropertiesExtension93.Append(openXmlUnknownElement113);

            nonVisualDrawingPropertiesExtensionList93.Append(nonVisualDrawingPropertiesExtension93);

            nonVisualDrawingProperties106.Append(nonVisualDrawingPropertiesExtensionList93);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties78 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks63 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties78.Append(shapeLocks63);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties106 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape63 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties106.Append(placeholderShape63);

            nonVisualShapeProperties78.Append(nonVisualDrawingProperties106);
            nonVisualShapeProperties78.Append(nonVisualShapeDrawingProperties78);
            nonVisualShapeProperties78.Append(applicationNonVisualDrawingProperties106);
            ShapeProperties shapeProperties87 = new ShapeProperties();

            TextBody textBody99 = new TextBody();
            A.BodyProperties bodyProperties100 = new A.BodyProperties();
            A.ListStyle listStyle100 = new A.ListStyle();

            A.Paragraph paragraph135 = new A.Paragraph();

            A.Field field23 = new A.Field() { Id = "{C453C8C6-607B-FF49-A93F-867C282599C0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties111 = new A.RunProperties() { Language = "en-US" };
            runProperties111.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text111 = new A.Text();
            text111.Text = "5/3/2024";

            field23.Append(runProperties111);
            field23.Append(text111);
            A.EndParagraphRunProperties endParagraphRunProperties89 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph135.Append(field23);
            paragraph135.Append(endParagraphRunProperties89);

            textBody99.Append(bodyProperties100);
            textBody99.Append(listStyle100);
            textBody99.Append(paragraph135);

            shape78.Append(nonVisualShapeProperties78);
            shape78.Append(shapeProperties87);
            shape78.Append(textBody99);

            Shape shape79 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties79 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties107 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Footer Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList94 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension94 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement114 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{91F6F58D-A31D-7D06-D289-47B1C8411146}\" />");

            // nonVisualDrawingPropertiesExtension94.Append(openXmlUnknownElement114);

            nonVisualDrawingPropertiesExtensionList94.Append(nonVisualDrawingPropertiesExtension94);

            nonVisualDrawingProperties107.Append(nonVisualDrawingPropertiesExtensionList94);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties79 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks64 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties79.Append(shapeLocks64);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties107 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape64 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties107.Append(placeholderShape64);

            nonVisualShapeProperties79.Append(nonVisualDrawingProperties107);
            nonVisualShapeProperties79.Append(nonVisualShapeDrawingProperties79);
            nonVisualShapeProperties79.Append(applicationNonVisualDrawingProperties107);
            ShapeProperties shapeProperties88 = new ShapeProperties();

            TextBody textBody100 = new TextBody();
            A.BodyProperties bodyProperties101 = new A.BodyProperties();
            A.ListStyle listStyle101 = new A.ListStyle();

            A.Paragraph paragraph136 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties90 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph136.Append(endParagraphRunProperties90);

            textBody100.Append(bodyProperties101);
            textBody100.Append(listStyle101);
            textBody100.Append(paragraph136);

            shape79.Append(nonVisualShapeProperties79);
            shape79.Append(shapeProperties88);
            shape79.Append(textBody100);

            Shape shape80 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties80 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties108 = new NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Slide Number Placeholder 6" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList95 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension95 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement115 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{99BA17FA-8DCC-B5C4-5646-0E6248B09F69}\" />");

            // nonVisualDrawingPropertiesExtension95.Append(openXmlUnknownElement115);

            nonVisualDrawingPropertiesExtensionList95.Append(nonVisualDrawingPropertiesExtension95);

            nonVisualDrawingProperties108.Append(nonVisualDrawingPropertiesExtensionList95);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties80 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks65 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties80.Append(shapeLocks65);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties108 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape65 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties108.Append(placeholderShape65);

            nonVisualShapeProperties80.Append(nonVisualDrawingProperties108);
            nonVisualShapeProperties80.Append(nonVisualShapeDrawingProperties80);
            nonVisualShapeProperties80.Append(applicationNonVisualDrawingProperties108);
            ShapeProperties shapeProperties89 = new ShapeProperties();

            TextBody textBody101 = new TextBody();
            A.BodyProperties bodyProperties102 = new A.BodyProperties();
            A.ListStyle listStyle102 = new A.ListStyle();

            A.Paragraph paragraph137 = new A.Paragraph();

            A.Field field24 = new A.Field() { Id = "{402C1264-1A6A-2D47-B539-28043A88F412}", Type = "slidenum" };

            A.RunProperties runProperties112 = new A.RunProperties() { Language = "en-US" };
            runProperties112.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text112 = new A.Text();
            text112.Text = "‹#›";

            field24.Append(runProperties112);
            field24.Append(text112);
            A.EndParagraphRunProperties endParagraphRunProperties91 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph137.Append(field24);
            paragraph137.Append(endParagraphRunProperties91);

            textBody101.Append(bodyProperties102);
            textBody101.Append(listStyle102);
            textBody101.Append(paragraph137);

            shape80.Append(nonVisualShapeProperties80);
            shape80.Append(shapeProperties89);
            shape80.Append(textBody101);

            shapeTree13.Append(nonVisualGroupShapeProperties19);
            shapeTree13.Append(groupShapeProperties19);
            shapeTree13.Append(shape75);
            shapeTree13.Append(shape76);
            shapeTree13.Append(shape77);
            shapeTree13.Append(shape78);
            shapeTree13.Append(shape79);
            shapeTree13.Append(shape80);

            CommonSlideDataExtensionList commonSlideDataExtensionList13 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension13 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId13 = new P14.CreationId() { Val = (UInt32Value)122466947U };
            creationId13.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension13.Append(creationId13);

            commonSlideDataExtensionList13.Append(commonSlideDataExtension13);

            commonSlideData13.Append(shapeTree13);
            commonSlideData13.Append(commonSlideDataExtensionList13);

            ColorMapOverride colorMapOverride12 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping12 = new A.MasterColorMapping();

            colorMapOverride12.Append(masterColorMapping12);

            slideLayout11.Append(commonSlideData13);
            slideLayout11.Append(colorMapOverride12);

            slideLayoutPart11.SlideLayout = slideLayout11;
        }

        // Generates content of imagePart4.
        private void GenerateImagePart4Content(ImagePart imagePart4)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart4Data);
            imagePart4.FeedData(data);
            data.Close();
        }

        // Generates content of imagePart5.
        private void GenerateImagePart5Content(ImagePart imagePart5)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart5Data);
            imagePart5.FeedData(data);
            data.Close();
        }

        // Generates content of imagePart6.
        private void GenerateImagePart6Content(ImagePart imagePart6)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart6Data);
            imagePart6.FeedData(data);
            data.Close();
        }

        // Generates content of tableStylesPart1.
        private void GenerateTableStylesPart1Content(TableStylesPart tableStylesPart1)
        {
            A.TableStyleList tableStyleList1 = new A.TableStyleList() { Default = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" };
            tableStyleList1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            tableStylesPart1.TableStyleList = tableStyleList1;
        }

        // Generates content of slidePart2.

        // Generates content of viewPropertiesPart1.
        private void GenerateViewPropertiesPart1Content(ViewPropertiesPart viewPropertiesPart1)
        {
            ViewProperties viewProperties1 = new ViewProperties() { LastView = ViewValues.SlideThumbnailView };
            viewProperties1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            viewProperties1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            viewProperties1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            NormalViewProperties normalViewProperties1 = new NormalViewProperties();
            RestoredLeft restoredLeft1 = new RestoredLeft() { Size = 15604 };
            RestoredTop restoredTop1 = new RestoredTop() { Size = 94694 };

            normalViewProperties1.Append(restoredLeft1);
            normalViewProperties1.Append(restoredTop1);

            SlideViewProperties slideViewProperties1 = new SlideViewProperties();

            CommonSlideViewProperties commonSlideViewProperties1 = new CommonSlideViewProperties() { SnapToGrid = false };

            CommonViewProperties commonViewProperties1 = new CommonViewProperties() { VariableScale = true };

            ScaleFactor scaleFactor1 = new ScaleFactor();
            A.ScaleX scaleX1 = new A.ScaleX() { Numerator = 104, Denominator = 100 };
            A.ScaleY scaleY1 = new A.ScaleY() { Numerator = 104, Denominator = 100 };

            scaleFactor1.Append(scaleX1);
            scaleFactor1.Append(scaleY1);
            Origin origin1 = new Origin() { X = 348L, Y = 114L };

            commonViewProperties1.Append(scaleFactor1);
            commonViewProperties1.Append(origin1);
            GuideList guideList1 = new GuideList();

            commonSlideViewProperties1.Append(commonViewProperties1);
            commonSlideViewProperties1.Append(guideList1);

            slideViewProperties1.Append(commonSlideViewProperties1);

            NotesTextViewProperties notesTextViewProperties1 = new NotesTextViewProperties();

            CommonViewProperties commonViewProperties2 = new CommonViewProperties();

            ScaleFactor scaleFactor2 = new ScaleFactor();
            A.ScaleX scaleX2 = new A.ScaleX() { Numerator = 1, Denominator = 1 };
            A.ScaleY scaleY2 = new A.ScaleY() { Numerator = 1, Denominator = 1 };

            scaleFactor2.Append(scaleX2);
            scaleFactor2.Append(scaleY2);
            Origin origin2 = new Origin() { X = 0L, Y = 0L };

            commonViewProperties2.Append(scaleFactor2);
            commonViewProperties2.Append(origin2);

            notesTextViewProperties1.Append(commonViewProperties2);
            GridSpacing gridSpacing1 = new GridSpacing() { Cx = 72008L, Cy = 72008L };

            viewProperties1.Append(normalViewProperties1);
            viewProperties1.Append(slideViewProperties1);
            viewProperties1.Append(notesTextViewProperties1);
            viewProperties1.Append(gridSpacing1);

            viewPropertiesPart1.ViewProperties = viewProperties1;
        }

        // Generates content of presentationPropertiesPart1.
        private void GeneratePresentationPropertiesPart1Content(PresentationPropertiesPart presentationPropertiesPart1)
        {
            PresentationProperties presentationProperties1 = new PresentationProperties();
            presentationProperties1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            presentationProperties1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            presentationProperties1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            ShowProperties showProperties1 = new ShowProperties() { ShowNarration = true };
            PresenterSlideMode presenterSlideMode1 = new PresenterSlideMode();
            SlideAll slideAll1 = new SlideAll();

            PenColor penColor1 = new PenColor();
            A.PresetColor presetColor17 = new A.PresetColor() { Val = A.PresetColorValues.Red };

            penColor1.Append(presetColor17);

            ShowPropertiesExtensionList showPropertiesExtensionList1 = new ShowPropertiesExtensionList();

            ShowPropertiesExtension showPropertiesExtension1 = new ShowPropertiesExtension() { Uri = "{EC167BDD-8182-4AB7-AECC-EB403E3ABB37}" };

            P14.LaserColor laserColor1 = new P14.LaserColor();
            laserColor1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");
            A.RgbColorModelHex rgbColorModelHex64 = new A.RgbColorModelHex() { Val = "FF0000" };

            laserColor1.Append(rgbColorModelHex64);

            showPropertiesExtension1.Append(laserColor1);

            ShowPropertiesExtension showPropertiesExtension2 = new ShowPropertiesExtension() { Uri = "{2FDB2607-1784-4EEB-B798-7EB5836EED8A}" };

            P14.ShowMediaControls showMediaControls1 = new P14.ShowMediaControls() { Val = true };
            showMediaControls1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            showPropertiesExtension2.Append(showMediaControls1);

            showPropertiesExtensionList1.Append(showPropertiesExtension1);
            showPropertiesExtensionList1.Append(showPropertiesExtension2);

            showProperties1.Append(presenterSlideMode1);
            showProperties1.Append(slideAll1);
            showProperties1.Append(penColor1);
            showProperties1.Append(showPropertiesExtensionList1);

            PresentationPropertiesExtensionList presentationPropertiesExtensionList1 = new PresentationPropertiesExtensionList();

            PresentationPropertiesExtension presentationPropertiesExtension1 = new PresentationPropertiesExtension() { Uri = "{E76CE94A-603C-4142-B9EB-6D1370010A27}" };

            P14.DiscardImageEditData discardImageEditData1 = new P14.DiscardImageEditData() { Val = false };
            discardImageEditData1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            presentationPropertiesExtension1.Append(discardImageEditData1);

            PresentationPropertiesExtension presentationPropertiesExtension2 = new PresentationPropertiesExtension() { Uri = "{D31A062A-798A-4329-ABDD-BBA856620510}" };

            P14.DefaultImageDpi defaultImageDpi1 = new P14.DefaultImageDpi() { Val = (UInt32Value)32767U };
            defaultImageDpi1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            presentationPropertiesExtension2.Append(defaultImageDpi1);

            PresentationPropertiesExtension presentationPropertiesExtension3 = new PresentationPropertiesExtension() { Uri = "{FD5EFAAD-0ECE-453E-9831-46B23BE46B34}" };

            P15.ChartTrackingReferenceBased chartTrackingReferenceBased1 = new P15.ChartTrackingReferenceBased() { Val = true };
            chartTrackingReferenceBased1.AddNamespaceDeclaration("p15", "http://schemas.microsoft.com/office/powerpoint/2012/main");

            presentationPropertiesExtension3.Append(chartTrackingReferenceBased1);

            presentationPropertiesExtensionList1.Append(presentationPropertiesExtension1);
            presentationPropertiesExtensionList1.Append(presentationPropertiesExtension2);
            presentationPropertiesExtensionList1.Append(presentationPropertiesExtension3);

            presentationProperties1.Append(showProperties1);
            presentationProperties1.Append(presentationPropertiesExtensionList1);

            presentationPropertiesPart1.PresentationProperties = presentationProperties1;
        }

        // Generates content of extendedPart1.
        private void GenerateExtendedPart1Content(ExtendedPart extendedPart1)
        {
            System.IO.Stream data = GetBinaryDataStream(extendedPart1Data);
            extendedPart1.FeedData(data);
            data.Close();
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.TotalTime totalTime1 = new Ap.TotalTime();
            totalTime1.Text = "63";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "310";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office PowerPoint";
            Ap.PresentationFormat presentationFormat1 = new Ap.PresentationFormat();
            presentationFormat1.Text = "Widescreen";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "46";
            Ap.Slides slides1 = new Ap.Slides();
            slides1.Text = "2";
            Ap.Notes notes1 = new Ap.Notes();
            notes1.Text = "0";
            Ap.HiddenSlides hiddenSlides1 = new Ap.HiddenSlides();
            hiddenSlides1.Text = "0";
            Ap.MultimediaClips multimediaClips1 = new Ap.MultimediaClips();
            multimediaClips1.Text = "0";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)6U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Fonts Used";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "6";

            variant2.Append(vTInt321);

            Vt.Variant variant3 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Theme";

            variant3.Append(vTLPSTR2);

            Vt.Variant variant4 = new Vt.Variant();
            Vt.VTInt32 vTInt322 = new Vt.VTInt32();
            vTInt322.Text = "1";

            variant4.Append(vTInt322);

            Vt.Variant variant5 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR3 = new Vt.VTLPSTR();
            vTLPSTR3.Text = "Slide Titles";

            variant5.Append(vTLPSTR3);

            Vt.Variant variant6 = new Vt.Variant();
            Vt.VTInt32 vTInt323 = new Vt.VTInt32();
            vTInt323.Text = "2";

            variant6.Append(vTInt323);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);
            vTVector1.Append(variant3);
            vTVector1.Append(variant4);
            vTVector1.Append(variant5);
            vTVector1.Append(variant6);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)9U };
            Vt.VTLPSTR vTLPSTR4 = new Vt.VTLPSTR();
            vTLPSTR4.Text = "Aptos";
            Vt.VTLPSTR vTLPSTR5 = new Vt.VTLPSTR();
            vTLPSTR5.Text = "Aptos Display";
            Vt.VTLPSTR vTLPSTR6 = new Vt.VTLPSTR();
            vTLPSTR6.Text = "Arial";
            Vt.VTLPSTR vTLPSTR7 = new Vt.VTLPSTR();
            vTLPSTR7.Text = "Segoe UI";
            Vt.VTLPSTR vTLPSTR8 = new Vt.VTLPSTR();
            vTLPSTR8.Text = "Segoe UI Regular";
            Vt.VTLPSTR vTLPSTR9 = new Vt.VTLPSTR();
            vTLPSTR9.Text = "Segoe UI Semibold";
            Vt.VTLPSTR vTLPSTR10 = new Vt.VTLPSTR();
            vTLPSTR10.Text = "Office Theme";
            Vt.VTLPSTR vTLPSTR11 = new Vt.VTLPSTR();
            vTLPSTR11.Text = "PowerPoint Presentation";
            Vt.VTLPSTR vTLPSTR12 = new Vt.VTLPSTR();
            vTLPSTR12.Text = "PowerPoint Presentation";

            vTVector2.Append(vTLPSTR4);
            vTVector2.Append(vTLPSTR5);
            vTVector2.Append(vTLPSTR6);
            vTVector2.Append(vTLPSTR7);
            vTVector2.Append(vTLPSTR8);
            vTVector2.Append(vTLPSTR9);
            vTVector2.Append(vTLPSTR10);
            vTVector2.Append(vTLPSTR11);
            vTVector2.Append(vTLPSTR12);

            titlesOfParts1.Append(vTVector2);
            Ap.Company company1 = new Ap.Company();
            company1.Text = "";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "16.0000";

            properties1.Append(totalTime1);
            properties1.Append(words1);
            properties1.Append(application1);
            properties1.Append(presentationFormat1);
            properties1.Append(paragraphs1);
            properties1.Append(slides1);
            properties1.Append(notes1);
            properties1.Append(hiddenSlides1);
            properties1.Append(multimediaClips1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Pankaj Kumar";
            document.PackageProperties.Title = "PowerPoint Presentation";
            document.PackageProperties.Revision = "2";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2024-05-02T11:39:38Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2024-05-03T07:23:35Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Pankaj Kumar";
        }

        #region Binary Data
        private string thumbnailPart1Data = "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCACQAQADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9K/iJ8RtE+F/hyfWteuGgtI1YgIu5pGCltq+5xxnFaPg/xHF4w8JaJr0EMltBqljBfRwzY3xrLGrhWx3AbBqfXPD+l+JrE2WsabZ6rZlg5t76BJoyw6HawIyKreLPEEPgvwnqmstZzXcOm2klx9ktAvmSBFJ2ICQMnGBkgVTat5nNGNb20pSkuTSytrfrd318tjZpCa8t0/8AaQ8GX8OpXH2i7itLOaSMXH2SSRZVSBJTIuwEgHcyqGwWZDgHjMF/+0h4as2u9trqTx28NzJ5rWrIrPDDbymPn7pIuUXLY+YEehqTouer0hNee6z8btC07RF1S0gvNUt9lxK4hRYiiQRCWcnzWQFkDbSgO7cGUgbG2x/8L+8GPfGyj1C4lvRJFCbeOymaQSSOI0QgLwd5KH+6yspwRimQ2eiFqYzV5lN+0V4Khm2vqMgja0W8jZbWZmdMurkKE/gZdrdw24EDaa63wn400nxxp819o9y11bRTG3aQxsnzgKSBuAzjcPxyO1MybN0tTd1NLVGWpmTkSFqbuqNmpm73pmbkS7/ek3ioS9J5lMjmJt9HmVX8z3pPM96QuYs7zR5lVvM96XzO+aB85Y8ylElVfM5pyyUi1ItrJT1eqgkp6yUGqkWw1PVqqq9PWT1pGikWt1LuqBXp4akaJkwanA1DupwakaIlopoalzQMWiiigAqC5jjuI2hmjWaOQFWjdQQw7gg1KTUbf6yL6n+RoJM9vDOkO6O2j2bPGGCMbaPKhlCNjjjKgKfUACmyeG9JdpWOj2ZaQEOTbJlwcZB456D8hVS90/xNJdSG11azitzKWVZLUswTHC5z64557/SujoCxz8PhDSINMg0/+yYZ7SGdrqOO5QTYmZ2kaXL5Jcu7MWJzlic81K3h3TmuEnOk2vnozOkvkJuVmfzGIOOCX+Y+p561t0UxciZgSeGNMlZGfRrR2Ri6s1vGSrEsSRxwSXc/8Db1NWLXTY9Ph8m1s0toslvLhRUXJOScCteii5Hsk+pmNDMf+WTfp/jTDbz/APPFvzH+Na1FPmZPsI9zHNvcf88W/Mf40021z/zxb8x/jW1RRzE/V49zD+x3P/PFvzH+NN+x3P8Azxb8x/jW9RRzMX1aPdmAbO5/54N+Y/xpPsd1/wA8G/Mf410FFHMxfVYd2c/9juv+eDfp/jR9juv+eDfmP8a6CijmYfVYd2c6LK7/AOeDfmP8actndf8APBvzH+NdBRRzMr6vHuYQtbn/AJ4N+Y/xp4trj/ni35j/ABraoo5ivYR7mQtvP/zxYfl/jUixTd4m/Ssz+zfFYAC63YEbhndYksRk55DgZxjHHbv1rpqVyvZLuZ4jlH/LNv0p6rJ/zzb9Ku0UXK5EVAsn9xv0pwD/ANxv0qzXK3Gm+LvO22+s2P2dmOXltf3irgYxg7d2c8kYwAMHOQi7HRjd/cb9KdvKjJVgBUtMm/1L/wC6aBiq2adUEbfKPpUitQAE4qnfMSq/WrDNVO8bhfr/AEpmV9TOj1S0meRI7yGRo5PKdVlBKvx8p54PI496im8QabbzQQy6laxS3BxDG86hpOQPlGeeWA49R61zusfCPwnr2qPqF/pK3F1JIZZWaV8SkqFwy7sEYA4xjj3OZ/8AhWHhf+zY9OGkRJYxxmMW6O6oyli3zAH5iGZiCckFjjGTSNTo01C3kICXMTFhkbZAcjrUVrrVhfQxzW1/b3EMg3JJFMrKw9QQea5LT/gp4M0zUbO8t9EhRrOPZbxEkxxklizYJ+ZjvIJbPAGMU8fBfwVuLHw/bu7HczyM7MzZB3Elslsgcnnr6mgDrNO1S01i1W5sLyG9tmJAmt5RIhI4OCDirW4+tZPhvwrpPhGxez0exjsLZ5DK0cecFyACTk9cKPyrVoAXcfWjcfWkooAXcfWjcfWkooAXcfWjcfWkooAXcfWjcfWkooAXcfWjcfWkooAXcfWjcfWkooAXcfWmyTCGNndwiKNzMxwAB1JNLSMoZSCMg8EGgCnHr2nSoXTUbV0G3LLOpA3EqvfuQQPUipv7StipYXUO0dT5gx0J9fQE/QGqq+G9IjjeNdLsljkwHRbdAGw24ZGOcEkj3NMTwvo8dq9uulWf2d87ojApVs5zkY5+8fzNAFv+1bQNt+2Q7s4x5oznGcdfarEcoljV0cOjDKspyCD3FZq+GdHXGNJsRg5GLZOOc+nrzV+3t4rSCOCCJIYY1CJHGoVVUDAAA6ACgCXcfWs5fEmmFrtf7Qt1No4SfdIFEbE4AJPqQR9QR2rQrK/4RXSd94/2CENeSLLcED/WMpJBP0JP50ATSeItLiZVfU7RGbOA1woJwSp7+qsPqp9Kvljg81iL4H8Pr5YXRrJRGcoqwqAvzFuAOnzEn689a2m+6fpQBoxt8o+lSK1Vo2+UfSplamQmIzdaiXDXMIIzyf5GnSHmoo2/0yD6n/0E0zG/vIvsIlKhgoLHCg45OM4H4A/lTvLT+6v5VQ1rw7pviKza11KyivLdmDmOQcEjOD+prBm+Efg+4z5mg2zkqykktlgwAbPPOQB1qTpOoimt55JEjeKR4ztdVIJQ+h9KcvlM7IuwuuNyjGRnpmuc1T4Y+GNauIJr7SYrkwhhGkjMUXLBjhc4HK5+pb1NM/4VX4U2yL/YlviRmduW5YkEnr1yAc+1AG7JqmmxTpC93apM7BVjaRQxJYqABnrlWH1BHararG4yoVhnGRiuaPwy8LtDaRHRrcx2kSwQKd2I0VmZQOeMFm596YPhb4VVpHj0eCKSSQyvImQzMepJz1P/AOrtQB1Plp/dX8qPLT+6v5U6igBvlp/dX8qPLT+6v5U6igBvlp/dX8qPLT+6v5U6igBvlp/dX8qPLT+6v5U6igBvlp/dX8qPLT+6v5U6igBvlp/dX8qPLT+6v5U6igBvlp/dX8qayxxqWYKqqMljgACpKa6LIjI6hlYYKkZBHpQBXW8spFys9uw+U5DqfvHC/mQQPXFTssalQQoLHAzjnvWTH4O0WGGWKPToUjl271UEBtrFlP1DEkHsTTIfBOhwWU1pHp0aW0wIkjUsAwIYEHn0Zh9CR0oA2vLT+6v5UeWn91fyrDHgTw8rBhpFqrBtwYR4IOSc59eT+dbFnZw6faQ2ttGsNvCixxxqOFUDAA+goAk8tP7q/lWSPE+ibr4HULSP7C6x3LSOEWJiSAGJ4GSCPqK2KzB4a0oSXTnTrZmupFlm3xBt7qcqxB7gkn6knvQBFL4o0GBlWTV9NjZs7Q1zGCcEqcc9irD6qfStG5jUW8p2j7h7e1Zv/CGeH9iJ/YWm7YySi/Y48KTnJHHGcn861Lr/AI9Zv9xv5UAZ0T/KPpVhWqjC3yj6VZjaqMExztyajhObyD6n/wBBNOk71Fb/APH/AAfU/wDoJpmX2kSa9d6pZ2qvplhHfzeYAYmlCZXBPU4xyFGecZzg4wcibxF4jivvK/sGz8hggWR9SVSrmMFgRt5w7Kgx1znvipfHvh5fEml21u+mQ6vFHP5j2s0pjVgY5EzkH/b9+/1riLz4a3F9Y3EUvhLSElCxxwSW9y6FUBCkdeSFVfm9edpIwYO07C01/wAVPMsdx4ahTCs7mPUEJ/1bFQB/tOAuegyeoGSW/iDxTPp02/wukGoIyoqNfIY2BVvnyB0BC/KMnDjng1h2/gtpPLtbnwdpqWSWjQq0N0SwDF2aMjjALtk8n73ftDp/gWZdei1S48I6Yt41ytzJcC7Z3WRptzOCe6h5SBjrjB7AEdDN4g8TwNLINAt57Rcus325IxsA9yc+zcZA5VabaeLPEM0tssnhmFY7hHZJY9ViYPhcoFGMnPOccAYOT25nTPBd9ZiS6s/BWj6dc+V5KBrptzpJGRIjFfu4KoO+cnGMZK2vg661O3vra98F6WixvNLFJLOxE0sjpuLckksgBZ8nLAjtlgDqU8V61h9+gwo8Jb7RH/aKZQBQ24EqARzjnHPXHJpF8ReJ2jiZfDEMkpj/AH0CammYJNxwhbbyShjbpxk9eM85q3gg3UwceBtMuJZN0bsbwoBFiEdQOpJkwMfwZ/iwVuPBd8bw6qnhXS21G6gkNyv211PmSeYrru7qVZegGCe+BgA2JfHOtNqVzZW3hyC4kgtzIXbVY0VpABlFyuSA29S+MAr65A1P7a1/bchPD8UjRzBE26ggEiHdlvu8EAJwRzv9qjX4e+H7zTbaC60WBY1Ak+y7iyI+dxIGcE5J+bqRx04qGf4T+ELht0mg2rNlTn5snacgHnoCAcdMigB1xr3imOO3lTwzbtutw8sTakqmObcRsDbcEAAHdjnd2xSXGveKYbZJR4ds2ndgi2f9pgEnc+SHKYOEVW24z8x/umrP/CvPDv2iaYaXCHnVll6/vM45Pvxwfc0xfhr4YVLpBo8G26eOSYHcd5QEJnnoAxAHTBxQBG+t+KtwZfDUBXAyh1Bck5bIB29PuEceucU5de8SG1vN3hmMXkKxtFCuooyzZYhhu2gqQBkZHPI4xUVx8KfC10tnG+lR/Z7WVp0gUkIZCUO9ucsRsXqccfSrUPw78N2+3ytIt49okHygjiTfvB55z5j9f71AEMniTxBDbWsknhuOOSaQxtE+pRjYcptO7HOcvwMn5RxzVb/hJvFUFvbtc+GbaOaS4aIxjU4+V2ZUgkDkt8uBnhSeM4FyP4a+GIrWW2TRrdbeSRZXiAO0uqlVOM9QpIpbX4b+GLHzvI0S0j85dsmE+8MOMH8JH/76+lAEa+Jdehiu3vPDkdsVGbYf2lGRLjBIYkDYQNzcbhhDz0qfSdY1+7u4kvvD62MDuVaVb1JQihWO7gZbJ2gDA6knpiobX4Y+FbNUWDRLaJVxtCggAjIB69cEjPXmpLH4c+GtLvoby00e3trqFzIksWVIYqVPQ9NrEY6YOKAOjVg6hlIZSMgjoaWsLQPA+heF7iSfStNhsZZFKs0eeQSCep/2R+VbtAFKPWtPmVyl9bOE5YrMp28455454qf7ZBtLedHtABJ3DHJwPzqhD4V0e3UiLTLWMMoUhIgMgHcB+B5+tNtvCWj2cDwwafDFG4CsqAjIDBsfmB+QoA0TeW69Z4xxn746etSK6yKGUhlPcHIrHXwXoMe0ppFmhUMF2wgYzyenrWpZ2cGn26W9tEsMEYwsaDAH0oAmrG/4TDRt2oD7fHmwkWK54b92zMVUHjuQRx3BFbNUf7D07zJ3/s+133DiSZvJXMjDozcckepoAzW8feHUEZOr2uJCwUh8glWZWH4MrD6jFbV5/wAek3+438qrNoOmSbd2nWjbTkZgU4/T3qzef8ec/wD1zb+VAGJC3yircbVSh+6v0q1GelWcsSaTqait/wDj+g+p/wDQTU8nU1FCP9Og+p/9BNBP2kX7q+trERm4uIrcSOI081wu5z0UZ6k4PHtUf9rWQnWE3luJm27Y/NXcd2duBnvtbHrg+lM1TRbHWoVivbZLiMP5m1s43AEZOOvBP51lx/D3w5DeQXS6TB9ogkWWKQ5JRl+6Rk8Hk1B2GpHr2mzKGTULV1JUArOpB3KGXv3Ugj1BzQuuac8ywrqFq0zKXEYmXcVBKk4z0BBH1BrHl+GvheY5fQ7Rvk8v7nAXcHwPT5lU/gKnuPAXh+6VFm0uCRUjMKBsnahYsQOeMkkmgDchmjuEDxOsiH+JCCKfVLSdGstDtTbafbR2kBcyGOMYG49TV2gAooooAKKKKACiiigAooooAKKKKACiiigApk0yW8LyyHbGilmb0A5NPooAxoPGGi3MLyxajC8aeXuYHhd7FUz6ZZSPwqdPEukyW7zpqdnJAmd8qTqUTCsxLEHAGFY8+hq1/Z1ptYfZYcMNpHljkdcdKYNIsFjMYsrcIeq+UuD36YoAqL4u0NmCrrWnlidoUXUeSeeOvXg/lWlb3EV3bxzwSpNBIodJI2DK6kZBBHUEd6j/ALPtTwbaHH/XMf4VNHGsaKiKERRgKowAPQUAOrmj8Q9Ejk1VZZ5Yv7NmSCctA5+d2KqFABLZKkcCulpqoqsxCgFjkkDrxjmgDk5fiz4ThEJbV1KzFljZYZGDFWZSMhfVG/DB6EE9Ref8ec//AFzb+VTVDef8ek3+438qYGFD91fpVqPtVeFflH0q1GtUcsSZzVa4HzRn3/pVuRetVJ8hk9M/0pF8uqKy6laNuxdQna5Q4kHDAkFevXIPHsaRtStFRnN1CEUbixkGAACc9emAT+Brnr74WeE9TuLue60S2mlu3L3DNu/ekkMd3PIyAcHjIBqCz+D/AIN0+8N1BoFslwY2iLksfkZGRl5PQqzD6GpNzpbfWtPu7dZ4L62mgYZWSOZWU/Qg0Q6xYXKu0N9bSqjtGxSVSFYHBU4PUHgiuUb4K+DJLy6u5tEjuLm5bfJNPLI7lsJk7i2ckxqxPUtk5yTSx/BXwRFt2eHrZSv3TufI5J4+bg80Abt5408P6fJHHda7pts8jiNFmvI0LMWZAoBPJLI649UYdjWgup2bglbuBgM5xIvbOe/sfyNc2/wn8JSQRwtodu0McSQIhLEBEeR1HXs0shz1+Y1FbfB3wZZ3EdxBoNvHcR7dkwZ967fukHdkEcc+woA7FXWRQysGX1ByKWqOh6Na+HdItNNsY/KtLWMRRr3wO5PcnqT3JNXqACiiigAooooAKKKKACiiigAooooAKRmEalmIVVGSScAClpsiLIjI6hkYYKsMgj0oAiXULWRcrcwsPl5Dg/eOF/MggfSpiwUgEgFjge/esmPwjo0UUkSafCkcm3eqggHaxZfxBJIPbNMi8F6LDZy2kenxpbSgh41LAMCCCDz6Mw+hIoA2qKwx4H0BWDDSbUMDuDCPBB55z68n8617S1hsbWG2t41ighRY4416KoGAB9BQBLWR/wAJdo26/B1CFfsDrHcliQImJIAJPuCPqDWvVD+wdN8yeT+z7YvcOJJWMSnewzhjx1GT+Z9aAK0njLQoWRX1azVpM7VMy5OGZT/48jL9QRWtL/qn+hrP/wCEZ0faq/2TY7VJKj7MmATnOOO+T+daEn+rb6UAEK/KPpVqNajhX5R9KsKtWYJEjLUaqPtUPHc/yNWGWowP9Ii+p/kak1sJfalY6XGJLu4ht4y2zdIwAzgn+QJ+gNVv+Em0XzpIf7UsfNjRXdfPTKqy7lJ56FefoQaq+MtOv9U02KHT7HS76bzcsusRl4VGx8Ngc53bR9Ca5fVfC3ib5Lu30vwpc3MKIyqbRg+4IqlVY9sb0ByPlI/BFHX2/irQ7rYItUspN+du2ZecKWJHPQBWJPbB9KefE2jCxa9/tKz+yqMmXzVwMqWA69cAnHXiuZXwfqcMayR6V4VN01u6P/oTIgkO5BggE7fLKqQfvc/dHFRWPgzU7XVJVj0Xwra6NJd72hS0fzTGrSAPgfJvKsOccFmznPDA6n/hKdEErxtqVokiHDK8gUjjOee2O/SpIfEWj3K5h1OxlGWGUnQ/dGW6HsOtc3pPhHUW1CN9Z03wvPalW837LYMJS5C4ILEjBIOeOy1vL4L8PrcLOuhaYJ1QxrILOPcFPVQdvQ+lICdPEGkSeVt1Kxbzf9XtnQ7/AKc80n/CQaRvRf7Rs90i70/fL8y5IJHPPKt09DTF8J6Kqun9k2ZR2LFGgUrkrtOARgZXg49T6mmHwX4fa2jtzoWmm3jTy0h+xx7EXcW2gbcAbiTj1OaALD67pMbFX1CzUhd5DToMLgHPXpgg/iKT+3tI2u39o2W1H8pj56YVySAp565Vhj2PpVOTwD4bmup7mXQdOnnmj8l3mtkfMeANgyDhflHyjg4q43hvSG83dpdkfOYPJm3T52GcE8cn5m59z60AMk8SaLDNHDJqdikskfnIrToCyE7Qw56E5Ge+D6UN4l0aNgralZrld24yqFxkjrnHVW4/2T6Uw+D9Ba1jtjountbxxeSkLWqFFjyTsAxgLyeOnNK3hHQpLH7G2i6c1nuD/ZzaRmPcGLA7cYzuZmz6sT3oAX/hKNE8xY/7UsSzBSMTL/Fu29++1sfSpP8AhINI8iSb+0rHyY2VXk89NqljhQTngk8D1NMfwrokiFH0ewZSCCrWqEHJJPbuSfzpF8JaGsc8a6NYLHcIqTRi2QLIoOVDDGGAPTPTNAEp17SVUMdRsgpDEHz0wcYz37blz/vD1qGPxVocsMUqarYtFK5jRxOm1mALEA59AT9KVfCOhLDDCNF08RQtvijFrHtjbjlRjg/KvI9B6UxfBnh9U2DQtNC7zJtFnHjeVClunUqAM+gAoAtxaxps4lMd7ayCH/WFZVOz688fjTbPWtM1CTZa3trcvxgRSq2cgkAYPoCfoKrL4N0BFmVdD05VmG2VRaRgSDKnDDHPKr1/uiprXwvo1jcRz22k2NvNE7SRyRWyKyMwIZgQOCQSCe+TQBo7V/uj8qNq/wB0flTqZNEs8TxuMo6lWGccGgBdq/3R+VG1f7o/KsCDwPYW1vLDHPehZBGCTcsSNjl1IJ6YJPsRwakj8JRx2clv/aeqOHz+9e8ZpFyGHyseVOGPI9j1ANAG3tX+6Pyo2r/dH5VhjwfArBhqOrZDbv8AkIykZye27GOenStizthZ2kNurySrEioHmcu7YGMsx5J9SaAJNq/3R+VMjkhkd0Rkdozh1UglTjOD6cVLXNH4faQ0mqsUmP8AaUyT3AMhILKxYYHYZYnHvQB0e1f7o/KmXCDyJOB909vauWl+Fnh+YQh4rphEzMu67kY/MzMcksSeXY9e/pXVTf6mT/dP8qAKca/KPpUqrSRr8o+lSqtMhIkZajZTuVhwVOamqtqFl9utWh86a33FT5kD7HGGB4P4YPqCRSLHmST/AGfy/wDr00zSf7P5H/GsZfCbCNAdZ1VmUY3faBzxzxj/ADinp4ZWKdXGoagVHPlG4O05OT2z+tMk02uJR/c/I/40w3kw/uf98n/Gsv8A4RgqhX+1dSI4OfPAII/4D3702PwyYwf+JpqMmX3ndMOvpwvTpx7fWgnU02vp/SP/AL5P+NMbUrgdo/8Avk/41kf8IniQn+1NTZcghGuOBj8M+n5fXK3HhozvG39o6hH5ecLHPgHJJOeOeuOegAxTJ1NM6pc+kf8A3yf8aadWuv7sX/fJ/wAa5q+8GXtxOHh1+/t4wqLs3Fs7cck5HJxzxzk0knhG/cf8hq4U7twKhuv0347UaE3l3OjbWbofwxf98n/GmnXLv+7D/wB8n/GuUm8E6lMqg+Ir5fkVGKgjdtUDJ+bgnGSRjkn1rpY7VkiRWJkZQAWIxk+tPQi8u5IddvP7sP8A3yf8aT+37z+7D/3wf8ajNufQ00259KehPNPuS/8ACQXn92H/AL5P+NJ/wkF5/ch/74P/AMVUX2c/3ab9nP8AdNGhPNPuT/8ACQXn9yH/AL4P/wAVR/wkF5/ch/74P/xVQfZz/dP5Uv2c/wB2jQOafcn/AOEgvP7kP/fDf40o168/uw/98n/GoPs59KcLc+n6UaD5p9yca5d/3Yf++T/jTxrN2f4Yf++T/jUK259DT1gPp+lGhacu5MurXR/hi/75P+NPXU7k9o/++T/jUSwn0qRYT6Ui05Ei6hcHtH/3yf8AGnrfTn/nn/3yf8aYsJ9KkEdIvUcLqb/Y/wC+T/jT1uJf9n8j/jTVj9qeF9qRWovnSf7P5H/GlZndSpK4Ix0/+vSqtOC0ihqrgAVIq0oWnUDP/9k=";

        private string imagePart1Data = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAAsTAAALEwEAmpwYAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAJWSURBVHgBlVNNaxNRFL13MgW/wGwUoUID/gCzcyPYjYgrXXfVf5DuxJXpQsSVunCdKIouRNOFSrvQ+gMEFyKUoiZgIYaaTGKSmXnv3Xu9702qxY01cHlvMu+ce86Z+xD+43fo1rUKEDTEmSoaWwbrNksHBt+pV5D4LTh3Fq07LM4BOJLoIOBTrxuVSPgFMC9oQShiEUfwT4LTG09rJYm/SG6rCkJfQoFA94SxP7TUHdSUdYXBLbBkHV2bz+bPrM5vPK4K012epgpwIOTErx4Y9qokXvq2uwz+kFhgMMBiKiSmfvnr++HHrU8X2DF4v1pSdCfffVYEsaDU9oE1ZCspZzDm9AZYSpQcVIVoYEE6svjnAA4ECqyygphzJCWZUoYjl8OIsuMiXPZBgbOocr1kLNRQsVfCeGqGMBdHqLJhwrmMbIojUgIlgcxoQmrB2CAZ1bN28mthhSmJusNe24GRn5RCouDEZTB0mVAvQUA9r97Jhyh//MusdJ9EW5+3W+0f37GfTz0BjPIUeHeEpcFUZapvYzHt9dWNeleCohzM5qEZm974Xsdurxw7eQJQpyKe5iDWeKnI1sq4s4NsQgZF8rMVidqRRO9K7uV6El08X86T5BynGYAx6MYTyPsDmex0vf8g3YeHYXxnxVQ3T9bWcW/ijty/2dDLsexfhtD8am0o0Gf9T8AqWW6UyNbNo9aqx/2+TPbVm7W5S4uoZhe9zNDVOlGpqPPgCfUTaupM183D57f3cPj37Psrq5/pina8qt0qYFSuNR/YuM2j+fhB0mwl+8//Ak9GDB97HnnQAAAAAElFTkSuQmCC";

        private string imagePart2Data = "PHN2ZyB3aWR0aD0iMTIiIGhlaWdodD0iMTIiIHZpZXdCb3g9IjAgMCAxMiAxMiIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIiB4bWxuczp4bGluaz0iaHR0cDovL3d3dy53My5vcmcvMTk5OS94bGluayIgZmlsbD0ibm9uZSIgb3ZlcmZsb3c9ImhpZGRlbiI+PHBhdGggZD0iTTYgN0M2LjU1MjI5IDcgNyA2LjU1MjI5IDcgNiA3IDUuNDQ3NzIgNi41NTIyOSA1IDYgNSA1LjQ0NzcyIDUgNSA1LjQ0NzcyIDUgNiA1IDYuNTUyMjkgNS40NDc3MiA3IDYgN1pNMi41IDZDMi41IDQuMDY3IDQuMDY3IDIuNSA2IDIuNSA3LjkzMyAyLjUgOS41IDQuMDY3IDkuNSA2IDkuNSA3LjkzMyA3LjkzMyA5LjUgNiA5LjUgNC4wNjcgOS41IDIuNSA3LjkzMyAyLjUgNlpNNiAzLjVDNC42MTkyOSAzLjUgMy41IDQuNjE5MjkgMy41IDYgMy41IDcuMzgwNzEgNC42MTkyOSA4LjUgNiA4LjUgNy4zODA3MSA4LjUgOC41IDcuMzgwNzEgOC41IDYgOC41IDQuNjE5MjkgNy4zODA3MSAzLjUgNiAzLjVaTTAgNi4wMDA2NkMwIDIuNjg2NTggMi42ODY1OCAwIDYuMDAwNjYgMCA5LjMxNDczIDAgMTIuMDAxMyAyLjY4NjU4IDEyLjAwMTMgNi4wMDA2NiAxMi4wMDEzIDkuMzE0NzMgOS4zMTQ3MyAxMi4wMDEzIDYuMDAwNjYgMTIuMDAxMyAyLjY4NjU4IDEyLjAwMTMgMCA5LjMxNDczIDAgNi4wMDA2NlpNNi4wMDA2NiAxQzMuMjM4ODcgMSAxIDMuMjM4ODcgMSA2LjAwMDY2IDEgOC43NjI0NCAzLjIzODg3IDExLjAwMTMgNi4wMDA2NiAxMS4wMDEzIDguNzYyNDQgMTEuMDAxMyAxMS4wMDEzIDguNzYyNDQgMTEuMDAxMyA2LjAwMDY2IDExLjAwMTMgMy4yMzg4NyA4Ljc2MjQ0IDEgNi4wMDA2NiAxWiIgZmlsbD0iIzQyNDI0MiIvPjwvc3ZnPg==";


        private string imagePart4Data = "iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAOw4AADsOAcy2oYMAAAWWSURBVGhDzZohUB1JEIYRERERiAhEREQEIiLiRMRVccATEREIBOKJJxARCEQE4gRVEQhExImIiIgIBCICERFxIhKBQEQgTiAQiCcQCAT3/Uvvq5nemd19j93j/qoOyfbfPT29PT0zS+a6wNLS0uPV1dU3Kysre8vLy1+Qv/n7L2SMXCFneoZ8FQdZk42ZPwwIYJ5ARgT1zYK8nVKusT3i5ya+nprb/qHMMfAOAyu7qcBmEU1mV0mxYfoBAw2Rf4KBu5YLZIuJPLIhu4FeMRlSXacGLeUczid+jlgPA2xeKqPIE2RRz6SD8xc/m5Jwis1zG/5+UCA4zA14SUB/wnll9NaQX9niQ1lP+mbSfxh9NuDgLY5SC/Qa2SOIe9csPrSmdvGXWlM3yKZRp4MF7x3eMtgPBu3m9QbA5wK+D1JjIu+M1g56vRilMr+PrtsF5mBl5ce9aV1OBPgUA1/zepUjo/QOxlpDfAIviO2ZUfIgA6lu0yp4BlhUBlVm2GgnVhCq7VOefZdOHKPXIlPCx7UVAEF93hvtmzoLOMrYSWDTJOKumXkWTFiL29tumToGM3uMMiodZbJuxuiewzsObaYUZfSFuUuCGA6dzSU21e4HUceDkHitAE1dwWAw+B1OrodPI7X9nhgW4PgWu2fqO2hGjaQAFnyqS92QCB3uhpo8ol1Yojc1NJ0aQsWOSQzMfQXY+c6k5C6YuqjhkSOkXxOwYFKZ/4XutdGyEEdcZyvRmMly4rnK+9zxt01dzFCZmSg1Y1NVgN7XvDKq/aH1+d4C2g98lHJslArQbYdcYvxRKMxZVA48S55t0KnbTHgmjV0qB9k6X5JkdyImvfmQd8OzefVb3aRCxbnZVIDOt0qVzcw3K0ueL6cTU1eA7tRxh3qoK97kIa/mk/EjMNhiyEOUgcaabwI+XslX6FtjmToCug8hr4iVP76ED5Hkrqt1EfL49zdTVUAAz9DrYHaJ6J7wVc9MXYH0zndyDaJbd7wjGUdHh1w7g6fjwYSHDE0VQYGiU+AhVzJGl9xX0G2EXMb6bqoI2KuDhT6PZRzVIKSXxo+Q4CWDYfDckViBHRgtgnw57qmpIsBTckLeuQKLNjBIyf6PzneqJ6aKgC6V/VKSDUK+HG9sqgoc73aawP63EzgLH+As1wHalpA/fE2kgxLSuSjkXfSxiBVM6px0hS55VEDXdhH/FvKQEwUWtTCkbRs9NFUFmoT08Mo2epALXkAftXKNZaoI6KKTALwjPfQb2UfjRyCAPjcyfeWY+NZYpo6Azm9knyuzQs6MXwG6Po4S3mfdUcJzR6WTthl4yMOc3wNUAXcfhVVLoZJ/7xSKBNA/1HF6K+QS409TFcrNUInoE0YyIJ6/QJ/q9Sqnxk+M4sD1pSDRUSN3oXmEPmr3TOC9qQuCvgVFZQQhe6nRHRZO8mqInbraBj79lXIDnbpNNI6JrpRvzH0F6KPsiy+fpr4Dzv0nDGUk+4sH7Rdw6nbdtjKuC54YdF+PrrDEWu2UGWJy5yyBjcqp788q+mQf2uQTi9K/qtpSKgGvlw9bcN4FNoXUxsPMtFj8tU3SOJiAfZefFlWifp2dYV/f7SBowfnavsLhW6P0Dgven6d0lkreVSrAQbLLkMFdo/QGxlHZpDrculHaAQO/NxTCJA7JROe/FsXnPL79gi3lg9GmA4Za1KlsjBlsp7EeWwAfWncaJ/eddbbgS1g55Zzra/Y2QWQ/AueAjc42CjzaYQPRGpiubHKwwZr6vbqXjrvr8F/Lxsxlv4DoMqJ2K05Ty1W3abdg2wKH5avuYvfNiUpTrfbepZkFzrVj6xKUOtPMKjpDfcT3f/p/JlRW2wysTSu10JtEQf9E3uNr6jXUKeyt6BcYnxHdLbRewu/5agIn0onD30fdZHtu7l+7mPH9zYWl+wAAAABJRU5ErkJggg==";

        private string imagePart5Data = "PHN2ZyB2aWV3Qm94PSIwIDAgOTYgOTYiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIgeG1sbnM6eGxpbms9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkveGxpbmsiIGlkPSJJY29uc19MaW5lU2xpZ2h0Q3VydmVfTSIgb3ZlcmZsb3c9ImhpZGRlbiI+PGcgaWQ9Ikljb25zIj48cGF0aCBkPSJNNzUuNzA3IDMyLjI5M0M3NS4zMDk3IDMxLjkwOTMgNzQuNjc2NyAzMS45MjAzIDc0LjI5MyAzMi4zMTc2IDczLjkxODcgMzIuNzA1MSA3My45MTg3IDMzLjMxOTUgNzQuMjkzIDMzLjcwN0w4Ny41NjkgNDYuOTgzQzg3LjU3MjkgNDYuOTg2OSA4Ny41NzI4IDQ2Ljk5MzMgODcuNTY4OSA0Ni45OTcxIDg3LjU2NyA0Ni45OTg5IDg3LjU2NDYgNDcgODcuNTYyIDQ3TDM0IDQ3QzI0LjQ5MTUgNDYuOTk5OSAxNS40NzM2IDQyLjc3ODggOS4zODMgMzUuNDc3TDkuMjgzIDM1LjM1OUM4LjkyODk5IDM0LjkzNDggOC4yOTgxNSAzNC44NzggNy44NzQgMzUuMjMyIDcuNDQ5ODUgMzUuNTg2IDcuMzkyOTkgMzYuMjE2OCA3Ljc0NyAzNi42NDFMNy44OTUgMzYuODE3QzE0LjM2NjEgNDQuNTM5NiAyMy45MjQ2IDQ5LjAwMDUgMzQgNDlMODcuNTYyIDQ5Qzg3LjU2NzUgNDkuMDAwMSA4Ny41NzE5IDQ5LjAwNDYgODcuNTcxOSA0OS4wMTAxIDg3LjU3MTggNDkuMDEyNyA4Ny41NzA4IDQ5LjAxNTIgODcuNTY5IDQ5LjAxN0w3NC4yOTMgNjIuMjkzQzczLjg5NTcgNjIuNjc2NyA3My44ODQ3IDYzLjMwOTcgNzQuMjY4NCA2My43MDcgNzQuNjUyMSA2NC4xMDQzIDc1LjI4NTIgNjQuMTE1MyA3NS42ODI0IDYzLjczMTYgNzUuNjkwOCA2My43MjM1IDc1LjY5OSA2My43MTUzIDc1LjcwNyA2My43MDdMOTAuNzA3IDQ4LjcwN0M5MS4wOTc0IDQ4LjMxNjUgOTEuMDk3NCA0Ny42ODM1IDkwLjcwNyA0Ny4yOTNaIi8+PC9nPjwvc3ZnPg==";

        private string imagePart6Data = "iVBORw0KGgoAAAANSUhEUgAAAYAAAAGACAYAAACkx7W/AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAOw4AADsOAcy2oYMAAAeKSURBVHhe7d0f0KZVGAfghSAIgoVgYWEhCIIgCIIgCIIgWFgIgmAhCBaaCYJgYSFYCIKFIAgWFoIgCBaCIAgWgoUgCIK6fzN9M++8ne993vdrazrnvq6Z30zSt03n2/t+/pxzP5cAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICjPVO5U/n+r3xaea4CwMJS/H+s/LGXXyqvVABY1EeV/eJ/ll8rmgDAoh5WRsX/LJoAwKJ+qIwK/240AYAF5YXvqOjvRxMAWMyVSl74jor+fjQBgMWkqKe4j4r+fjQBgMVoAgCNaQIAjWkCAI1pAovIPI/M+PipkuPe9yrPVwAO0QQmd7mSwj9arJcrAIdoAhPLlf9ooZKfK+4EgC2awKQy1nW0SGdJE7hWAThEE5jQ1pCnJHNAzP0GtmgCk7ldGS3OftIEnq0AHKIJTCRX9scu1jeVfBgC4BBNYCKnLNbXlacqAIdoAhN5o/J7ZbQ4+/miogkAWzSBibxVObYJfFYB2KIJTOTdymhhRskLZIAtmsBEblZGCzPKrQrAFk1gIh9XRgszygcVgC2awESOPSOQ3K14MQxs0QQmksmgo4UZ5UHFOQFgiyYwiVzVZ9vnaGFGyYnhqxWAQzSBSaQJ5BTwaGFGyQA5o6SBLZrAJPJo55QmkMXKuQKAQzSBSZzaBBI7hIAtmsAk8jjo88poYc6LHULAFk1gIh9WRgtzXuwQArZoAhO5Xjl2dlCSHUI+MQkcoglMJP/zH1dGizPKb5X3Kx4JAefRBCaSq/pc3Y8W57zkuwLOCwDn0QQmcrly6g6hLFqmjwKMaAITucgOoeSrypUKwD5NYDKn7hBK8h7BwTFgRBOYzKk7hM6SuUPPVgB2aQKTebHysDJaoEN5VMk3igF2aQKTyXuBjyoXuRvI+wTvBoBdmsCEMh301K2iSc4NfFLxWAg4owlM6OnKncpokbbySyWD5fIzADSBSb1a+akyWqit5FsDNypOEgOawKTySOezymihjkkeJ71ZAXrTBCaW3T6nzBLaT04f544C6EsTmFjGSJzy3eFRvqy8UAF60gQm93blx8powY7N/UpOFHtHAP1oApNL4X6vkpe9o0U7Nvn3P66YOAq9aAILyHbPzBTK9s/Rwp0SdwXQiyawiOwWul3JgbDR4p2S3BXkVLLTxbA+TWAhKdrZNnqRkRL7yc/ICOrsQHJXAOvSBBaTr49lx89oAS+SDJ67W8kjIh+th/VoAgvKbKF8VnK0iBdN7gzyM29VXqoAa9AEFpVDYHmcM1rIf5q8M8hjp2xPNYwO5qYJLCyPhjJo7tgFPjW5O8ipY3cHMC9NYHG5Un+/ctFhc8cmvxzfVtJ03qmkKZhUCv9/mkAD2d2TYXEZMfEktpAek9wl5Ato+TNzCO3dymuVaxXg/0MTaCSzhlKMc8U+WuD/KhlzkZfM9yq5c8hBt9w9pFHlFyynlneT/27g36EJNJSr8RwG+6czh0SkVzSBxWR6aF7qflcZLbiIyG7SBMwXW1BOGucx0YPKkzhtLCJrJgdRWVhOBL9eyYvcvDfQEETkLBlQSSO7DSFnAY59cSQi6yV//2ku7w+uVzKpVFMQ6ZOMl4e/SVPIls6blWzzzC9KPl7vEZLIGsmFXqYOwEnykjlzi9IgblRyBiB3D5k1lLMBOUCW08u7eRIfwxGRJ5MUf9tAARaQK/nHlVGx34/iD7CIFP9jvzuu+AMsQvEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8Qdo6ErlcWVU7Pej+AMs5PPKqNjvR/EHWMyjyqjg70bxB1jQ1uMfxR9gUV9WRoU/UfwBFpYdQCn0ij9AQ2kC9ysp+kn++WoFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFjGpUt/AtrE7R9IdGLVAAAAAElFTkSuQmCC";

        private string extendedPart1Data = "PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiIHN0YW5kYWxvbmU9InllcyI/PjxjbGJsOmxhYmVsTGlzdCB4bWxuczpjbGJsPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS8yMDIwL21pcExhYmVsTWV0YWRhdGEiPjxjbGJsOmxhYmVsIGlkPSJ7ZjQyYWEzNDItODcwNi00Mjg4LWJkMTEtZWJiODU5OTUwMjhjfSIgZW5hYmxlZD0iMSIgbWV0aG9kPSJTdGFuZGFyZCIgc2l0ZUlkPSJ7NzJmOTg4YmYtODZmMS00MWFmLTkxYWItMmQ3Y2QwMTFkYjQ3fSIgY29udGVudEJpdHM9IjAiIHJlbW92ZWQ9IjAiIC8+PC9jbGJsOmxhYmVsTGlzdD4=";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        #endregion

    }
}