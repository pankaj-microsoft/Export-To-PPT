// ---------------------------------------------------------------
// Copyright (c) Microsoft Corporation. All rights reserved.
// ---------------------------------------------------------------

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using P15 = DocumentFormat.OpenXml.Office2013.PowerPoint;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;

namespace CoreGoals
{
    public class CreatePresentationService
    {

        // This is temporary data for the presentation. 
        public List<Goal.GoalDetail> GetGoalsArray()
        {
            List<Goal.GoalDetail> goalsArray = new List<Goal.GoalDetail>();

            for (int i = 1; i <= 23; i++)
            {
                Random rnd = new Random();

                goalsArray.Add(
                    new Goal.GoalDetail(rnd.Next(1000)%2 == 0, "Goal Name - " + i.ToString(), "Goal Owner - " + i.ToString(), (i + 1 * 10) + "K")
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

            NotesMasterPart notesMasterPart1 = presentationPart1.AddNewPart<NotesMasterPart>("rId4");
            GenerateNotesMasterPart1Content(notesMasterPart1);

            SlideMasterPart slideMasterPart1 = presentationPart1.AddNewPart<SlideMasterPart>("rId1");
            GenerateSlideMasterPart1Content(slideMasterPart1);

            SlideLayoutPart slideLayoutPart1 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId12");
            GenerateSlideLayoutPart1Content(slideLayoutPart1);

            slideLayoutPart1.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart2 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId8");
            GenerateSlideLayoutPart2Content(slideLayoutPart2);

            slideLayoutPart2.AddPart(slideMasterPart1, "rId1");

            ThemePart themePart1 = notesMasterPart1.AddNewPart<ThemePart>("rId1");
            GenerateThemePart1Content(themePart1);

            ThemePart themePart2 = slideMasterPart1.AddNewPart<ThemePart>("rId13");
            GenerateThemePart2Content(themePart2);

            presentationPart1.AddPart(themePart2, "rId7");

            SlideLayoutPart slideLayoutPart3 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId3");
            GenerateSlideLayoutPart3Content(slideLayoutPart3);

            slideLayoutPart3.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart4 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId7");
            GenerateSlideLayoutPart4Content(slideLayoutPart4);

            slideLayoutPart4.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart5 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId2");
            GenerateSlideLayoutPart5Content(slideLayoutPart5);

            slideLayoutPart5.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart6 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId1");
            GenerateSlideLayoutPart6Content(slideLayoutPart6);

            slideLayoutPart6.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart7 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId6");
            GenerateSlideLayoutPart7Content(slideLayoutPart7);

            slideLayoutPart7.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart8 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId11");
            GenerateSlideLayoutPart8Content(slideLayoutPart8);

            slideLayoutPart8.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart9 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId5");
            GenerateSlideLayoutPart9Content(slideLayoutPart9);

            slideLayoutPart9.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart10 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId10");
            GenerateSlideLayoutPart10Content(slideLayoutPart10);

            slideLayoutPart10.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart11 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId4");
            GenerateSlideLayoutPart11Content(slideLayoutPart11);

            slideLayoutPart11.AddPart(slideMasterPart1, "rId1");

            SlideLayoutPart slideLayoutPart12 = slideMasterPart1.AddNewPart<SlideLayoutPart>("rId9");
            GenerateSlideLayoutPart12Content(slideLayoutPart12);

            slideLayoutPart12.AddPart(slideMasterPart1, "rId1");

            TableStylesPart tableStylesPart1 = presentationPart1.AddNewPart<TableStylesPart>("rId8");
            GenerateTableStylesPart1Content(tableStylesPart1);

            // ****************************** Presentation level settings End *************************************


            // ****************************** Slide Creation Start *************************************

            for (int i = 0; i < totalSlides; i++)
            {
                List<Goal.GoalDetail> goalsArrayForSlide = goalsArray.Skip(i * 6).Take(6).ToList();

                SlidePart slidePart1 = presentationPart1.AddNewPart<SlidePart>("rId" + (i + 20).ToString());
                GenerateSlidePartContent(slidePart1, goalsArrayForSlide);

                NotesSlidePart notesSlidePart1 = slidePart1.AddNewPart<NotesSlidePart>("rId2");
                GenerateNotesSlidePartContent(notesSlidePart1);

                notesSlidePart1.AddPart(slidePart1, "rId2");

                notesSlidePart1.AddPart(notesMasterPart1, "rId1");

                slidePart1.AddPart(slideLayoutPart1, "rId1");

                ImagePart imagePart1 = slidePart1.AddNewPart<ImagePart>("image/png", "rId8");
                GenerateImagePart1Content(imagePart1);

                ImagePart imagePart2 = slidePart1.AddNewPart<ImagePart>("image/png", "rId3");
                GenerateImagePart2Content(imagePart2);

                ImagePart imagePart3 = slidePart1.AddNewPart<ImagePart>("image/png", "rId7");
                GenerateImagePart3Content(imagePart3);

                ImagePart imagePart4 = slidePart1.AddNewPart<ImagePart>("image/svg+xml", "rId12");
                GenerateImagePart4Content(imagePart4);

                ImagePart imagePart5 = slidePart1.AddNewPart<ImagePart>("image/svg+xml", "rId6");
                GenerateImagePart5Content(imagePart5);

                ImagePart imagePart6 = slidePart1.AddNewPart<ImagePart>("image/png", "rId11");
                GenerateImagePart6Content(imagePart6);

                ImagePart imagePart7 = slidePart1.AddNewPart<ImagePart>("image/png", "rId5");
                GenerateImagePart7Content(imagePart7);

                ImagePart imagePart8 = slidePart1.AddNewPart<ImagePart>("image/png", "rId10");
                GenerateImagePart8Content(imagePart8);

                ImagePart imagePart9 = slidePart1.AddNewPart<ImagePart>("image/png", "rId4");
                GenerateImagePart9Content(imagePart9);

                ImagePart imagePart10 = slidePart1.AddNewPart<ImagePart>("image/svg+xml", "rId9");
                GenerateImagePart10Content(imagePart10);
            }

            // ****************************** Slide Creation End *************************************

            ViewPropertiesPart viewPropertiesPart1 = presentationPart1.AddNewPart<ViewPropertiesPart>("rId6");
            GenerateViewPropertiesPart1Content(viewPropertiesPart1);

            PresentationPropertiesPart presentationPropertiesPart1 = presentationPart1.AddNewPart<PresentationPropertiesPart>("rId5");
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
            Presentation presentation1 = new Presentation() { SaveSubsetFonts = true };
            presentation1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            presentation1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            presentation1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList();
            SlideMasterId slideMasterId1 = new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" };

            slideMasterIdList1.Append(slideMasterId1);

            NotesMasterIdList notesMasterIdList1 = new NotesMasterIdList();
            NotesMasterId notesMasterId1 = new NotesMasterId() { Id = "rId4" };

            notesMasterIdList1.Append(notesMasterId1);

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
            presentation1.Append(notesMasterIdList1);
            presentation1.Append(slideIdList1);
            presentation1.Append(slideSize1);
            presentation1.Append(notesSize1);
            presentation1.Append(defaultTextStyle1);
            presentation1.Append(presentationExtensionList1);

            presentationPart1.Presentation = presentation1;
        }

        // Generates content of notesMasterPart1.
        private void GenerateNotesMasterPart1Content(NotesMasterPart notesMasterPart1)
        {
            NotesMaster notesMaster1 = new NotesMaster();
            notesMaster1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            notesMaster1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            notesMaster1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData1 = new CommonSlideData();

            Background background1 = new Background();

            BackgroundStyleReference backgroundStyleReference1 = new BackgroundStyleReference() { Index = (UInt32Value)1001U };
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            backgroundStyleReference1.Append(schemeColor10);

            background1.Append(backgroundStyleReference1);

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
            NonVisualDrawingProperties nonVisualDrawingProperties2 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Header Placeholder 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties1 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties1.Append(shapeLocks1);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape1 = new PlaceholderShape() { Type = PlaceholderValues.Header, Size = PlaceholderSizeValues.Quarter };

            applicationNonVisualDrawingProperties2.Append(placeholderShape1);

            nonVisualShapeProperties1.Append(nonVisualDrawingProperties2);
            nonVisualShapeProperties1.Append(nonVisualShapeDrawingProperties1);
            nonVisualShapeProperties1.Append(applicationNonVisualDrawingProperties2);

            ShapeProperties shapeProperties1 = new ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset2 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents2 = new A.Extents() { Cx = 2971800L, Cy = 458788L };

            transform2D1.Append(offset2);
            transform2D1.Append(extents2);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            TextBody textBody1 = new TextBody();
            A.BodyProperties bodyProperties1 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false };

            A.ListStyle listStyle1 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties2 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Left };
            A.DefaultRunProperties defaultRunProperties11 = new A.DefaultRunProperties() { FontSize = 1200 };

            level1ParagraphProperties2.Append(defaultRunProperties11);

            listStyle1.Append(level1ParagraphProperties2);

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
            NonVisualDrawingProperties nonVisualDrawingProperties3 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Date Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties2 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks2 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties2.Append(shapeLocks2);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties3 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape2 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties3.Append(placeholderShape2);

            nonVisualShapeProperties2.Append(nonVisualDrawingProperties3);
            nonVisualShapeProperties2.Append(nonVisualShapeDrawingProperties2);
            nonVisualShapeProperties2.Append(applicationNonVisualDrawingProperties3);

            ShapeProperties shapeProperties2 = new ShapeProperties();

            A.Transform2D transform2D2 = new A.Transform2D();
            A.Offset offset3 = new A.Offset() { X = 3884613L, Y = 0L };
            A.Extents extents3 = new A.Extents() { Cx = 2971800L, Cy = 458788L };

            transform2D2.Append(offset3);
            transform2D2.Append(extents3);

            A.PresetGeometry presetGeometry2 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList2 = new A.AdjustValueList();

            presetGeometry2.Append(adjustValueList2);

            shapeProperties2.Append(transform2D2);
            shapeProperties2.Append(presetGeometry2);

            TextBody textBody2 = new TextBody();
            A.BodyProperties bodyProperties2 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false };

            A.ListStyle listStyle2 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties3 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Right };
            A.DefaultRunProperties defaultRunProperties12 = new A.DefaultRunProperties() { FontSize = 1200 };

            level1ParagraphProperties3.Append(defaultRunProperties12);

            listStyle2.Append(level1ParagraphProperties3);

            A.Paragraph paragraph2 = new A.Paragraph();

            A.Field field1 = new A.Field() { Id = "{E6E44BD1-00CF-4A58-8AC8-7E34A8FF1681}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties1 = new A.RunProperties() { Language = "en-US" };
            runProperties1.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text1 = new A.Text();
            text1.Text = "4/24/2024";

            field1.Append(runProperties1);
            field1.Append(text1);
            A.EndParagraphRunProperties endParagraphRunProperties2 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph2.Append(field1);
            paragraph2.Append(endParagraphRunProperties2);

            textBody2.Append(bodyProperties2);
            textBody2.Append(listStyle2);
            textBody2.Append(paragraph2);

            shape2.Append(nonVisualShapeProperties2);
            shape2.Append(shapeProperties2);
            shape2.Append(textBody2);

            Shape shape3 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties3 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties4 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Slide Image Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties3 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks3 = new A.ShapeLocks() { NoGrouping = true, NoRotation = true, NoChangeAspect = true };

            nonVisualShapeDrawingProperties3.Append(shapeLocks3);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties4 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape3 = new PlaceholderShape() { Type = PlaceholderValues.SlideImage, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties4.Append(placeholderShape3);

            nonVisualShapeProperties3.Append(nonVisualDrawingProperties4);
            nonVisualShapeProperties3.Append(nonVisualShapeDrawingProperties3);
            nonVisualShapeProperties3.Append(applicationNonVisualDrawingProperties4);

            ShapeProperties shapeProperties3 = new ShapeProperties();

            A.Transform2D transform2D3 = new A.Transform2D();
            A.Offset offset4 = new A.Offset() { X = 685800L, Y = 1143000L };
            A.Extents extents4 = new A.Extents() { Cx = 5486400L, Cy = 3086100L };

            transform2D3.Append(offset4);
            transform2D3.Append(extents4);

            A.PresetGeometry presetGeometry3 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList3 = new A.AdjustValueList();

            presetGeometry3.Append(adjustValueList3);
            A.NoFill noFill1 = new A.NoFill();

            A.Outline outline1 = new A.Outline() { Width = 12700 };

            A.SolidFill solidFill10 = new A.SolidFill();
            A.PresetColor presetColor1 = new A.PresetColor() { Val = A.PresetColorValues.Black };

            solidFill10.Append(presetColor1);

            outline1.Append(solidFill10);

            shapeProperties3.Append(transform2D3);
            shapeProperties3.Append(presetGeometry3);
            shapeProperties3.Append(noFill1);
            shapeProperties3.Append(outline1);

            TextBody textBody3 = new TextBody();
            A.BodyProperties bodyProperties3 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle3 = new A.ListStyle();

            A.Paragraph paragraph3 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties3 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph3.Append(endParagraphRunProperties3);

            textBody3.Append(bodyProperties3);
            textBody3.Append(listStyle3);
            textBody3.Append(paragraph3);

            shape3.Append(nonVisualShapeProperties3);
            shape3.Append(shapeProperties3);
            shape3.Append(textBody3);

            Shape shape4 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties4 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties5 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Notes Placeholder 4" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties4 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks4 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties4.Append(shapeLocks4);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties5 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape4 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)3U };

            applicationNonVisualDrawingProperties5.Append(placeholderShape4);

            nonVisualShapeProperties4.Append(nonVisualDrawingProperties5);
            nonVisualShapeProperties4.Append(nonVisualShapeDrawingProperties4);
            nonVisualShapeProperties4.Append(applicationNonVisualDrawingProperties5);

            ShapeProperties shapeProperties4 = new ShapeProperties();

            A.Transform2D transform2D4 = new A.Transform2D();
            A.Offset offset5 = new A.Offset() { X = 685800L, Y = 4400550L };
            A.Extents extents5 = new A.Extents() { Cx = 5486400L, Cy = 3600450L };

            transform2D4.Append(offset5);
            transform2D4.Append(extents5);

            A.PresetGeometry presetGeometry4 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList4 = new A.AdjustValueList();

            presetGeometry4.Append(adjustValueList4);

            shapeProperties4.Append(transform2D4);
            shapeProperties4.Append(presetGeometry4);

            TextBody textBody4 = new TextBody();
            A.BodyProperties bodyProperties4 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false };
            A.ListStyle listStyle4 = new A.ListStyle();

            A.Paragraph paragraph4 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties1 = new A.ParagraphProperties() { Level = 0 };

            A.Run run1 = new A.Run();
            A.RunProperties runProperties2 = new A.RunProperties() { Language = "en-US" };
            A.Text text2 = new A.Text();
            text2.Text = "Click to edit Master text styles";

            run1.Append(runProperties2);
            run1.Append(text2);

            paragraph4.Append(paragraphProperties1);
            paragraph4.Append(run1);

            A.Paragraph paragraph5 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties2 = new A.ParagraphProperties() { Level = 1 };

            A.Run run2 = new A.Run();
            A.RunProperties runProperties3 = new A.RunProperties() { Language = "en-US" };
            A.Text text3 = new A.Text();
            text3.Text = "Second level";

            run2.Append(runProperties3);
            run2.Append(text3);

            paragraph5.Append(paragraphProperties2);
            paragraph5.Append(run2);

            A.Paragraph paragraph6 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties3 = new A.ParagraphProperties() { Level = 2 };

            A.Run run3 = new A.Run();
            A.RunProperties runProperties4 = new A.RunProperties() { Language = "en-US" };
            A.Text text4 = new A.Text();
            text4.Text = "Third level";

            run3.Append(runProperties4);
            run3.Append(text4);

            paragraph6.Append(paragraphProperties3);
            paragraph6.Append(run3);

            A.Paragraph paragraph7 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties4 = new A.ParagraphProperties() { Level = 3 };

            A.Run run4 = new A.Run();
            A.RunProperties runProperties5 = new A.RunProperties() { Language = "en-US" };
            A.Text text5 = new A.Text();
            text5.Text = "Fourth level";

            run4.Append(runProperties5);
            run4.Append(text5);

            paragraph7.Append(paragraphProperties4);
            paragraph7.Append(run4);

            A.Paragraph paragraph8 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties5 = new A.ParagraphProperties() { Level = 4 };

            A.Run run5 = new A.Run();
            A.RunProperties runProperties6 = new A.RunProperties() { Language = "en-US" };
            A.Text text6 = new A.Text();
            text6.Text = "Fifth level";

            run5.Append(runProperties6);
            run5.Append(text6);

            paragraph8.Append(paragraphProperties5);
            paragraph8.Append(run5);

            textBody4.Append(bodyProperties4);
            textBody4.Append(listStyle4);
            textBody4.Append(paragraph4);
            textBody4.Append(paragraph5);
            textBody4.Append(paragraph6);
            textBody4.Append(paragraph7);
            textBody4.Append(paragraph8);

            shape4.Append(nonVisualShapeProperties4);
            shape4.Append(shapeProperties4);
            shape4.Append(textBody4);

            Shape shape5 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties5 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties6 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Footer Placeholder 5" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties5 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks5 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties5.Append(shapeLocks5);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties6 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape5 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)4U };

            applicationNonVisualDrawingProperties6.Append(placeholderShape5);

            nonVisualShapeProperties5.Append(nonVisualDrawingProperties6);
            nonVisualShapeProperties5.Append(nonVisualShapeDrawingProperties5);
            nonVisualShapeProperties5.Append(applicationNonVisualDrawingProperties6);

            ShapeProperties shapeProperties5 = new ShapeProperties();

            A.Transform2D transform2D5 = new A.Transform2D();
            A.Offset offset6 = new A.Offset() { X = 0L, Y = 8685213L };
            A.Extents extents6 = new A.Extents() { Cx = 2971800L, Cy = 458787L };

            transform2D5.Append(offset6);
            transform2D5.Append(extents6);

            A.PresetGeometry presetGeometry5 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList5 = new A.AdjustValueList();

            presetGeometry5.Append(adjustValueList5);

            shapeProperties5.Append(transform2D5);
            shapeProperties5.Append(presetGeometry5);

            TextBody textBody5 = new TextBody();
            A.BodyProperties bodyProperties5 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle5 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties4 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Left };
            A.DefaultRunProperties defaultRunProperties13 = new A.DefaultRunProperties() { FontSize = 1200 };

            level1ParagraphProperties4.Append(defaultRunProperties13);

            listStyle5.Append(level1ParagraphProperties4);

            A.Paragraph paragraph9 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties4 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph9.Append(endParagraphRunProperties4);

            textBody5.Append(bodyProperties5);
            textBody5.Append(listStyle5);
            textBody5.Append(paragraph9);

            shape5.Append(nonVisualShapeProperties5);
            shape5.Append(shapeProperties5);
            shape5.Append(textBody5);

            Shape shape6 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties6 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties7 = new NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Slide Number Placeholder 6" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties6 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks6 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties6.Append(shapeLocks6);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties7 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape6 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)5U };

            applicationNonVisualDrawingProperties7.Append(placeholderShape6);

            nonVisualShapeProperties6.Append(nonVisualDrawingProperties7);
            nonVisualShapeProperties6.Append(nonVisualShapeDrawingProperties6);
            nonVisualShapeProperties6.Append(applicationNonVisualDrawingProperties7);

            ShapeProperties shapeProperties6 = new ShapeProperties();

            A.Transform2D transform2D6 = new A.Transform2D();
            A.Offset offset7 = new A.Offset() { X = 3884613L, Y = 8685213L };
            A.Extents extents7 = new A.Extents() { Cx = 2971800L, Cy = 458787L };

            transform2D6.Append(offset7);
            transform2D6.Append(extents7);

            A.PresetGeometry presetGeometry6 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList6 = new A.AdjustValueList();

            presetGeometry6.Append(adjustValueList6);

            shapeProperties6.Append(transform2D6);
            shapeProperties6.Append(presetGeometry6);

            TextBody textBody6 = new TextBody();
            A.BodyProperties bodyProperties6 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle6 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties5 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Right };
            A.DefaultRunProperties defaultRunProperties14 = new A.DefaultRunProperties() { FontSize = 1200 };

            level1ParagraphProperties5.Append(defaultRunProperties14);

            listStyle6.Append(level1ParagraphProperties5);

            A.Paragraph paragraph10 = new A.Paragraph();

            A.Field field2 = new A.Field() { Id = "{C5EE4AD6-E395-487C-BCB9-5F2C75A2E299}", Type = "slidenum" };

            A.RunProperties runProperties7 = new A.RunProperties() { Language = "en-US" };
            runProperties7.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text7 = new A.Text();
            text7.Text = "‹#›";

            field2.Append(runProperties7);
            field2.Append(text7);
            A.EndParagraphRunProperties endParagraphRunProperties5 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph10.Append(field2);
            paragraph10.Append(endParagraphRunProperties5);

            textBody6.Append(bodyProperties6);
            textBody6.Append(listStyle6);
            textBody6.Append(paragraph10);

            shape6.Append(nonVisualShapeProperties6);
            shape6.Append(shapeProperties6);
            shape6.Append(textBody6);

            shapeTree1.Append(nonVisualGroupShapeProperties1);
            shapeTree1.Append(groupShapeProperties1);
            shapeTree1.Append(shape1);
            shapeTree1.Append(shape2);
            shapeTree1.Append(shape3);
            shapeTree1.Append(shape4);
            shapeTree1.Append(shape5);
            shapeTree1.Append(shape6);

            CommonSlideDataExtensionList commonSlideDataExtensionList1 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension1 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId1 = new P14.CreationId() { Val = (UInt32Value)2031944351U };
            creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension1.Append(creationId1);

            commonSlideDataExtensionList1.Append(commonSlideDataExtension1);

            commonSlideData1.Append(background1);
            commonSlideData1.Append(shapeTree1);
            commonSlideData1.Append(commonSlideDataExtensionList1);
            ColorMap colorMap1 = new ColorMap() { Background1 = A.ColorSchemeIndexValues.Light1, Text1 = A.ColorSchemeIndexValues.Dark1, Background2 = A.ColorSchemeIndexValues.Light2, Text2 = A.ColorSchemeIndexValues.Dark2, Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2, Accent3 = A.ColorSchemeIndexValues.Accent3, Accent4 = A.ColorSchemeIndexValues.Accent4, Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6, Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink };

            NotesStyle notesStyle1 = new NotesStyle();

            A.Level1ParagraphProperties level1ParagraphProperties6 = new A.Level1ParagraphProperties() { LeftMargin = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties15 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill11 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill11.Append(schemeColor11);
            A.LatinFont latinFont10 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont10 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont10 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties15.Append(solidFill11);
            defaultRunProperties15.Append(latinFont10);
            defaultRunProperties15.Append(eastAsianFont10);
            defaultRunProperties15.Append(complexScriptFont10);

            level1ParagraphProperties6.Append(defaultRunProperties15);

            A.Level2ParagraphProperties level2ParagraphProperties2 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties16 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill12 = new A.SolidFill();
            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill12.Append(schemeColor12);
            A.LatinFont latinFont11 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont11 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont11 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties16.Append(solidFill12);
            defaultRunProperties16.Append(latinFont11);
            defaultRunProperties16.Append(eastAsianFont11);
            defaultRunProperties16.Append(complexScriptFont11);

            level2ParagraphProperties2.Append(defaultRunProperties16);

            A.Level3ParagraphProperties level3ParagraphProperties2 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties17 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill13 = new A.SolidFill();
            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill13.Append(schemeColor13);
            A.LatinFont latinFont12 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont12 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont12 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties17.Append(solidFill13);
            defaultRunProperties17.Append(latinFont12);
            defaultRunProperties17.Append(eastAsianFont12);
            defaultRunProperties17.Append(complexScriptFont12);

            level3ParagraphProperties2.Append(defaultRunProperties17);

            A.Level4ParagraphProperties level4ParagraphProperties2 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties18 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill14 = new A.SolidFill();
            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill14.Append(schemeColor14);
            A.LatinFont latinFont13 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont13 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont13 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties18.Append(solidFill14);
            defaultRunProperties18.Append(latinFont13);
            defaultRunProperties18.Append(eastAsianFont13);
            defaultRunProperties18.Append(complexScriptFont13);

            level4ParagraphProperties2.Append(defaultRunProperties18);

            A.Level5ParagraphProperties level5ParagraphProperties2 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties19 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill15 = new A.SolidFill();
            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill15.Append(schemeColor15);
            A.LatinFont latinFont14 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont14 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont14 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties19.Append(solidFill15);
            defaultRunProperties19.Append(latinFont14);
            defaultRunProperties19.Append(eastAsianFont14);
            defaultRunProperties19.Append(complexScriptFont14);

            level5ParagraphProperties2.Append(defaultRunProperties19);

            A.Level6ParagraphProperties level6ParagraphProperties2 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties20 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill16 = new A.SolidFill();
            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill16.Append(schemeColor16);
            A.LatinFont latinFont15 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont15 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont15 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties20.Append(solidFill16);
            defaultRunProperties20.Append(latinFont15);
            defaultRunProperties20.Append(eastAsianFont15);
            defaultRunProperties20.Append(complexScriptFont15);

            level6ParagraphProperties2.Append(defaultRunProperties20);

            A.Level7ParagraphProperties level7ParagraphProperties2 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties21 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill17 = new A.SolidFill();
            A.SchemeColor schemeColor17 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill17.Append(schemeColor17);
            A.LatinFont latinFont16 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont16 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont16 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties21.Append(solidFill17);
            defaultRunProperties21.Append(latinFont16);
            defaultRunProperties21.Append(eastAsianFont16);
            defaultRunProperties21.Append(complexScriptFont16);

            level7ParagraphProperties2.Append(defaultRunProperties21);

            A.Level8ParagraphProperties level8ParagraphProperties2 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties22 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill18 = new A.SolidFill();
            A.SchemeColor schemeColor18 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill18.Append(schemeColor18);
            A.LatinFont latinFont17 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont17 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont17 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties22.Append(solidFill18);
            defaultRunProperties22.Append(latinFont17);
            defaultRunProperties22.Append(eastAsianFont17);
            defaultRunProperties22.Append(complexScriptFont17);

            level8ParagraphProperties2.Append(defaultRunProperties22);

            A.Level9ParagraphProperties level9ParagraphProperties2 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties23 = new A.DefaultRunProperties() { FontSize = 1200, Kerning = 1200 };

            A.SolidFill solidFill19 = new A.SolidFill();
            A.SchemeColor schemeColor19 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill19.Append(schemeColor19);
            A.LatinFont latinFont18 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont18 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont18 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties23.Append(solidFill19);
            defaultRunProperties23.Append(latinFont18);
            defaultRunProperties23.Append(eastAsianFont18);
            defaultRunProperties23.Append(complexScriptFont18);

            level9ParagraphProperties2.Append(defaultRunProperties23);

            notesStyle1.Append(level1ParagraphProperties6);
            notesStyle1.Append(level2ParagraphProperties2);
            notesStyle1.Append(level3ParagraphProperties2);
            notesStyle1.Append(level4ParagraphProperties2);
            notesStyle1.Append(level5ParagraphProperties2);
            notesStyle1.Append(level6ParagraphProperties2);
            notesStyle1.Append(level7ParagraphProperties2);
            notesStyle1.Append(level8ParagraphProperties2);
            notesStyle1.Append(level9ParagraphProperties2);

            notesMaster1.Append(commonSlideData1);
            notesMaster1.Append(colorMap1);
            notesMaster1.Append(notesStyle1);

            notesMasterPart1.NotesMaster = notesMaster1;
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
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "0E2841" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "E8E8E8" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "156082" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "E97132" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "196B24" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "0F9ED5" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "A02B93" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "4EA72E" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "467886" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "96607D" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

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
            A.LatinFont latinFont19 = new A.LatinFont() { Typeface = "Aptos Display", Panose = "02110004020202020204" };
            A.EastAsianFont eastAsianFont19 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont19 = new A.ComplexScriptFont() { Typeface = "" };
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

            majorFont1.Append(latinFont19);
            majorFont1.Append(eastAsianFont19);
            majorFont1.Append(complexScriptFont19);
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
            A.LatinFont latinFont20 = new A.LatinFont() { Typeface = "Aptos", Panose = "02110004020202020204" };
            A.EastAsianFont eastAsianFont20 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont20 = new A.ComplexScriptFont() { Typeface = "" };
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

            minorFont1.Append(latinFont20);
            minorFont1.Append(eastAsianFont20);
            minorFont1.Append(complexScriptFont20);
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

            A.SolidFill solidFill20 = new A.SolidFill();
            A.SchemeColor schemeColor20 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill20.Append(schemeColor20);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor21 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 105000 };
            A.Tint tint1 = new A.Tint() { Val = 67000 };

            schemeColor21.Append(luminanceModulation1);
            schemeColor21.Append(saturationModulation1);
            schemeColor21.Append(tint1);

            gradientStop1.Append(schemeColor21);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor22 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 103000 };
            A.Tint tint2 = new A.Tint() { Val = 73000 };

            schemeColor22.Append(luminanceModulation2);
            schemeColor22.Append(saturationModulation2);
            schemeColor22.Append(tint2);

            gradientStop2.Append(schemeColor22);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor23 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 109000 };
            A.Tint tint3 = new A.Tint() { Val = 81000 };

            schemeColor23.Append(luminanceModulation3);
            schemeColor23.Append(saturationModulation3);
            schemeColor23.Append(tint3);

            gradientStop3.Append(schemeColor23);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor24 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 103000 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 102000 };
            A.Tint tint4 = new A.Tint() { Val = 94000 };

            schemeColor24.Append(saturationModulation4);
            schemeColor24.Append(luminanceModulation4);
            schemeColor24.Append(tint4);

            gradientStop4.Append(schemeColor24);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor25 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 110000 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 100000 };
            A.Shade shade1 = new A.Shade() { Val = 100000 };

            schemeColor25.Append(saturationModulation5);
            schemeColor25.Append(luminanceModulation5);
            schemeColor25.Append(shade1);

            gradientStop5.Append(schemeColor25);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor26 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 120000 };
            A.Shade shade2 = new A.Shade() { Val = 78000 };

            schemeColor26.Append(luminanceModulation6);
            schemeColor26.Append(saturationModulation6);
            schemeColor26.Append(shade2);

            gradientStop6.Append(schemeColor26);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill20);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline2 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill21 = new A.SolidFill();
            A.SchemeColor schemeColor27 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill21.Append(schemeColor27);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter() { Limit = 800000 };

            outline2.Append(solidFill21);
            outline2.Append(presetDash1);
            outline2.Append(miter1);

            A.Outline outline3 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill22 = new A.SolidFill();
            A.SchemeColor schemeColor28 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill22.Append(schemeColor28);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter() { Limit = 800000 };

            outline3.Append(solidFill22);
            outline3.Append(presetDash2);
            outline3.Append(miter2);

            A.Outline outline4 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill23 = new A.SolidFill();
            A.SchemeColor schemeColor29 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill23.Append(schemeColor29);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter() { Limit = 800000 };

            outline4.Append(solidFill23);
            outline4.Append(presetDash3);
            outline4.Append(miter3);

            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);
            lineStyleList1.Append(outline4);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();
            A.EffectList effectList1 = new A.EffectList();

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();
            A.EffectList effectList2 = new A.EffectList();

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 63000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList3.Append(outerShadow1);

            effectStyle3.Append(effectList3);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill24 = new A.SolidFill();
            A.SchemeColor schemeColor30 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill24.Append(schemeColor30);

            A.SolidFill solidFill25 = new A.SolidFill();

            A.SchemeColor schemeColor31 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

            schemeColor31.Append(tint5);
            schemeColor31.Append(saturationModulation7);

            solidFill25.Append(schemeColor31);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor32 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 150000 };
            A.Shade shade3 = new A.Shade() { Val = 98000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 102000 };

            schemeColor32.Append(tint6);
            schemeColor32.Append(saturationModulation8);
            schemeColor32.Append(shade3);
            schemeColor32.Append(luminanceModulation7);

            gradientStop7.Append(schemeColor32);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor33 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint7 = new A.Tint() { Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 130000 };
            A.Shade shade4 = new A.Shade() { Val = 90000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 103000 };

            schemeColor33.Append(tint7);
            schemeColor33.Append(saturationModulation9);
            schemeColor33.Append(shade4);
            schemeColor33.Append(luminanceModulation8);

            gradientStop8.Append(schemeColor33);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor34 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade5 = new A.Shade() { Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 120000 };

            schemeColor34.Append(shade5);
            schemeColor34.Append(saturationModulation10);

            gradientStop9.Append(schemeColor34);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList1.Append(solidFill24);
            backgroundFillStyleList1.Append(solidFill25);
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
            A.ShapeProperties shapeProperties7 = new A.ShapeProperties();
            A.BodyProperties bodyProperties7 = new A.BodyProperties();
            A.ListStyle listStyle7 = new A.ListStyle();

            A.ShapeStyle shapeStyle1 = new A.ShapeStyle();

            A.LineReference lineReference1 = new A.LineReference() { Index = (UInt32Value)2U };
            A.SchemeColor schemeColor35 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            lineReference1.Append(schemeColor35);

            A.FillReference fillReference1 = new A.FillReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor36 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference1.Append(schemeColor36);

            A.EffectReference effectReference1 = new A.EffectReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor37 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference1.Append(schemeColor37);

            A.FontReference fontReference1 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor38 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference1.Append(schemeColor38);

            shapeStyle1.Append(lineReference1);
            shapeStyle1.Append(fillReference1);
            shapeStyle1.Append(effectReference1);
            shapeStyle1.Append(fontReference1);

            lineDefault1.Append(shapeProperties7);
            lineDefault1.Append(bodyProperties7);
            lineDefault1.Append(listStyle7);
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

        // Generates content of tableStylesPart1.
        private void GenerateTableStylesPart1Content(TableStylesPart tableStylesPart1)
        {
            A.TableStyleList tableStyleList1 = new A.TableStyleList() { Default = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" };
            tableStyleList1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            tableStylesPart1.TableStyleList = tableStyleList1;
        }

        // Generates content of slidePart1.
        private void GenerateSlidePartContent(SlidePart slidePart1, List<Goal.GoalDetail> goalsArrayForSlide)
        {
            Slide slide1 = new Slide();
            slide1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slide1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slide1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData2 = new CommonSlideData();

            ShapeTree shapeTree2 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties2 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties8 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties2 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties8 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties2.Append(nonVisualDrawingProperties8);
            nonVisualGroupShapeProperties2.Append(nonVisualGroupShapeDrawingProperties2);
            nonVisualGroupShapeProperties2.Append(applicationNonVisualDrawingProperties8);

            GroupShapeProperties groupShapeProperties2 = new GroupShapeProperties();

            A.TransformGroup transformGroup2 = new A.TransformGroup();
            A.Offset offset8 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents8 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset2 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents2 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup2.Append(offset8);
            transformGroup2.Append(extents8);
            transformGroup2.Append(childOffset2);
            transformGroup2.Append(childExtents2);

            groupShapeProperties2.Append(transformGroup2);

            Picture picture1 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties1 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties9 = new NonVisualDrawingProperties() { Id = (UInt32Value)180U, Name = "Picture 179", Description = "A blurry image of a dark green and white background\n\nDescription automatically generated" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{E298DCBF-C458-2956-E870-6D3AEA1DBDBA}\" />");

            // nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

            nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

            nonVisualDrawingProperties9.Append(nonVisualDrawingPropertiesExtensionList1);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoGrouping = true, NoRotation = true, NoChangeAspect = true, NoMove = true, NoResize = true, NoEditPoints = true, NoAdjustHandles = true, NoChangeArrowheads = true, NoChangeShapeType = true, NoCrop = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties9 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties9);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);
            nonVisualPictureProperties1.Append(applicationNonVisualDrawingProperties9);

            BlipFill blipFill1 = new BlipFill();
            A.Blip blip1 = new A.Blip() { Embed = "rId3" };

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            ShapeProperties shapeProperties8 = new ShapeProperties();

            A.Transform2D transform2D7 = new A.Transform2D();
            A.Offset offset9 = new A.Offset() { X = 12241L, Y = -14682L };
            A.Extents extents9 = new A.Extents() { Cx = 12188190L, Cy = 6860144L };

            transform2D7.Append(offset9);
            transform2D7.Append(extents9);

            A.PresetGeometry presetGeometry7 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList7 = new A.AdjustValueList();

            presetGeometry7.Append(adjustValueList7);

            shapeProperties8.Append(transform2D7);
            shapeProperties8.Append(presetGeometry7);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties8);

            Shape shape7 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties7 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties10 = new NonVisualDrawingProperties() { Id = (UInt32Value)172U, Name = "Shape 125" };
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties7 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties10 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties7.Append(nonVisualDrawingProperties10);
            nonVisualShapeProperties7.Append(nonVisualShapeDrawingProperties7);
            nonVisualShapeProperties7.Append(applicationNonVisualDrawingProperties10);

            ShapeProperties shapeProperties9 = new ShapeProperties();

            A.Transform2D transform2D8 = new A.Transform2D();
            A.Offset offset10 = new A.Offset() { X = 686195L, Y = 697634L };
            A.Extents extents10 = new A.Extents() { Cx = 10852432L, Cy = 365713L };

            transform2D8.Append(offset10);
            transform2D8.Append(extents10);

            A.PresetGeometry presetGeometry8 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList8 = new A.AdjustValueList();

            presetGeometry8.Append(adjustValueList8);
            A.NoFill noFill2 = new A.NoFill();
            A.Outline outline5 = new A.Outline();

            shapeProperties9.Append(transform2D8);
            shapeProperties9.Append(presetGeometry8);
            shapeProperties9.Append(noFill2);
            shapeProperties9.Append(outline5);

            TextBody textBody7 = new TextBody();
            A.BodyProperties bodyProperties8 = new A.BodyProperties();
            A.ListStyle listStyle8 = new A.ListStyle();

            A.Paragraph paragraph11 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties6 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1620 };

            paragraph11.Append(endParagraphRunProperties6);

            textBody7.Append(bodyProperties8);
            textBody7.Append(listStyle8);
            textBody7.Append(paragraph11);

            shape7.Append(nonVisualShapeProperties7);
            shape7.Append(shapeProperties9);
            shape7.Append(textBody7);

            Shape shape8 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties8 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties11 = new NonVisualDrawingProperties() { Id = (UInt32Value)173U, Name = "Text 126" };
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties8 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties11 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties8.Append(nonVisualDrawingProperties11);
            nonVisualShapeProperties8.Append(nonVisualShapeDrawingProperties8);
            nonVisualShapeProperties8.Append(applicationNonVisualDrawingProperties11);

            ShapeProperties shapeProperties10 = new ShapeProperties();

            A.Transform2D transform2D9 = new A.Transform2D();
            A.Offset offset11 = new A.Offset() { X = 686194L, Y = 495574L };
            A.Extents extents11 = new A.Extents() { Cx = 8639872L, Cy = 457142L };

            transform2D9.Append(offset11);
            transform2D9.Append(extents11);

            A.PresetGeometry presetGeometry9 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList9 = new A.AdjustValueList();

            presetGeometry9.Append(adjustValueList9);
            A.NoFill noFill3 = new A.NoFill();
            A.Outline outline6 = new A.Outline();

            shapeProperties10.Append(transform2D9);
            shapeProperties10.Append(presetGeometry9);
            shapeProperties10.Append(noFill3);
            shapeProperties10.Append(outline6);

            TextBody textBody8 = new TextBody();
            A.BodyProperties bodyProperties9 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle9 = new A.ListStyle();

            A.Paragraph paragraph12 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties6 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing1 = new A.LineSpacing();
            A.SpacingPercent spacingPercent1 = new A.SpacingPercent() { Val = 85744 };

            lineSpacing1.Append(spacingPercent1);

            paragraphProperties6.Append(lineSpacing1);

            A.Run run6 = new A.Run();

            A.RunProperties runProperties8 = new A.RunProperties() { Language = "en-US", FontSize = 2000 };

            A.SolidFill solidFill26 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "242424" };
            A.Alpha alpha2 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex12.Append(alpha2);

            solidFill26.Append(rgbColorModelHex12);
            A.LatinFont latinFont21 = new A.LatinFont() { Typeface = "Segoe UI Semibold", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont21 = new A.EastAsianFont() { Typeface = "Segoe UI Semibold", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont21 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", PitchFamily = 34, CharacterSet = -120 };

            runProperties8.Append(solidFill26);
            runProperties8.Append(latinFont21);
            runProperties8.Append(eastAsianFont21);
            runProperties8.Append(complexScriptFont21);
            A.Text text8 = new A.Text();
            text8.Text = "Monthly Overview";

            run6.Append(runProperties8);
            run6.Append(text8);
            A.EndParagraphRunProperties endParagraphRunProperties7 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 2000 };

            paragraph12.Append(paragraphProperties6);
            paragraph12.Append(run6);
            paragraph12.Append(endParagraphRunProperties7);

            textBody8.Append(bodyProperties9);
            textBody8.Append(listStyle9);
            textBody8.Append(paragraph12);

            shape8.Append(nonVisualShapeProperties8);
            shape8.Append(shapeProperties10);
            shape8.Append(textBody8);

            Shape shape9 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties9 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties12 = new NonVisualDrawingProperties() { Id = (UInt32Value)178U, Name = "Text 52" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList2 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension2 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{AC4AD6D9-E7A7-7E4C-56EF-1A76D1B5BC69}\" />");

            // nonVisualDrawingPropertiesExtension2.Append(openXmlUnknownElement2);

            nonVisualDrawingPropertiesExtensionList2.Append(nonVisualDrawingPropertiesExtension2);

            nonVisualDrawingProperties12.Append(nonVisualDrawingPropertiesExtensionList2);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties9 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties12 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties9.Append(nonVisualDrawingProperties12);
            nonVisualShapeProperties9.Append(nonVisualShapeDrawingProperties9);
            nonVisualShapeProperties9.Append(applicationNonVisualDrawingProperties12);

            ShapeProperties shapeProperties11 = new ShapeProperties();

            A.Transform2D transform2D10 = new A.Transform2D();
            A.Offset offset12 = new A.Offset() { X = 663320L, Y = 6469053L };
            A.Extents extents12 = new A.Extents() { Cx = 3888430L, Cy = 228571L };

            transform2D10.Append(offset12);
            transform2D10.Append(extents12);

            A.PresetGeometry presetGeometry10 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList10 = new A.AdjustValueList();

            presetGeometry10.Append(adjustValueList10);
            A.NoFill noFill4 = new A.NoFill();
            A.Outline outline7 = new A.Outline();

            shapeProperties11.Append(transform2D10);
            shapeProperties11.Append(presetGeometry10);
            shapeProperties11.Append(noFill4);
            shapeProperties11.Append(outline7);

            TextBody textBody9 = new TextBody();
            A.BodyProperties bodyProperties10 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle10 = new A.ListStyle();

            A.Paragraph paragraph13 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties7 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing2 = new A.LineSpacing();
            A.SpacingPercent spacingPercent2 = new A.SpacingPercent() { Val = 85714 };

            lineSpacing2.Append(spacingPercent2);

            paragraphProperties7.Append(lineSpacing2);

            A.Run run7 = new A.Run();

            A.RunProperties runProperties9 = new A.RunProperties() { Language = "en-US", FontSize = 1000 };

            A.SolidFill solidFill27 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "616161" };
            A.Alpha alpha3 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex13.Append(alpha3);

            solidFill27.Append(rgbColorModelHex13);
            A.LatinFont latinFont22 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont22 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont22 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties9.Append(solidFill27);
            runProperties9.Append(latinFont22);
            runProperties9.Append(eastAsianFont22);
            runProperties9.Append(complexScriptFont22);
            A.Text text9 = new A.Text();
            text9.Text = "Exported from Viva Goals on March 3, 2024";

            run7.Append(runProperties9);
            run7.Append(text9);

            A.Run run8 = new A.Run();

            A.RunProperties runProperties10 = new A.RunProperties() { Language = "en-US", FontSize = 1000 };

            A.SolidFill solidFill28 = new A.SolidFill();

            A.SchemeColor schemeColor39 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation9 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset1 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor39.Append(luminanceModulation9);
            schemeColor39.Append(luminanceOffset1);

            solidFill28.Append(schemeColor39);
            A.LatinFont latinFont23 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont23 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont23 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties10.Append(solidFill28);
            runProperties10.Append(latinFont23);
            runProperties10.Append(eastAsianFont23);
            runProperties10.Append(complexScriptFont23);
            A.Text text10 = new A.Text();
            text10.Text = ". ";

            run8.Append(runProperties10);
            run8.Append(text10);

            A.Run run9 = new A.Run();

            A.RunProperties runProperties11 = new A.RunProperties() { Language = "en-US", FontSize = 1000, Underline = A.TextUnderlineValues.Single };

            A.SolidFill solidFill29 = new A.SolidFill();

            A.SchemeColor schemeColor40 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation10 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset2 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor40.Append(luminanceModulation10);
            schemeColor40.Append(luminanceOffset2);

            solidFill29.Append(schemeColor40);
            A.LatinFont latinFont24 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont24 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont24 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties11.Append(solidFill29);
            runProperties11.Append(latinFont24);
            runProperties11.Append(eastAsianFont24);
            runProperties11.Append(complexScriptFont24);
            A.Text text11 = new A.Text();
            text11.Text = "Open in Viva Goals";

            run9.Append(runProperties11);
            run9.Append(text11);

            A.EndParagraphRunProperties endParagraphRunProperties8 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1000, Underline = A.TextUnderlineValues.Single };

            A.SolidFill solidFill30 = new A.SolidFill();

            A.SchemeColor schemeColor41 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.LuminanceModulation luminanceModulation11 = new A.LuminanceModulation() { Val = 75000 };
            A.LuminanceOffset luminanceOffset3 = new A.LuminanceOffset() { Val = 25000 };

            schemeColor41.Append(luminanceModulation11);
            schemeColor41.Append(luminanceOffset3);

            solidFill30.Append(schemeColor41);

            endParagraphRunProperties8.Append(solidFill30);

            paragraph13.Append(paragraphProperties7);
            paragraph13.Append(run7);
            paragraph13.Append(run8);
            paragraph13.Append(run9);
            paragraph13.Append(endParagraphRunProperties8);

            textBody9.Append(bodyProperties10);
            textBody9.Append(listStyle10);
            textBody9.Append(paragraph13);

            shape9.Append(nonVisualShapeProperties9);
            shape9.Append(shapeProperties11);
            shape9.Append(textBody9);

            Picture picture2 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties2 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties13 = new NonVisualDrawingProperties() { Id = (UInt32Value)181U, Name = "Picture 180" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList3 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension3 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement3 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{FE8A5E5C-713E-5E5A-9D49-4FBC8D902FBF}\" />");

            // nonVisualDrawingPropertiesExtension3.Append(openXmlUnknownElement3);

            nonVisualDrawingPropertiesExtensionList3.Append(nonVisualDrawingPropertiesExtension3);

            nonVisualDrawingProperties13.Append(nonVisualDrawingPropertiesExtensionList3);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties2 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks2 = new A.PictureLocks() { NoGrouping = true, NoRotation = true, NoMove = true, NoResize = true, NoEditPoints = true, NoAdjustHandles = true, NoChangeArrowheads = true, NoChangeShapeType = true, NoCrop = true };

            nonVisualPictureDrawingProperties2.Append(pictureLocks2);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties13 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties2.Append(nonVisualDrawingProperties13);
            nonVisualPictureProperties2.Append(nonVisualPictureDrawingProperties2);
            nonVisualPictureProperties2.Append(applicationNonVisualDrawingProperties13);

            BlipFill blipFill2 = new BlipFill();
            A.Blip blip2 = new A.Blip() { Embed = "rId4" };

            A.Stretch stretch2 = new A.Stretch();
            A.FillRectangle fillRectangle2 = new A.FillRectangle();

            stretch2.Append(fillRectangle2);

            blipFill2.Append(blip2);
            blipFill2.Append(stretch2);

            ShapeProperties shapeProperties12 = new ShapeProperties();

            A.Transform2D transform2D11 = new A.Transform2D();
            A.Offset offset13 = new A.Offset() { X = 223167L, Y = 223408L };
            A.Extents extents13 = new A.Extents() { Cx = 203200L, Cy = 203200L };

            transform2D11.Append(offset13);
            transform2D11.Append(extents13);

            A.PresetGeometry presetGeometry11 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList11 = new A.AdjustValueList();

            presetGeometry11.Append(adjustValueList11);

            shapeProperties12.Append(transform2D11);
            shapeProperties12.Append(presetGeometry11);

            picture2.Append(nonVisualPictureProperties2);
            picture2.Append(blipFill2);
            picture2.Append(shapeProperties12);

            GraphicFrame graphicFrame1 = new GraphicFrame();

            NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties1 = new NonVisualGraphicFrameProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties14 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Table 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList4 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension4 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement4 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{B45883C4-42C0-0D76-223C-38CCB477A61E}\" />");

            // nonVisualDrawingPropertiesExtension4.Append(openXmlUnknownElement4);

            nonVisualDrawingPropertiesExtensionList4.Append(nonVisualDrawingPropertiesExtension4);

            nonVisualDrawingProperties14.Append(nonVisualDrawingPropertiesExtensionList4);

            NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new NonVisualGraphicFrameDrawingProperties();
            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoGrouping = true };

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties14 = new ApplicationNonVisualDrawingProperties();

            ApplicationNonVisualDrawingPropertiesExtensionList applicationNonVisualDrawingPropertiesExtensionList1 = new ApplicationNonVisualDrawingPropertiesExtensionList();

            ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 = new ApplicationNonVisualDrawingPropertiesExtension() { Uri = "{D42A27DB-BD31-4B8C-83A1-F6EECF244321}" };

            P14.ModificationId modificationId1 = new P14.ModificationId() { Val = (UInt32Value)3988155089U };
            modificationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            applicationNonVisualDrawingPropertiesExtension1.Append(modificationId1);

            applicationNonVisualDrawingPropertiesExtensionList1.Append(applicationNonVisualDrawingPropertiesExtension1);

            applicationNonVisualDrawingProperties14.Append(applicationNonVisualDrawingPropertiesExtensionList1);

            nonVisualGraphicFrameProperties1.Append(nonVisualDrawingProperties14);
            nonVisualGraphicFrameProperties1.Append(nonVisualGraphicFrameDrawingProperties1);
            nonVisualGraphicFrameProperties1.Append(applicationNonVisualDrawingProperties14);

            Transform transform1 = new Transform();
            A.Offset offset14 = new A.Offset() { X = 663320L, Y = 1055367L };
            A.Extents extents14 = new A.Extents() { Cx = 10729953L, Cy = 5187778L };

            transform1.Append(offset14);
            transform1.Append(extents14);

            A.Graphic graphic1 = new A.Graphic();

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" };

            A.Table table1 = new A.Table();

            A.TableProperties tableProperties1 = new A.TableProperties() { FirstRow = true, BandRow = true };

            A.EffectList effectList4 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow() { BlurRadius = 38100L, HorizontalRatio = 101000, VerticalRatio = 101000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.PresetColor presetColor2 = new A.PresetColor() { Val = A.PresetColorValues.Black };
            A.Alpha alpha4 = new A.Alpha() { Val = 3000 };

            presetColor2.Append(alpha4);

            outerShadow2.Append(presetColor2);

            effectList4.Append(outerShadow2);
            A.TableStyleId tableStyleId1 = new A.TableStyleId();
            tableStyleId1.Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";

            tableProperties1.Append(effectList4);
            tableProperties1.Append(tableStyleId1);

            A.TableGrid tableGrid1 = new A.TableGrid();

            A.GridColumn gridColumn1 = new A.GridColumn() { Width = 7646290L };

            A.ExtensionList extensionList1 = new A.ExtensionList();

            A.Extension extension1 = new A.Extension() { Uri = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement5 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:colId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"3835035813\" />");

            // extension1.Append(openXmlUnknownElement5);

            extensionList1.Append(extension1);

            gridColumn1.Append(extensionList1);

            A.GridColumn gridColumn2 = new A.GridColumn() { Width = 1828800L };

            A.ExtensionList extensionList2 = new A.ExtensionList();

            A.Extension extension2 = new A.Extension() { Uri = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement6 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:colId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"622552574\" />");

            // extension2.Append(openXmlUnknownElement6);

            extensionList2.Append(extension2);

            gridColumn2.Append(extensionList2);

            A.GridColumn gridColumn3 = new A.GridColumn() { Width = 1254863L };

            A.ExtensionList extensionList3 = new A.ExtensionList();

            A.Extension extension3 = new A.Extension() { Uri = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement7 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:colId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"3617387509\" />");

            // extension3.Append(openXmlUnknownElement7);

            extensionList3.Append(extension3);

            gridColumn3.Append(extensionList3);

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);

            A.TableRow tableRow1 = new A.TableRow() { Height = 475571L };

            A.TableCell tableCell1 = new A.TableCell();

            A.TextBody textBody10 = new A.TextBody();
            A.BodyProperties bodyProperties11 = new A.BodyProperties();
            A.ListStyle listStyle11 = new A.ListStyle();

            A.Paragraph paragraph14 = new A.Paragraph();

            A.Run run10 = new A.Run();

            A.RunProperties runProperties12 = new A.RunProperties() { Language = "en-US", FontSize = 1000, Bold = false, Italic = false };

            A.SolidFill solidFill31 = new A.SolidFill();
            A.SchemeColor schemeColor42 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill31.Append(schemeColor42);
            A.LatinFont latinFont25 = new A.LatinFont() { Typeface = "Segoe UI" };
            A.ComplexScriptFont complexScriptFont25 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            runProperties12.Append(solidFill31);
            runProperties12.Append(latinFont25);
            runProperties12.Append(complexScriptFont25);
            A.Text text12 = new A.Text();
            text12.Text = "Goals";

            run10.Append(runProperties12);
            run10.Append(text12);

            paragraph14.Append(run10);

            textBody10.Append(bodyProperties11);
            textBody10.Append(listStyle11);
            textBody10.Append(paragraph14);

            A.TableCellProperties tableCellProperties1 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties1 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill5 = new A.NoFill();

            leftBorderLineProperties1.Append(noFill5);

            A.RightBorderLineProperties rightBorderLineProperties1 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill6 = new A.NoFill();

            rightBorderLineProperties1.Append(noFill6);

            A.TopBorderLineProperties topBorderLineProperties1 = new A.TopBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill7 = new A.NoFill();

            topBorderLineProperties1.Append(noFill7);

            A.BottomBorderLineProperties bottomBorderLineProperties1 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill8 = new A.NoFill();
            A.PresetDash presetDash4 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round1 = new A.Round();
            A.HeadEnd headEnd1 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd1 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties1.Append(noFill8);
            bottomBorderLineProperties1.Append(presetDash4);
            bottomBorderLineProperties1.Append(round1);
            bottomBorderLineProperties1.Append(headEnd1);
            bottomBorderLineProperties1.Append(tailEnd1);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties1 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill9 = new A.NoFill();
            A.PresetDash presetDash5 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties1.Append(noFill9);
            topLeftToBottomRightBorderLineProperties1.Append(presetDash5);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties1 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill10 = new A.NoFill();
            A.PresetDash presetDash6 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties1.Append(noFill10);
            bottomLeftToTopRightBorderLineProperties1.Append(presetDash6);

            A.SolidFill solidFill32 = new A.SolidFill();
            A.SchemeColor schemeColor43 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill32.Append(schemeColor43);

            tableCellProperties1.Append(leftBorderLineProperties1);
            tableCellProperties1.Append(rightBorderLineProperties1);
            tableCellProperties1.Append(topBorderLineProperties1);
            tableCellProperties1.Append(bottomBorderLineProperties1);
            tableCellProperties1.Append(topLeftToBottomRightBorderLineProperties1);
            tableCellProperties1.Append(bottomLeftToTopRightBorderLineProperties1);
            tableCellProperties1.Append(solidFill32);

            tableCell1.Append(textBody10);
            tableCell1.Append(tableCellProperties1);

            A.TableCell tableCell2 = new A.TableCell();

            A.TextBody textBody11 = new A.TextBody();
            A.BodyProperties bodyProperties12 = new A.BodyProperties();
            A.ListStyle listStyle12 = new A.ListStyle();

            A.Paragraph paragraph15 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties8 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

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
            A.DefaultRunProperties defaultRunProperties24 = new A.DefaultRunProperties();

            paragraphProperties8.Append(lineSpacing3);
            paragraphProperties8.Append(spaceBefore1);
            paragraphProperties8.Append(spaceAfter1);
            paragraphProperties8.Append(bulletColorText1);
            paragraphProperties8.Append(bulletSizeText1);
            paragraphProperties8.Append(bulletFontText1);
            paragraphProperties8.Append(noBullet1);
            paragraphProperties8.Append(tabStopList1);
            paragraphProperties8.Append(defaultRunProperties24);

            A.Run run11 = new A.Run();

            A.RunProperties runProperties13 = new A.RunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1000, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline8 = new A.Outline();
            A.NoFill noFill11 = new A.NoFill();

            outline8.Append(noFill11);

            A.SolidFill solidFill33 = new A.SolidFill();
            A.PresetColor presetColor3 = new A.PresetColor() { Val = A.PresetColorValues.Black };

            solidFill33.Append(presetColor3);
            A.EffectList effectList5 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText1 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText1 = new A.UnderlineFillText();
            A.LatinFont latinFont26 = new A.LatinFont() { Typeface = "Segoe UI" };
            A.EastAsianFont eastAsianFont25 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont26 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            runProperties13.Append(outline8);
            runProperties13.Append(solidFill33);
            runProperties13.Append(effectList5);
            runProperties13.Append(underlineFollowsText1);
            runProperties13.Append(underlineFillText1);
            runProperties13.Append(latinFont26);
            runProperties13.Append(eastAsianFont25);
            runProperties13.Append(complexScriptFont26);
            A.Text text13 = new A.Text();
            text13.Text = "Owner";

            run11.Append(runProperties13);
            run11.Append(text13);

            paragraph15.Append(paragraphProperties8);
            paragraph15.Append(run11);

            textBody11.Append(bodyProperties12);
            textBody11.Append(listStyle12);
            textBody11.Append(paragraph15);

            A.TableCellProperties tableCellProperties2 = new A.TableCellProperties() { LeftMargin = 82296, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties2 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill12 = new A.NoFill();

            leftBorderLineProperties2.Append(noFill12);

            A.RightBorderLineProperties rightBorderLineProperties2 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill13 = new A.NoFill();

            rightBorderLineProperties2.Append(noFill13);

            A.TopBorderLineProperties topBorderLineProperties2 = new A.TopBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill14 = new A.NoFill();

            topBorderLineProperties2.Append(noFill14);

            A.BottomBorderLineProperties bottomBorderLineProperties2 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill15 = new A.NoFill();
            A.PresetDash presetDash7 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round2 = new A.Round();
            A.HeadEnd headEnd2 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd2 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties2.Append(noFill15);
            bottomBorderLineProperties2.Append(presetDash7);
            bottomBorderLineProperties2.Append(round2);
            bottomBorderLineProperties2.Append(headEnd2);
            bottomBorderLineProperties2.Append(tailEnd2);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties2 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill16 = new A.NoFill();
            A.PresetDash presetDash8 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties2.Append(noFill16);
            topLeftToBottomRightBorderLineProperties2.Append(presetDash8);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties2 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill17 = new A.NoFill();
            A.PresetDash presetDash9 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties2.Append(noFill17);
            bottomLeftToTopRightBorderLineProperties2.Append(presetDash9);

            A.SolidFill solidFill34 = new A.SolidFill();
            A.SchemeColor schemeColor44 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill34.Append(schemeColor44);

            tableCellProperties2.Append(leftBorderLineProperties2);
            tableCellProperties2.Append(rightBorderLineProperties2);
            tableCellProperties2.Append(topBorderLineProperties2);
            tableCellProperties2.Append(bottomBorderLineProperties2);
            tableCellProperties2.Append(topLeftToBottomRightBorderLineProperties2);
            tableCellProperties2.Append(bottomLeftToTopRightBorderLineProperties2);
            tableCellProperties2.Append(solidFill34);

            tableCell2.Append(textBody11);
            tableCell2.Append(tableCellProperties2);

            A.TableCell tableCell3 = new A.TableCell();

            A.TextBody textBody12 = new A.TextBody();
            A.BodyProperties bodyProperties13 = new A.BodyProperties();
            A.ListStyle listStyle13 = new A.ListStyle();

            A.Paragraph paragraph16 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties9 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

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
            A.DefaultRunProperties defaultRunProperties25 = new A.DefaultRunProperties();

            paragraphProperties9.Append(lineSpacing4);
            paragraphProperties9.Append(spaceBefore2);
            paragraphProperties9.Append(spaceAfter2);
            paragraphProperties9.Append(bulletColorText2);
            paragraphProperties9.Append(bulletSizeText2);
            paragraphProperties9.Append(bulletFontText2);
            paragraphProperties9.Append(noBullet2);
            paragraphProperties9.Append(tabStopList2);
            paragraphProperties9.Append(defaultRunProperties25);

            A.Run run12 = new A.Run();

            A.RunProperties runProperties14 = new A.RunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1000, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline9 = new A.Outline();
            A.NoFill noFill18 = new A.NoFill();

            outline9.Append(noFill18);

            A.SolidFill solidFill35 = new A.SolidFill();
            A.SchemeColor schemeColor45 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill35.Append(schemeColor45);
            A.EffectList effectList6 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText2 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText2 = new A.UnderlineFillText();
            A.LatinFont latinFont27 = new A.LatinFont() { Typeface = "Segoe UI" };
            A.EastAsianFont eastAsianFont26 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont27 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            runProperties14.Append(outline9);
            runProperties14.Append(solidFill35);
            runProperties14.Append(effectList6);
            runProperties14.Append(underlineFollowsText2);
            runProperties14.Append(underlineFillText2);
            runProperties14.Append(latinFont27);
            runProperties14.Append(eastAsianFont26);
            runProperties14.Append(complexScriptFont27);
            A.Text text14 = new A.Text();
            text14.Text = "Target";

            run12.Append(runProperties14);
            run12.Append(text14);

            paragraph16.Append(paragraphProperties9);
            paragraph16.Append(run12);

            textBody12.Append(bodyProperties13);
            textBody12.Append(listStyle13);
            textBody12.Append(paragraph16);

            A.TableCellProperties tableCellProperties3 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties3 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill19 = new A.NoFill();

            leftBorderLineProperties3.Append(noFill19);

            A.RightBorderLineProperties rightBorderLineProperties3 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill20 = new A.NoFill();

            rightBorderLineProperties3.Append(noFill20);

            A.TopBorderLineProperties topBorderLineProperties3 = new A.TopBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill21 = new A.NoFill();

            topBorderLineProperties3.Append(noFill21);

            A.BottomBorderLineProperties bottomBorderLineProperties3 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill22 = new A.NoFill();
            A.PresetDash presetDash10 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round3 = new A.Round();
            A.HeadEnd headEnd3 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd3 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties3.Append(noFill22);
            bottomBorderLineProperties3.Append(presetDash10);
            bottomBorderLineProperties3.Append(round3);
            bottomBorderLineProperties3.Append(headEnd3);
            bottomBorderLineProperties3.Append(tailEnd3);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties3 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill23 = new A.NoFill();
            A.PresetDash presetDash11 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties3.Append(noFill23);
            topLeftToBottomRightBorderLineProperties3.Append(presetDash11);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties3 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill24 = new A.NoFill();
            A.PresetDash presetDash12 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties3.Append(noFill24);
            bottomLeftToTopRightBorderLineProperties3.Append(presetDash12);

            A.SolidFill solidFill36 = new A.SolidFill();
            A.SchemeColor schemeColor46 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill36.Append(schemeColor46);

            tableCellProperties3.Append(leftBorderLineProperties3);
            tableCellProperties3.Append(rightBorderLineProperties3);
            tableCellProperties3.Append(topBorderLineProperties3);
            tableCellProperties3.Append(bottomBorderLineProperties3);
            tableCellProperties3.Append(topLeftToBottomRightBorderLineProperties3);
            tableCellProperties3.Append(bottomLeftToTopRightBorderLineProperties3);
            tableCellProperties3.Append(solidFill36);

            tableCell3.Append(textBody12);
            tableCell3.Append(tableCellProperties3);

            A.ExtensionList extensionList4 = new A.ExtensionList();

            A.Extension extension4 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement8 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"2429191634\" />");

            // extension4.Append(openXmlUnknownElement8);

            extensionList4.Append(extension4);

            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);
            tableRow1.Append(extensionList4);

            A.TableRow tableRow2 = new A.TableRow() { Height = 781011L };

            A.TableCell tableCell4 = new A.TableCell();

            A.TextBody textBody13 = new A.TextBody();
            A.BodyProperties bodyProperties14 = new A.BodyProperties();
            A.ListStyle listStyle14 = new A.ListStyle();

            A.Paragraph paragraph17 = new A.Paragraph();

            A.Run run13 = new A.Run();

            A.RunProperties runProperties15 = new A.RunProperties() { Language = "en-US", FontSize = 1300, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Dirty = false };

            A.SolidFill solidFill37 = new A.SolidFill();
            A.SchemeColor schemeColor47 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            solidFill37.Append(schemeColor47);
            A.EffectList effectList7 = new A.EffectList();
            string fontName = goalsArrayForSlide.ElementAt(0).isMainGoal ? "Segoe UI Semibold" : "Segoe UI";
            A.LatinFont latinFont28 = new A.LatinFont() { Typeface = "Segoe UI Semibold" };
            A.EastAsianFont eastAsianFont27 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont28 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            runProperties15.Append(solidFill37);
            runProperties15.Append(effectList7);
            runProperties15.Append(latinFont28);
            runProperties15.Append(eastAsianFont27);
            runProperties15.Append(complexScriptFont28);
            A.Text text15 = new A.Text();


            text15.Text = goalsArrayForSlide.ElementAt(0).goalName.ToString() ?? "";

            run13.Append(runProperties15);
            run13.Append(text15);

            A.Run run14 = new A.Run();

            A.RunProperties runProperties16 = new A.RunProperties() { Language = "en-US", FontSize = 1300, Bold = true, Italic = false, Kerning = 1200, Dirty = false };

            A.SolidFill solidFill38 = new A.SolidFill();
            A.SchemeColor schemeColor48 = new A.SchemeColor() { Val = A.SchemeColorValues.Dark1 };

            solidFill38.Append(schemeColor48);
            A.EffectList effectList8 = new A.EffectList();
            A.LatinFont latinFont29 = new A.LatinFont() { Typeface = "Segoe UI Semibold" };
            A.EastAsianFont eastAsianFont28 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont29 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            runProperties16.Append(solidFill38);
            runProperties16.Append(effectList8);
            runProperties16.Append(latinFont29);
            runProperties16.Append(eastAsianFont28);
            runProperties16.Append(complexScriptFont29);
            A.Text text16 = new A.Text();
            text16.Text = "​";

            run14.Append(runProperties16);
            run14.Append(text16);

            A.EndParagraphRunProperties endParagraphRunProperties9 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1300, Bold = true, Italic = false, Dirty = false };

            A.SolidFill solidFill39 = new A.SolidFill();
            A.SchemeColor schemeColor49 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill39.Append(schemeColor49);
            A.EffectList effectList9 = new A.EffectList();
            A.LatinFont latinFont30 = new A.LatinFont() { Typeface = "Segoe UI Semibold" };
            A.ComplexScriptFont complexScriptFont30 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            endParagraphRunProperties9.Append(solidFill39);
            endParagraphRunProperties9.Append(effectList9);
            endParagraphRunProperties9.Append(latinFont30);
            endParagraphRunProperties9.Append(complexScriptFont30);

            paragraph17.Append(run13);
            paragraph17.Append(run14);
            paragraph17.Append(endParagraphRunProperties9);

            textBody13.Append(bodyProperties14);
            textBody13.Append(listStyle14);
            textBody13.Append(paragraph17);

            A.TableCellProperties tableCellProperties4 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties4 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill25 = new A.NoFill();

            leftBorderLineProperties4.Append(noFill25);

            A.RightBorderLineProperties rightBorderLineProperties4 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill26 = new A.NoFill();

            rightBorderLineProperties4.Append(noFill26);

            A.TopBorderLineProperties topBorderLineProperties4 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill27 = new A.NoFill();
            A.PresetDash presetDash13 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round4 = new A.Round();
            A.HeadEnd headEnd4 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd4 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties4.Append(noFill27);
            topBorderLineProperties4.Append(presetDash13);
            topBorderLineProperties4.Append(round4);
            topBorderLineProperties4.Append(headEnd4);
            topBorderLineProperties4.Append(tailEnd4);

            A.BottomBorderLineProperties bottomBorderLineProperties4 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill28 = new A.NoFill();
            A.PresetDash presetDash14 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round5 = new A.Round();
            A.HeadEnd headEnd5 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd5 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties4.Append(noFill28);
            bottomBorderLineProperties4.Append(presetDash14);
            bottomBorderLineProperties4.Append(round5);
            bottomBorderLineProperties4.Append(headEnd5);
            bottomBorderLineProperties4.Append(tailEnd5);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties4 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill29 = new A.NoFill();
            A.PresetDash presetDash15 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties4.Append(noFill29);
            topLeftToBottomRightBorderLineProperties4.Append(presetDash15);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties4 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill30 = new A.NoFill();
            A.PresetDash presetDash16 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties4.Append(noFill30);
            bottomLeftToTopRightBorderLineProperties4.Append(presetDash16);

            A.SolidFill solidFill40 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex14 = new A.RgbColorModelHex() { Val = "F5F5F5" };

            solidFill40.Append(rgbColorModelHex14);

            tableCellProperties4.Append(leftBorderLineProperties4);
            tableCellProperties4.Append(rightBorderLineProperties4);
            tableCellProperties4.Append(topBorderLineProperties4);
            tableCellProperties4.Append(bottomBorderLineProperties4);
            tableCellProperties4.Append(topLeftToBottomRightBorderLineProperties4);
            tableCellProperties4.Append(bottomLeftToTopRightBorderLineProperties4);
            tableCellProperties4.Append(solidFill40);

            tableCell4.Append(textBody13);
            tableCell4.Append(tableCellProperties4);

            A.TableCell tableCell5 = new A.TableCell();

            A.TextBody textBody14 = new A.TextBody();
            A.BodyProperties bodyProperties15 = new A.BodyProperties();
            A.ListStyle listStyle15 = new A.ListStyle();

            A.Paragraph paragraph18 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties10 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

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
            A.DefaultRunProperties defaultRunProperties26 = new A.DefaultRunProperties();

            paragraphProperties10.Append(lineSpacing5);
            paragraphProperties10.Append(spaceBefore3);
            paragraphProperties10.Append(spaceAfter3);
            paragraphProperties10.Append(bulletColorText3);
            paragraphProperties10.Append(bulletSizeText3);
            paragraphProperties10.Append(bulletFontText3);
            paragraphProperties10.Append(noBullet3);
            paragraphProperties10.Append(tabStopList3);
            paragraphProperties10.Append(defaultRunProperties26);

            A.EndParagraphRunProperties endParagraphRunProperties10 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline10 = new A.Outline();
            A.NoFill noFill31 = new A.NoFill();

            outline10.Append(noFill31);

            A.SolidFill solidFill41 = new A.SolidFill();
            A.PresetColor presetColor4 = new A.PresetColor() { Val = A.PresetColorValues.Black };

            solidFill41.Append(presetColor4);
            A.EffectList effectList10 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText3 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText3 = new A.UnderlineFillText();
            A.LatinFont latinFont31 = new A.LatinFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont29 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont31 = new A.ComplexScriptFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties10.Append(outline10);
            endParagraphRunProperties10.Append(solidFill41);
            endParagraphRunProperties10.Append(effectList10);
            endParagraphRunProperties10.Append(underlineFollowsText3);
            endParagraphRunProperties10.Append(underlineFillText3);
            endParagraphRunProperties10.Append(latinFont31);
            endParagraphRunProperties10.Append(eastAsianFont29);
            endParagraphRunProperties10.Append(complexScriptFont31);

            paragraph18.Append(paragraphProperties10);
            paragraph18.Append(endParagraphRunProperties10);

            textBody14.Append(bodyProperties15);
            textBody14.Append(listStyle15);
            textBody14.Append(paragraph18);

            A.TableCellProperties tableCellProperties5 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties5 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill32 = new A.NoFill();

            leftBorderLineProperties5.Append(noFill32);

            A.RightBorderLineProperties rightBorderLineProperties5 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill33 = new A.NoFill();

            rightBorderLineProperties5.Append(noFill33);

            A.TopBorderLineProperties topBorderLineProperties5 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill34 = new A.NoFill();
            A.PresetDash presetDash17 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round6 = new A.Round();
            A.HeadEnd headEnd6 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd6 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties5.Append(noFill34);
            topBorderLineProperties5.Append(presetDash17);
            topBorderLineProperties5.Append(round6);
            topBorderLineProperties5.Append(headEnd6);
            topBorderLineProperties5.Append(tailEnd6);

            A.BottomBorderLineProperties bottomBorderLineProperties5 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill35 = new A.NoFill();
            A.PresetDash presetDash18 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round7 = new A.Round();
            A.HeadEnd headEnd7 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd7 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties5.Append(noFill35);
            bottomBorderLineProperties5.Append(presetDash18);
            bottomBorderLineProperties5.Append(round7);
            bottomBorderLineProperties5.Append(headEnd7);
            bottomBorderLineProperties5.Append(tailEnd7);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties5 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill36 = new A.NoFill();
            A.PresetDash presetDash19 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties5.Append(noFill36);
            topLeftToBottomRightBorderLineProperties5.Append(presetDash19);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties5 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill37 = new A.NoFill();
            A.PresetDash presetDash20 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties5.Append(noFill37);
            bottomLeftToTopRightBorderLineProperties5.Append(presetDash20);

            A.SolidFill solidFill42 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex15 = new A.RgbColorModelHex(){ Val = "F5F5F5" };

            solidFill42.Append(rgbColorModelHex15);

            tableCellProperties5.Append(leftBorderLineProperties5);
            tableCellProperties5.Append(rightBorderLineProperties5);
            tableCellProperties5.Append(topBorderLineProperties5);
            tableCellProperties5.Append(bottomBorderLineProperties5);
            tableCellProperties5.Append(topLeftToBottomRightBorderLineProperties5);
            tableCellProperties5.Append(bottomLeftToTopRightBorderLineProperties5);
            tableCellProperties5.Append(solidFill42);

            tableCell5.Append(textBody14);
            tableCell5.Append(tableCellProperties5);

            A.TableCell tableCell6 = new A.TableCell();

            A.TextBody textBody15 = new A.TextBody();
            A.BodyProperties bodyProperties16 = new A.BodyProperties();
            A.ListStyle listStyle16 = new A.ListStyle();

            A.Paragraph paragraph19 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties11 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

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
            A.DefaultRunProperties defaultRunProperties27 = new A.DefaultRunProperties();

            paragraphProperties11.Append(lineSpacing6);
            paragraphProperties11.Append(spaceBefore4);
            paragraphProperties11.Append(spaceAfter4);
            paragraphProperties11.Append(bulletColorText4);
            paragraphProperties11.Append(bulletSizeText4);
            paragraphProperties11.Append(bulletFontText4);
            paragraphProperties11.Append(noBullet4);
            paragraphProperties11.Append(tabStopList4);
            paragraphProperties11.Append(defaultRunProperties27);

            A.EndParagraphRunProperties endParagraphRunProperties11 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline11 = new A.Outline();
            A.NoFill noFill38 = new A.NoFill();

            outline11.Append(noFill38);

            A.SolidFill solidFill43 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex16 = new A.RgbColorModelHex() { Val = "C50F1F" };

            solidFill43.Append(rgbColorModelHex16);
            A.EffectList effectList11 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText4 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText4 = new A.UnderlineFillText();
            A.LatinFont latinFont32 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont30 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont32 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties11.Append(outline11);
            endParagraphRunProperties11.Append(solidFill43);
            endParagraphRunProperties11.Append(effectList11);
            endParagraphRunProperties11.Append(underlineFollowsText4);
            endParagraphRunProperties11.Append(underlineFillText4);
            endParagraphRunProperties11.Append(latinFont32);
            endParagraphRunProperties11.Append(eastAsianFont30);
            endParagraphRunProperties11.Append(complexScriptFont32);

            paragraph19.Append(paragraphProperties11);
            paragraph19.Append(endParagraphRunProperties11);

            textBody15.Append(bodyProperties16);
            textBody15.Append(listStyle16);
            textBody15.Append(paragraph19);

            A.TableCellProperties tableCellProperties6 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties6 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill39 = new A.NoFill();

            leftBorderLineProperties6.Append(noFill39);

            A.RightBorderLineProperties rightBorderLineProperties6 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill40 = new A.NoFill();

            rightBorderLineProperties6.Append(noFill40);

            A.TopBorderLineProperties topBorderLineProperties6 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill41 = new A.NoFill();
            A.PresetDash presetDash21 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round8 = new A.Round();
            A.HeadEnd headEnd8 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd8 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties6.Append(noFill41);
            topBorderLineProperties6.Append(presetDash21);
            topBorderLineProperties6.Append(round8);
            topBorderLineProperties6.Append(headEnd8);
            topBorderLineProperties6.Append(tailEnd8);

            A.BottomBorderLineProperties bottomBorderLineProperties6 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill42 = new A.NoFill();
            A.PresetDash presetDash22 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round9 = new A.Round();
            A.HeadEnd headEnd9 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd9 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties6.Append(noFill42);
            bottomBorderLineProperties6.Append(presetDash22);
            bottomBorderLineProperties6.Append(round9);
            bottomBorderLineProperties6.Append(headEnd9);
            bottomBorderLineProperties6.Append(tailEnd9);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties6 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill43 = new A.NoFill();
            A.PresetDash presetDash23 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties6.Append(noFill43);
            topLeftToBottomRightBorderLineProperties6.Append(presetDash23);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties6 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill44 = new A.NoFill();
            A.PresetDash presetDash24 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties6.Append(noFill44);
            bottomLeftToTopRightBorderLineProperties6.Append(presetDash24);

            A.SolidFill solidFill44 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex17 = new A.RgbColorModelHex() { Val = "F5F5F5" };

            solidFill44.Append(rgbColorModelHex17);

            tableCellProperties6.Append(leftBorderLineProperties6);
            tableCellProperties6.Append(rightBorderLineProperties6);
            tableCellProperties6.Append(topBorderLineProperties6);
            tableCellProperties6.Append(bottomBorderLineProperties6);
            tableCellProperties6.Append(topLeftToBottomRightBorderLineProperties6);
            tableCellProperties6.Append(bottomLeftToTopRightBorderLineProperties6);
            tableCellProperties6.Append(solidFill44);

            tableCell6.Append(textBody15);
            tableCell6.Append(tableCellProperties6);

            A.ExtensionList extensionList5 = new A.ExtensionList();

            A.Extension extension5 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement9 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"587416212\" />");

            // extension5.Append(openXmlUnknownElement9);

            extensionList5.Append(extension5);

            tableRow2.Append(tableCell4);
            tableRow2.Append(tableCell5);
            tableRow2.Append(tableCell6);
            tableRow2.Append(extensionList5);

            A.TableRow tableRow3 = new A.TableRow() { Height = 748571L };

            A.TableCell tableCell7 = new A.TableCell();

            A.TextBody textBody16 = new A.TextBody();
            A.BodyProperties bodyProperties17 = new A.BodyProperties();
            A.ListStyle listStyle17 = new A.ListStyle();

            A.Paragraph paragraph20 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties12 = new A.ParagraphProperties() { LeftMargin = 171450, RightMargin = 0, Indent = -171450, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing7 = new A.LineSpacing();
            A.SpacingPercent spacingPercent7 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing7.Append(spacingPercent7);

            A.SpaceBefore spaceBefore5 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints9 = new A.SpacingPoints() { Val = 0 };

            spaceBefore5.Append(spacingPoints9);

            A.SpaceAfter spaceAfter5 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints10 = new A.SpacingPoints() { Val = 0 };

            spaceAfter5.Append(spacingPoints10);
            A.BulletColorText bulletColorText5 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText5 = new A.BulletSizeText();
            A.BulletFontText bulletFontText5 = new A.BulletFontText();

            A.PictureBullet pictureBullet1 = new A.PictureBullet();

            A.Blip blip3 = new A.Blip() { Embed = "rId5" };

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement10 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId6\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension1.Append(openXmlUnknownElement10);

            blipExtensionList1.Append(blipExtension1);

            blip3.Append(blipExtensionList1);

            pictureBullet1.Append(blip3);
            A.TabStopList tabStopList5 = new A.TabStopList();
            A.DefaultRunProperties defaultRunProperties28 = new A.DefaultRunProperties();

            paragraphProperties12.Append(lineSpacing7);
            paragraphProperties12.Append(spaceBefore5);
            paragraphProperties12.Append(spaceAfter5);
            paragraphProperties12.Append(bulletColorText5);
            paragraphProperties12.Append(bulletSizeText5);
            paragraphProperties12.Append(bulletFontText5);

            // ********************************* Custom Code Start *********************************
            if (!goalsArrayForSlide.ElementAt(1).isMainGoal)
                paragraphProperties12.Append(pictureBullet1);
            // ********************************* Custom Code End *********************************

            paragraphProperties12.Append(tabStopList5);
            paragraphProperties12.Append(defaultRunProperties28);


            A.Run run15 = new A.Run();

            A.RunProperties runProperties17 = new A.RunProperties() { Language = "en-US", FontSize = goalsArrayForSlide.ElementAt(1).isMainGoal ? 1300 : 1200, Bold = goalsArrayForSlide.ElementAt(1).isMainGoal, Italic = false, Dirty = false };
            A.EffectList effectList12 = new A.EffectList();

            // ********************************* Custom Code Start *********************************
            A.LatinFont latinFont33 = new A.LatinFont() { Typeface = goalsArrayForSlide.ElementAt(1).isMainGoal ? "Segoe UI semibold" : "Segoe UI" };
            // ********************************* Custom Code End *********************************

            A.ComplexScriptFont complexScriptFont33 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            runProperties17.Append(effectList12);
            runProperties17.Append(latinFont33);
            runProperties17.Append(complexScriptFont33);
            A.Text text17 = new A.Text();

            // ********************************* Custom Code Start *********************************
            if (goalsArrayForSlide.Count() > 1)
                text17.Text = goalsArrayForSlide.ElementAt(1).goalName ?? "";
            else
                text17.Text = "NO DATA";
            // ********************************* Custom Code End *********************************

            run15.Append(runProperties17);
            run15.Append(text17);

            A.EndParagraphRunProperties endParagraphRunProperties12 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Dirty = false };
            A.LatinFont latinFont34 = new A.LatinFont() { Typeface = "Segoe UI" };
            A.ComplexScriptFont complexScriptFont34 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            endParagraphRunProperties12.Append(latinFont34);
            endParagraphRunProperties12.Append(complexScriptFont34);

            paragraph20.Append(paragraphProperties12);
            paragraph20.Append(run15);
            paragraph20.Append(endParagraphRunProperties12);

            textBody16.Append(bodyProperties17);
            textBody16.Append(listStyle17);
            textBody16.Append(paragraph20);

            A.TableCellProperties tableCellProperties7 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties7 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill45 = new A.NoFill();

            leftBorderLineProperties7.Append(noFill45);

            A.RightBorderLineProperties rightBorderLineProperties7 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill46 = new A.NoFill();

            rightBorderLineProperties7.Append(noFill46);

            A.TopBorderLineProperties topBorderLineProperties7 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill47 = new A.NoFill();
            A.PresetDash presetDash25 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round10 = new A.Round();
            A.HeadEnd headEnd10 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd10 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties7.Append(noFill47);
            topBorderLineProperties7.Append(presetDash25);
            topBorderLineProperties7.Append(round10);
            topBorderLineProperties7.Append(headEnd10);
            topBorderLineProperties7.Append(tailEnd10);

            A.BottomBorderLineProperties bottomBorderLineProperties7 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill48 = new A.NoFill();
            A.PresetDash presetDash26 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round11 = new A.Round();
            A.HeadEnd headEnd11 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd11 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties7.Append(noFill48);
            bottomBorderLineProperties7.Append(presetDash26);
            bottomBorderLineProperties7.Append(round11);
            bottomBorderLineProperties7.Append(headEnd11);
            bottomBorderLineProperties7.Append(tailEnd11);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties7 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill49 = new A.NoFill();
            A.PresetDash presetDash27 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties7.Append(noFill49);
            topLeftToBottomRightBorderLineProperties7.Append(presetDash27);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties7 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill50 = new A.NoFill();
            A.PresetDash presetDash28 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties7.Append(noFill50);
            bottomLeftToTopRightBorderLineProperties7.Append(presetDash28);

            A.SolidFill solidFill45 = new A.SolidFill();

            
            // A.SchemeColor schemeColor50 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            // ********************************* Custom Code Start *********************************
            // Adding the background color according to the goal type
            A.RgbColorModelHex rgbColorModelHex100 = new A.RgbColorModelHex() { Val = goalsArrayForSlide.ElementAt(1).isMainGoal ? "F5F5F5" : "FFFFFF" };
            // ********************************* Custom Code End *********************************

            solidFill45.Append(rgbColorModelHex100);

            tableCellProperties7.Append(leftBorderLineProperties7);
            tableCellProperties7.Append(rightBorderLineProperties7);
            tableCellProperties7.Append(topBorderLineProperties7);
            tableCellProperties7.Append(bottomBorderLineProperties7);
            tableCellProperties7.Append(topLeftToBottomRightBorderLineProperties7);
            tableCellProperties7.Append(bottomLeftToTopRightBorderLineProperties7);
            tableCellProperties7.Append(solidFill45);

            tableCell7.Append(textBody16);
            tableCell7.Append(tableCellProperties7);

            A.TableCell tableCell8 = new A.TableCell();

            A.TextBody textBody17 = new A.TextBody();
            A.BodyProperties bodyProperties18 = new A.BodyProperties();
            A.ListStyle listStyle18 = new A.ListStyle();

            A.Paragraph paragraph21 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties13 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

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
            A.DefaultRunProperties defaultRunProperties29 = new A.DefaultRunProperties();

            paragraphProperties13.Append(lineSpacing8);
            paragraphProperties13.Append(spaceBefore6);
            paragraphProperties13.Append(spaceAfter6);
            paragraphProperties13.Append(bulletColorText6);
            paragraphProperties13.Append(bulletSizeText6);
            paragraphProperties13.Append(bulletFontText6);
            paragraphProperties13.Append(noBullet5);
            paragraphProperties13.Append(tabStopList6);
            paragraphProperties13.Append(defaultRunProperties29);

            A.EndParagraphRunProperties endParagraphRunProperties13 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline12 = new A.Outline();
            A.NoFill noFill51 = new A.NoFill();

            outline12.Append(noFill51);

            A.SolidFill solidFill46 = new A.SolidFill();
            A.PresetColor presetColor5 = new A.PresetColor() { Val = A.PresetColorValues.Black };

            solidFill46.Append(presetColor5);
            A.EffectList effectList13 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText5 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText5 = new A.UnderlineFillText();
            A.LatinFont latinFont35 = new A.LatinFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont31 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont35 = new A.ComplexScriptFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties13.Append(outline12);
            endParagraphRunProperties13.Append(solidFill46);
            endParagraphRunProperties13.Append(effectList13);
            endParagraphRunProperties13.Append(underlineFollowsText5);
            endParagraphRunProperties13.Append(underlineFillText5);
            endParagraphRunProperties13.Append(latinFont35);
            endParagraphRunProperties13.Append(eastAsianFont31);
            endParagraphRunProperties13.Append(complexScriptFont35);

            paragraph21.Append(paragraphProperties13);
            paragraph21.Append(endParagraphRunProperties13);

            textBody17.Append(bodyProperties18);
            textBody17.Append(listStyle18);
            textBody17.Append(paragraph21);

            A.TableCellProperties tableCellProperties8 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties8 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill52 = new A.NoFill();

            leftBorderLineProperties8.Append(noFill52);

            A.RightBorderLineProperties rightBorderLineProperties8 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill53 = new A.NoFill();

            rightBorderLineProperties8.Append(noFill53);

            A.TopBorderLineProperties topBorderLineProperties8 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill54 = new A.NoFill();
            A.PresetDash presetDash29 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round12 = new A.Round();
            A.HeadEnd headEnd12 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd12 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties8.Append(noFill54);
            topBorderLineProperties8.Append(presetDash29);
            topBorderLineProperties8.Append(round12);
            topBorderLineProperties8.Append(headEnd12);
            topBorderLineProperties8.Append(tailEnd12);

            A.BottomBorderLineProperties bottomBorderLineProperties8 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill55 = new A.NoFill();
            A.PresetDash presetDash30 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round13 = new A.Round();
            A.HeadEnd headEnd13 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd13 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties8.Append(noFill55);
            bottomBorderLineProperties8.Append(presetDash30);
            bottomBorderLineProperties8.Append(round13);
            bottomBorderLineProperties8.Append(headEnd13);
            bottomBorderLineProperties8.Append(tailEnd13);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties8 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill56 = new A.NoFill();
            A.PresetDash presetDash31 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties8.Append(noFill56);
            topLeftToBottomRightBorderLineProperties8.Append(presetDash31);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties8 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill57 = new A.NoFill();
            A.PresetDash presetDash32 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties8.Append(noFill57);
            bottomLeftToTopRightBorderLineProperties8.Append(presetDash32);

            A.SolidFill solidFill47 = new A.SolidFill();
            A.SchemeColor schemeColor51 = new A.SchemeColor(){ Val = A.SchemeColorValues.Background1 };

            solidFill47.Append(schemeColor51);

            tableCellProperties8.Append(leftBorderLineProperties8);
            tableCellProperties8.Append(rightBorderLineProperties8);
            tableCellProperties8.Append(topBorderLineProperties8);
            tableCellProperties8.Append(bottomBorderLineProperties8);
            tableCellProperties8.Append(topLeftToBottomRightBorderLineProperties8);
            tableCellProperties8.Append(bottomLeftToTopRightBorderLineProperties8);
            tableCellProperties8.Append(solidFill47);

            tableCell8.Append(textBody17);
            tableCell8.Append(tableCellProperties8);

            A.TableCell tableCell9 = new A.TableCell();

            A.TextBody textBody18 = new A.TextBody();
            A.BodyProperties bodyProperties19 = new A.BodyProperties();
            A.ListStyle listStyle19 = new A.ListStyle();

            A.Paragraph paragraph22 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties14 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

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
            A.DefaultRunProperties defaultRunProperties30 = new A.DefaultRunProperties();

            paragraphProperties14.Append(lineSpacing9);
            paragraphProperties14.Append(spaceBefore7);
            paragraphProperties14.Append(spaceAfter7);
            paragraphProperties14.Append(bulletColorText7);
            paragraphProperties14.Append(bulletSizeText7);
            paragraphProperties14.Append(bulletFontText7);
            paragraphProperties14.Append(noBullet6);
            paragraphProperties14.Append(tabStopList7);
            paragraphProperties14.Append(defaultRunProperties30);

            A.EndParagraphRunProperties endParagraphRunProperties14 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline13 = new A.Outline();
            A.NoFill noFill58 = new A.NoFill();

            outline13.Append(noFill58);

            A.SolidFill solidFill48 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex18 = new A.RgbColorModelHex() { Val = "EAA300" };

            solidFill48.Append(rgbColorModelHex18);
            A.EffectList effectList14 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText6 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText6 = new A.UnderlineFillText();
            A.LatinFont latinFont36 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont32 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont36 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties14.Append(outline13);
            endParagraphRunProperties14.Append(solidFill48);
            endParagraphRunProperties14.Append(effectList14);
            endParagraphRunProperties14.Append(underlineFollowsText6);
            endParagraphRunProperties14.Append(underlineFillText6);
            endParagraphRunProperties14.Append(latinFont36);
            endParagraphRunProperties14.Append(eastAsianFont32);
            endParagraphRunProperties14.Append(complexScriptFont36);

            paragraph22.Append(paragraphProperties14);
            paragraph22.Append(endParagraphRunProperties14);

            textBody18.Append(bodyProperties19);
            textBody18.Append(listStyle19);
            textBody18.Append(paragraph22);

            A.TableCellProperties tableCellProperties9 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties9 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill59 = new A.NoFill();

            leftBorderLineProperties9.Append(noFill59);

            A.RightBorderLineProperties rightBorderLineProperties9 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill60 = new A.NoFill();

            rightBorderLineProperties9.Append(noFill60);

            A.TopBorderLineProperties topBorderLineProperties9 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill61 = new A.NoFill();
            A.PresetDash presetDash33 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round14 = new A.Round();
            A.HeadEnd headEnd14 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd14 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties9.Append(noFill61);
            topBorderLineProperties9.Append(presetDash33);
            topBorderLineProperties9.Append(round14);
            topBorderLineProperties9.Append(headEnd14);
            topBorderLineProperties9.Append(tailEnd14);

            A.BottomBorderLineProperties bottomBorderLineProperties9 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill62 = new A.NoFill();
            A.PresetDash presetDash34 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round15 = new A.Round();
            A.HeadEnd headEnd15 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd15 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties9.Append(noFill62);
            bottomBorderLineProperties9.Append(presetDash34);
            bottomBorderLineProperties9.Append(round15);
            bottomBorderLineProperties9.Append(headEnd15);
            bottomBorderLineProperties9.Append(tailEnd15);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties9 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill63 = new A.NoFill();
            A.PresetDash presetDash35 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties9.Append(noFill63);
            topLeftToBottomRightBorderLineProperties9.Append(presetDash35);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties9 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill64 = new A.NoFill();
            A.PresetDash presetDash36 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties9.Append(noFill64);
            bottomLeftToTopRightBorderLineProperties9.Append(presetDash36);

            A.SolidFill solidFill49 = new A.SolidFill();
            A.SchemeColor schemeColor52 = new A.SchemeColor(){ Val = A.SchemeColorValues.Background1 };

            solidFill49.Append(schemeColor52);

            tableCellProperties9.Append(leftBorderLineProperties9);
            tableCellProperties9.Append(rightBorderLineProperties9);
            tableCellProperties9.Append(topBorderLineProperties9);
            tableCellProperties9.Append(bottomBorderLineProperties9);
            tableCellProperties9.Append(topLeftToBottomRightBorderLineProperties9);
            tableCellProperties9.Append(bottomLeftToTopRightBorderLineProperties9);
            tableCellProperties9.Append(solidFill49);

            tableCell9.Append(textBody18);
            tableCell9.Append(tableCellProperties9);

            A.ExtensionList extensionList6 = new A.ExtensionList();

            A.Extension extension6 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement11 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"3520631283\" />");

            // extension6.Append(openXmlUnknownElement11);

            extensionList6.Append(extension6);

            tableRow3.Append(tableCell7);
            tableRow3.Append(tableCell8);
            tableRow3.Append(tableCell9);
            tableRow3.Append(extensionList6);

            A.TableRow tableRow4 = new A.TableRow() { Height = 766152L };

            A.TableCell tableCell10 = new A.TableCell();

            A.TextBody textBody19 = new A.TextBody();
            A.BodyProperties bodyProperties20 = new A.BodyProperties();
            A.ListStyle listStyle20 = new A.ListStyle();

            A.Paragraph paragraph23 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties15 = new A.ParagraphProperties() { LeftMargin = 171450, RightMargin = 0, Level = 0, Indent = -171450, Alignment = A.TextAlignmentTypeValues.Left, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing10 = new A.LineSpacing();
            A.SpacingPercent spacingPercent10 = new A.SpacingPercent() { Val = 100000 };

            lineSpacing10.Append(spacingPercent10);

            A.SpaceBefore spaceBefore8 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints15 = new A.SpacingPoints() { Val = 0 };

            spaceBefore8.Append(spacingPoints15);

            A.SpaceAfter spaceAfter8 = new A.SpaceAfter();
            A.SpacingPoints spacingPoints16 = new A.SpacingPoints() { Val = 0 };

            spaceAfter8.Append(spacingPoints16);
            A.BulletColorText bulletColorText8 = new A.BulletColorText();
            A.BulletSizeText bulletSizeText8 = new A.BulletSizeText();
            A.BulletFontText bulletFontText8 = new A.BulletFontText();

            A.PictureBullet pictureBullet2 = new A.PictureBullet();

            A.Blip blip4 = new A.Blip() { Embed = "rId5" };

            A.BlipExtensionList blipExtensionList2 = new A.BlipExtensionList();

            A.BlipExtension blipExtension2 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement12 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId6\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension2.Append(openXmlUnknownElement12);

            blipExtensionList2.Append(blipExtension2);

            blip4.Append(blipExtensionList2);

            pictureBullet2.Append(blip4);

            paragraphProperties15.Append(lineSpacing10);
            paragraphProperties15.Append(spaceBefore8);
            paragraphProperties15.Append(spaceAfter8);
            paragraphProperties15.Append(bulletColorText8);
            paragraphProperties15.Append(bulletSizeText8);
            paragraphProperties15.Append(bulletFontText8);
            paragraphProperties15.Append(pictureBullet2);

            A.Run run16 = new A.Run();

            A.RunProperties runProperties18 = new A.RunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Dirty = false };
            A.EffectList effectList15 = new A.EffectList();
            A.LatinFont latinFont37 = new A.LatinFont() { Typeface = "Segoe UI" };
            A.ComplexScriptFont complexScriptFont37 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            runProperties18.Append(effectList15);
            runProperties18.Append(latinFont37);
            runProperties18.Append(complexScriptFont37);
            A.Text text18 = new A.Text();

            // ********************************* Custom Code Start *********************************
            if (goalsArrayForSlide.Count() > 2)
                text18.Text = goalsArrayForSlide.ElementAt(2).goalName.ToString() ?? "";
            else
                text18.Text = "NO DATA";
            // ********************************* Custom Code End *********************************

            run16.Append(runProperties18);
            run16.Append(text18);

            paragraph23.Append(paragraphProperties15);
            paragraph23.Append(run16);

            textBody19.Append(bodyProperties20);
            textBody19.Append(listStyle20);
            textBody19.Append(paragraph23);

            A.TableCellProperties tableCellProperties10 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 365760, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties10 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill65 = new A.NoFill();

            leftBorderLineProperties10.Append(noFill65);

            A.RightBorderLineProperties rightBorderLineProperties10 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill66 = new A.NoFill();

            rightBorderLineProperties10.Append(noFill66);

            A.TopBorderLineProperties topBorderLineProperties10 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill67 = new A.NoFill();
            A.PresetDash presetDash37 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round16 = new A.Round();
            A.HeadEnd headEnd16 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd16 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties10.Append(noFill67);
            topBorderLineProperties10.Append(presetDash37);
            topBorderLineProperties10.Append(round16);
            topBorderLineProperties10.Append(headEnd16);
            topBorderLineProperties10.Append(tailEnd16);

            A.BottomBorderLineProperties bottomBorderLineProperties10 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill68 = new A.NoFill();
            A.PresetDash presetDash38 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round17 = new A.Round();
            A.HeadEnd headEnd17 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd17 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties10.Append(noFill68);
            bottomBorderLineProperties10.Append(presetDash38);
            bottomBorderLineProperties10.Append(round17);
            bottomBorderLineProperties10.Append(headEnd17);
            bottomBorderLineProperties10.Append(tailEnd17);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties10 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill69 = new A.NoFill();
            A.PresetDash presetDash39 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties10.Append(noFill69);
            topLeftToBottomRightBorderLineProperties10.Append(presetDash39);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties10 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill70 = new A.NoFill();
            A.PresetDash presetDash40 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties10.Append(noFill70);
            bottomLeftToTopRightBorderLineProperties10.Append(presetDash40);

            A.SolidFill solidFill50 = new A.SolidFill();
            A.SchemeColor schemeColor53 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill50.Append(schemeColor53);

            tableCellProperties10.Append(leftBorderLineProperties10);
            tableCellProperties10.Append(rightBorderLineProperties10);
            tableCellProperties10.Append(topBorderLineProperties10);
            tableCellProperties10.Append(bottomBorderLineProperties10);
            tableCellProperties10.Append(topLeftToBottomRightBorderLineProperties10);
            tableCellProperties10.Append(bottomLeftToTopRightBorderLineProperties10);
            tableCellProperties10.Append(solidFill50);

            tableCell10.Append(textBody19);
            tableCell10.Append(tableCellProperties10);

            A.TableCell tableCell11 = new A.TableCell();

            A.TextBody textBody20 = new A.TextBody();
            A.BodyProperties bodyProperties21 = new A.BodyProperties();
            A.ListStyle listStyle21 = new A.ListStyle();

            A.Paragraph paragraph24 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties16 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

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
            A.DefaultRunProperties defaultRunProperties31 = new A.DefaultRunProperties();

            paragraphProperties16.Append(lineSpacing11);
            paragraphProperties16.Append(spaceBefore9);
            paragraphProperties16.Append(spaceAfter9);
            paragraphProperties16.Append(bulletColorText9);
            paragraphProperties16.Append(bulletSizeText9);
            paragraphProperties16.Append(bulletFontText9);
            paragraphProperties16.Append(noBullet7);
            paragraphProperties16.Append(tabStopList8);
            paragraphProperties16.Append(defaultRunProperties31);

            A.EndParagraphRunProperties endParagraphRunProperties15 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline14 = new A.Outline();
            A.NoFill noFill71 = new A.NoFill();

            outline14.Append(noFill71);

            A.SolidFill solidFill51 = new A.SolidFill();
            A.PresetColor presetColor6 = new A.PresetColor() { Val = A.PresetColorValues.Black };

            solidFill51.Append(presetColor6);
            A.EffectList effectList16 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText7 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText7 = new A.UnderlineFillText();
            A.LatinFont latinFont38 = new A.LatinFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont33 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont38 = new A.ComplexScriptFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties15.Append(outline14);
            endParagraphRunProperties15.Append(solidFill51);
            endParagraphRunProperties15.Append(effectList16);
            endParagraphRunProperties15.Append(underlineFollowsText7);
            endParagraphRunProperties15.Append(underlineFillText7);
            endParagraphRunProperties15.Append(latinFont38);
            endParagraphRunProperties15.Append(eastAsianFont33);
            endParagraphRunProperties15.Append(complexScriptFont38);

            paragraph24.Append(paragraphProperties16);
            paragraph24.Append(endParagraphRunProperties15);

            textBody20.Append(bodyProperties21);
            textBody20.Append(listStyle21);
            textBody20.Append(paragraph24);

            A.TableCellProperties tableCellProperties11 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties11 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill72 = new A.NoFill();

            leftBorderLineProperties11.Append(noFill72);

            A.RightBorderLineProperties rightBorderLineProperties11 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill73 = new A.NoFill();

            rightBorderLineProperties11.Append(noFill73);

            A.TopBorderLineProperties topBorderLineProperties11 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill74 = new A.NoFill();
            A.PresetDash presetDash41 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round18 = new A.Round();
            A.HeadEnd headEnd18 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd18 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties11.Append(noFill74);
            topBorderLineProperties11.Append(presetDash41);
            topBorderLineProperties11.Append(round18);
            topBorderLineProperties11.Append(headEnd18);
            topBorderLineProperties11.Append(tailEnd18);

            A.BottomBorderLineProperties bottomBorderLineProperties11 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill75 = new A.NoFill();
            A.PresetDash presetDash42 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round19 = new A.Round();
            A.HeadEnd headEnd19 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd19 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties11.Append(noFill75);
            bottomBorderLineProperties11.Append(presetDash42);
            bottomBorderLineProperties11.Append(round19);
            bottomBorderLineProperties11.Append(headEnd19);
            bottomBorderLineProperties11.Append(tailEnd19);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties11 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill76 = new A.NoFill();
            A.PresetDash presetDash43 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties11.Append(noFill76);
            topLeftToBottomRightBorderLineProperties11.Append(presetDash43);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties11 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill77 = new A.NoFill();
            A.PresetDash presetDash44 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties11.Append(noFill77);
            bottomLeftToTopRightBorderLineProperties11.Append(presetDash44);

            A.SolidFill solidFill52 = new A.SolidFill();
            A.SchemeColor schemeColor54 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill52.Append(schemeColor54);

            tableCellProperties11.Append(leftBorderLineProperties11);
            tableCellProperties11.Append(rightBorderLineProperties11);
            tableCellProperties11.Append(topBorderLineProperties11);
            tableCellProperties11.Append(bottomBorderLineProperties11);
            tableCellProperties11.Append(topLeftToBottomRightBorderLineProperties11);
            tableCellProperties11.Append(bottomLeftToTopRightBorderLineProperties11);
            tableCellProperties11.Append(solidFill52);

            tableCell11.Append(textBody20);
            tableCell11.Append(tableCellProperties11);

            A.TableCell tableCell12 = new A.TableCell();

            A.TextBody textBody21 = new A.TextBody();
            A.BodyProperties bodyProperties22 = new A.BodyProperties();
            A.ListStyle listStyle22 = new A.ListStyle();

            A.Paragraph paragraph25 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties17 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

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
            A.DefaultRunProperties defaultRunProperties32 = new A.DefaultRunProperties();

            paragraphProperties17.Append(lineSpacing12);
            paragraphProperties17.Append(spaceBefore10);
            paragraphProperties17.Append(spaceAfter10);
            paragraphProperties17.Append(bulletColorText10);
            paragraphProperties17.Append(bulletSizeText10);
            paragraphProperties17.Append(bulletFontText10);
            paragraphProperties17.Append(noBullet8);
            paragraphProperties17.Append(tabStopList9);
            paragraphProperties17.Append(defaultRunProperties32);

            A.EndParagraphRunProperties endParagraphRunProperties16 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline15 = new A.Outline();
            A.NoFill noFill78 = new A.NoFill();

            outline15.Append(noFill78);

            A.SolidFill solidFill53 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex19 = new A.RgbColorModelHex() { Val = "5EC65A" };

            solidFill53.Append(rgbColorModelHex19);
            A.EffectList effectList17 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText8 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText8 = new A.UnderlineFillText();
            A.LatinFont latinFont39 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont34 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont39 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties16.Append(outline15);
            endParagraphRunProperties16.Append(solidFill53);
            endParagraphRunProperties16.Append(effectList17);
            endParagraphRunProperties16.Append(underlineFollowsText8);
            endParagraphRunProperties16.Append(underlineFillText8);
            endParagraphRunProperties16.Append(latinFont39);
            endParagraphRunProperties16.Append(eastAsianFont34);
            endParagraphRunProperties16.Append(complexScriptFont39);

            paragraph25.Append(paragraphProperties17);
            paragraph25.Append(endParagraphRunProperties16);

            textBody21.Append(bodyProperties22);
            textBody21.Append(listStyle22);
            textBody21.Append(paragraph25);

            A.TableCellProperties tableCellProperties12 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties12 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill79 = new A.NoFill();

            leftBorderLineProperties12.Append(noFill79);

            A.RightBorderLineProperties rightBorderLineProperties12 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill80 = new A.NoFill();

            rightBorderLineProperties12.Append(noFill80);

            A.TopBorderLineProperties topBorderLineProperties12 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill81 = new A.NoFill();
            A.PresetDash presetDash45 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round20 = new A.Round();
            A.HeadEnd headEnd20 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd20 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties12.Append(noFill81);
            topBorderLineProperties12.Append(presetDash45);
            topBorderLineProperties12.Append(round20);
            topBorderLineProperties12.Append(headEnd20);
            topBorderLineProperties12.Append(tailEnd20);

            A.BottomBorderLineProperties bottomBorderLineProperties12 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill82 = new A.NoFill();
            A.PresetDash presetDash46 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round21 = new A.Round();
            A.HeadEnd headEnd21 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd21 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties12.Append(noFill82);
            bottomBorderLineProperties12.Append(presetDash46);
            bottomBorderLineProperties12.Append(round21);
            bottomBorderLineProperties12.Append(headEnd21);
            bottomBorderLineProperties12.Append(tailEnd21);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties12 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill83 = new A.NoFill();
            A.PresetDash presetDash47 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties12.Append(noFill83);
            topLeftToBottomRightBorderLineProperties12.Append(presetDash47);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties12 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill84 = new A.NoFill();
            A.PresetDash presetDash48 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties12.Append(noFill84);
            bottomLeftToTopRightBorderLineProperties12.Append(presetDash48);

            A.SolidFill solidFill54 = new A.SolidFill();
            A.SchemeColor schemeColor55 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill54.Append(schemeColor55);

            tableCellProperties12.Append(leftBorderLineProperties12);
            tableCellProperties12.Append(rightBorderLineProperties12);
            tableCellProperties12.Append(topBorderLineProperties12);
            tableCellProperties12.Append(bottomBorderLineProperties12);
            tableCellProperties12.Append(topLeftToBottomRightBorderLineProperties12);
            tableCellProperties12.Append(bottomLeftToTopRightBorderLineProperties12);
            tableCellProperties12.Append(solidFill54);

            tableCell12.Append(textBody21);
            tableCell12.Append(tableCellProperties12);

            A.ExtensionList extensionList7 = new A.ExtensionList();

            A.Extension extension7 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement13 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"3020163789\" />");

            // extension7.Append(openXmlUnknownElement13);

            extensionList7.Append(extension7);

            tableRow4.Append(tableCell10);
            tableRow4.Append(tableCell11);
            tableRow4.Append(tableCell12);
            tableRow4.Append(extensionList7);

            A.TableRow tableRow5 = new A.TableRow() { Height = 747144L };

            A.TableCell tableCell13 = new A.TableCell();

            A.TextBody textBody22 = new A.TextBody();
            A.BodyProperties bodyProperties23 = new A.BodyProperties();
            A.ListStyle listStyle23 = new A.ListStyle();

            A.Paragraph paragraph26 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties18 = new A.ParagraphProperties() { LeftMargin = 171450, Indent = -171450 };
            A.BulletFontText bulletFontText11 = new A.BulletFontText();

            A.PictureBullet pictureBullet3 = new A.PictureBullet();

            A.Blip blip5 = new A.Blip() { Embed = "rId5" };

            A.BlipExtensionList blipExtensionList3 = new A.BlipExtensionList();

            A.BlipExtension blipExtension3 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement14 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId6\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension3.Append(openXmlUnknownElement14);

            blipExtensionList3.Append(blipExtension3);

            blip5.Append(blipExtensionList3);

            pictureBullet3.Append(blip5);

            paragraphProperties18.Append(bulletFontText11);
            paragraphProperties18.Append(pictureBullet3);

            A.Run run17 = new A.Run();

            A.RunProperties runProperties19 = new A.RunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Dirty = false };
            A.LatinFont latinFont40 = new A.LatinFont() { Typeface = "Segoe UI" };
            A.ComplexScriptFont complexScriptFont40 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            runProperties19.Append(latinFont40);
            runProperties19.Append(complexScriptFont40);
            A.Text text19 = new A.Text();

            // ********************************* Custom Code Start *********************************
            if (goalsArrayForSlide.Count() > 3)
                text19.Text = goalsArrayForSlide.ElementAt(3).goalName.ToString() ?? "";
            else
                text19.Text = "NO DATA";
            // ********************************* Custom Code End *********************************

            run17.Append(runProperties19);
            run17.Append(text19);

            paragraph26.Append(paragraphProperties18);
            paragraph26.Append(run17);

            textBody22.Append(bodyProperties23);
            textBody22.Append(listStyle23);
            textBody22.Append(paragraph26);

            A.TableCellProperties tableCellProperties13 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties13 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill85 = new A.NoFill();

            leftBorderLineProperties13.Append(noFill85);

            A.RightBorderLineProperties rightBorderLineProperties13 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill86 = new A.NoFill();

            rightBorderLineProperties13.Append(noFill86);

            A.TopBorderLineProperties topBorderLineProperties13 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill87 = new A.NoFill();
            A.PresetDash presetDash49 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round22 = new A.Round();
            A.HeadEnd headEnd22 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd22 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties13.Append(noFill87);
            topBorderLineProperties13.Append(presetDash49);
            topBorderLineProperties13.Append(round22);
            topBorderLineProperties13.Append(headEnd22);
            topBorderLineProperties13.Append(tailEnd22);

            A.BottomBorderLineProperties bottomBorderLineProperties13 = new A.BottomBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill88 = new A.NoFill();
            A.PresetDash presetDash50 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round23 = new A.Round();
            A.HeadEnd headEnd23 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd23 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties13.Append(noFill88);
            bottomBorderLineProperties13.Append(presetDash50);
            bottomBorderLineProperties13.Append(round23);
            bottomBorderLineProperties13.Append(headEnd23);
            bottomBorderLineProperties13.Append(tailEnd23);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties13 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill89 = new A.NoFill();
            A.PresetDash presetDash51 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties13.Append(noFill89);
            topLeftToBottomRightBorderLineProperties13.Append(presetDash51);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties13 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill90 = new A.NoFill();
            A.PresetDash presetDash52 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties13.Append(noFill90);
            bottomLeftToTopRightBorderLineProperties13.Append(presetDash52);

            A.SolidFill solidFill55 = new A.SolidFill();
            A.SchemeColor schemeColor56 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill55.Append(schemeColor56);

            tableCellProperties13.Append(leftBorderLineProperties13);
            tableCellProperties13.Append(rightBorderLineProperties13);
            tableCellProperties13.Append(topBorderLineProperties13);
            tableCellProperties13.Append(bottomBorderLineProperties13);
            tableCellProperties13.Append(topLeftToBottomRightBorderLineProperties13);
            tableCellProperties13.Append(bottomLeftToTopRightBorderLineProperties13);
            tableCellProperties13.Append(solidFill55);

            tableCell13.Append(textBody22);
            tableCell13.Append(tableCellProperties13);

            A.TableCell tableCell14 = new A.TableCell();

            A.TextBody textBody23 = new A.TextBody();
            A.BodyProperties bodyProperties24 = new A.BodyProperties();
            A.ListStyle listStyle24 = new A.ListStyle();

            A.Paragraph paragraph27 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties19 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

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
            A.DefaultRunProperties defaultRunProperties33 = new A.DefaultRunProperties();

            paragraphProperties19.Append(lineSpacing13);
            paragraphProperties19.Append(spaceBefore11);
            paragraphProperties19.Append(spaceAfter11);
            paragraphProperties19.Append(bulletColorText11);
            paragraphProperties19.Append(bulletSizeText11);
            paragraphProperties19.Append(bulletFontText12);
            paragraphProperties19.Append(noBullet9);
            paragraphProperties19.Append(tabStopList10);
            paragraphProperties19.Append(defaultRunProperties33);

            A.EndParagraphRunProperties endParagraphRunProperties17 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline16 = new A.Outline();
            A.NoFill noFill91 = new A.NoFill();

            outline16.Append(noFill91);

            A.SolidFill solidFill56 = new A.SolidFill();
            A.PresetColor presetColor7 = new A.PresetColor() { Val = A.PresetColorValues.Black };

            solidFill56.Append(presetColor7);
            A.EffectList effectList18 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText9 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText9 = new A.UnderlineFillText();
            A.LatinFont latinFont41 = new A.LatinFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont35 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont41 = new A.ComplexScriptFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties17.Append(outline16);
            endParagraphRunProperties17.Append(solidFill56);
            endParagraphRunProperties17.Append(effectList18);
            endParagraphRunProperties17.Append(underlineFollowsText9);
            endParagraphRunProperties17.Append(underlineFillText9);
            endParagraphRunProperties17.Append(latinFont41);
            endParagraphRunProperties17.Append(eastAsianFont35);
            endParagraphRunProperties17.Append(complexScriptFont41);

            paragraph27.Append(paragraphProperties19);
            paragraph27.Append(endParagraphRunProperties17);

            textBody23.Append(bodyProperties24);
            textBody23.Append(listStyle24);
            textBody23.Append(paragraph27);

            A.TableCellProperties tableCellProperties14 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties14 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill92 = new A.NoFill();

            leftBorderLineProperties14.Append(noFill92);

            A.RightBorderLineProperties rightBorderLineProperties14 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill93 = new A.NoFill();

            rightBorderLineProperties14.Append(noFill93);

            A.TopBorderLineProperties topBorderLineProperties14 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill94 = new A.NoFill();
            A.PresetDash presetDash53 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round24 = new A.Round();
            A.HeadEnd headEnd24 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd24 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties14.Append(noFill94);
            topBorderLineProperties14.Append(presetDash53);
            topBorderLineProperties14.Append(round24);
            topBorderLineProperties14.Append(headEnd24);
            topBorderLineProperties14.Append(tailEnd24);

            A.BottomBorderLineProperties bottomBorderLineProperties14 = new A.BottomBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill95 = new A.NoFill();
            A.PresetDash presetDash54 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round25 = new A.Round();
            A.HeadEnd headEnd25 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd25 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties14.Append(noFill95);
            bottomBorderLineProperties14.Append(presetDash54);
            bottomBorderLineProperties14.Append(round25);
            bottomBorderLineProperties14.Append(headEnd25);
            bottomBorderLineProperties14.Append(tailEnd25);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties14 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill96 = new A.NoFill();
            A.PresetDash presetDash55 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties14.Append(noFill96);
            topLeftToBottomRightBorderLineProperties14.Append(presetDash55);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties14 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill97 = new A.NoFill();
            A.PresetDash presetDash56 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties14.Append(noFill97);
            bottomLeftToTopRightBorderLineProperties14.Append(presetDash56);

            A.SolidFill solidFill57 = new A.SolidFill();
            A.SchemeColor schemeColor57 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill57.Append(schemeColor57);

            tableCellProperties14.Append(leftBorderLineProperties14);
            tableCellProperties14.Append(rightBorderLineProperties14);
            tableCellProperties14.Append(topBorderLineProperties14);
            tableCellProperties14.Append(bottomBorderLineProperties14);
            tableCellProperties14.Append(topLeftToBottomRightBorderLineProperties14);
            tableCellProperties14.Append(bottomLeftToTopRightBorderLineProperties14);
            tableCellProperties14.Append(solidFill57);

            tableCell14.Append(textBody23);
            tableCell14.Append(tableCellProperties14);

            A.TableCell tableCell15 = new A.TableCell();

            A.TextBody textBody24 = new A.TextBody();
            A.BodyProperties bodyProperties25 = new A.BodyProperties();
            A.ListStyle listStyle25 = new A.ListStyle();

            A.Paragraph paragraph28 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties20 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

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
            A.DefaultRunProperties defaultRunProperties34 = new A.DefaultRunProperties();

            paragraphProperties20.Append(lineSpacing14);
            paragraphProperties20.Append(spaceBefore12);
            paragraphProperties20.Append(spaceAfter12);
            paragraphProperties20.Append(bulletColorText12);
            paragraphProperties20.Append(bulletSizeText12);
            paragraphProperties20.Append(bulletFontText13);
            paragraphProperties20.Append(noBullet10);
            paragraphProperties20.Append(tabStopList11);
            paragraphProperties20.Append(defaultRunProperties34);

            A.EndParagraphRunProperties endParagraphRunProperties18 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline17 = new A.Outline();
            A.NoFill noFill98 = new A.NoFill();

            outline17.Append(noFill98);

            A.SolidFill solidFill58 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex20 = new A.RgbColorModelHex() { Val = "5EC65A" };

            solidFill58.Append(rgbColorModelHex20);
            A.EffectList effectList19 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText10 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText10 = new A.UnderlineFillText();
            A.LatinFont latinFont42 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont36 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont42 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties18.Append(outline17);
            endParagraphRunProperties18.Append(solidFill58);
            endParagraphRunProperties18.Append(effectList19);
            endParagraphRunProperties18.Append(underlineFollowsText10);
            endParagraphRunProperties18.Append(underlineFillText10);
            endParagraphRunProperties18.Append(latinFont42);
            endParagraphRunProperties18.Append(eastAsianFont36);
            endParagraphRunProperties18.Append(complexScriptFont42);

            paragraph28.Append(paragraphProperties20);
            paragraph28.Append(endParagraphRunProperties18);

            textBody24.Append(bodyProperties25);
            textBody24.Append(listStyle25);
            textBody24.Append(paragraph28);

            A.TableCellProperties tableCellProperties15 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties15 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill99 = new A.NoFill();

            leftBorderLineProperties15.Append(noFill99);

            A.RightBorderLineProperties rightBorderLineProperties15 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill100 = new A.NoFill();

            rightBorderLineProperties15.Append(noFill100);

            A.TopBorderLineProperties topBorderLineProperties15 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill101 = new A.NoFill();
            A.PresetDash presetDash57 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round26 = new A.Round();
            A.HeadEnd headEnd26 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd26 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties15.Append(noFill101);
            topBorderLineProperties15.Append(presetDash57);
            topBorderLineProperties15.Append(round26);
            topBorderLineProperties15.Append(headEnd26);
            topBorderLineProperties15.Append(tailEnd26);

            A.BottomBorderLineProperties bottomBorderLineProperties15 = new A.BottomBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill102 = new A.NoFill();
            A.PresetDash presetDash58 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round27 = new A.Round();
            A.HeadEnd headEnd27 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd27 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties15.Append(noFill102);
            bottomBorderLineProperties15.Append(presetDash58);
            bottomBorderLineProperties15.Append(round27);
            bottomBorderLineProperties15.Append(headEnd27);
            bottomBorderLineProperties15.Append(tailEnd27);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties15 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill103 = new A.NoFill();
            A.PresetDash presetDash59 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties15.Append(noFill103);
            topLeftToBottomRightBorderLineProperties15.Append(presetDash59);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties15 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill104 = new A.NoFill();
            A.PresetDash presetDash60 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties15.Append(noFill104);
            bottomLeftToTopRightBorderLineProperties15.Append(presetDash60);

            A.SolidFill solidFill59 = new A.SolidFill();
            A.SchemeColor schemeColor58 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill59.Append(schemeColor58);

            tableCellProperties15.Append(leftBorderLineProperties15);
            tableCellProperties15.Append(rightBorderLineProperties15);
            tableCellProperties15.Append(topBorderLineProperties15);
            tableCellProperties15.Append(bottomBorderLineProperties15);
            tableCellProperties15.Append(topLeftToBottomRightBorderLineProperties15);
            tableCellProperties15.Append(bottomLeftToTopRightBorderLineProperties15);
            tableCellProperties15.Append(solidFill59);

            tableCell15.Append(textBody24);
            tableCell15.Append(tableCellProperties15);

            A.ExtensionList extensionList8 = new A.ExtensionList();

            A.Extension extension8 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement15 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"3973020123\" />");

            // extension8.Append(openXmlUnknownElement15);

            extensionList8.Append(extension8);

            tableRow5.Append(tableCell13);
            tableRow5.Append(tableCell14);
            tableRow5.Append(tableCell15);
            tableRow5.Append(extensionList8);

            A.TableRow tableRow6 = new A.TableRow() { Height = 777502L };

            A.TableCell tableCell16 = new A.TableCell();

            A.TextBody textBody25 = new A.TextBody();
            A.BodyProperties bodyProperties26 = new A.BodyProperties();
            A.ListStyle listStyle26 = new A.ListStyle();

            A.Paragraph paragraph29 = new A.Paragraph();

            A.Run run18 = new A.Run();

            A.RunProperties runProperties20 = new A.RunProperties() { Language = "en-US", FontSize = 1300, Bold = true, Italic = false, Dirty = false };
            A.EffectList effectList20 = new A.EffectList();
            A.LatinFont latinFont43 = new A.LatinFont() { Typeface = "Segoe UI Semibold" };
            A.ComplexScriptFont complexScriptFont43 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold" };

            runProperties20.Append(effectList20);
            runProperties20.Append(latinFont43);
            runProperties20.Append(complexScriptFont43);
            A.Text text20 = new A.Text();

            // ********************************* Custom Code Start *********************************
            if (goalsArrayForSlide.Count() > 4)
                text20.Text = goalsArrayForSlide.ElementAt(4).goalName.ToString() ?? "";
            else
                text20.Text = "NO DATA";
            // ********************************* Custom Code End *********************************

            run18.Append(runProperties20);
            run18.Append(text20);

            A.EndParagraphRunProperties endParagraphRunProperties19 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1300, Bold = true, Italic = false, Dirty = false };

            A.SolidFill solidFill60 = new A.SolidFill();
            A.SchemeColor schemeColor59 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill60.Append(schemeColor59);
            A.LatinFont latinFont44 = new A.LatinFont() { Typeface = "Segoe UI Semibold" };
            A.ComplexScriptFont complexScriptFont44 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold" };

            endParagraphRunProperties19.Append(solidFill60);
            endParagraphRunProperties19.Append(latinFont44);
            endParagraphRunProperties19.Append(complexScriptFont44);

            paragraph29.Append(run18);
            paragraph29.Append(endParagraphRunProperties19);

            textBody25.Append(bodyProperties26);
            textBody25.Append(listStyle26);
            textBody25.Append(paragraph29);

            A.TableCellProperties tableCellProperties16 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties16 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill105 = new A.NoFill();

            leftBorderLineProperties16.Append(noFill105);

            A.RightBorderLineProperties rightBorderLineProperties16 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill106 = new A.NoFill();

            rightBorderLineProperties16.Append(noFill106);

            A.TopBorderLineProperties topBorderLineProperties16 = new A.TopBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill107 = new A.NoFill();
            A.PresetDash presetDash61 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round28 = new A.Round();
            A.HeadEnd headEnd28 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd28 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties16.Append(noFill107);
            topBorderLineProperties16.Append(presetDash61);
            topBorderLineProperties16.Append(round28);
            topBorderLineProperties16.Append(headEnd28);
            topBorderLineProperties16.Append(tailEnd28);

            A.BottomBorderLineProperties bottomBorderLineProperties16 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill108 = new A.NoFill();
            A.PresetDash presetDash62 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round29 = new A.Round();
            A.HeadEnd headEnd29 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd29 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties16.Append(noFill108);
            bottomBorderLineProperties16.Append(presetDash62);
            bottomBorderLineProperties16.Append(round29);
            bottomBorderLineProperties16.Append(headEnd29);
            bottomBorderLineProperties16.Append(tailEnd29);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties16 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill109 = new A.NoFill();
            A.PresetDash presetDash63 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties16.Append(noFill109);
            topLeftToBottomRightBorderLineProperties16.Append(presetDash63);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties16 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill110 = new A.NoFill();
            A.PresetDash presetDash64 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties16.Append(noFill110);
            bottomLeftToTopRightBorderLineProperties16.Append(presetDash64);

            A.SolidFill solidFill61 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex21 = new A.RgbColorModelHex() { Val = "F5F5F5" };

            solidFill61.Append(rgbColorModelHex21);

            tableCellProperties16.Append(leftBorderLineProperties16);
            tableCellProperties16.Append(rightBorderLineProperties16);
            tableCellProperties16.Append(topBorderLineProperties16);
            tableCellProperties16.Append(bottomBorderLineProperties16);
            tableCellProperties16.Append(topLeftToBottomRightBorderLineProperties16);
            tableCellProperties16.Append(bottomLeftToTopRightBorderLineProperties16);
            tableCellProperties16.Append(solidFill61);

            tableCell16.Append(textBody25);
            tableCell16.Append(tableCellProperties16);

            A.TableCell tableCell17 = new A.TableCell();

            A.TextBody textBody26 = new A.TextBody();
            A.BodyProperties bodyProperties27 = new A.BodyProperties();
            A.ListStyle listStyle27 = new A.ListStyle();

            A.Paragraph paragraph30 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties21 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

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
            A.DefaultRunProperties defaultRunProperties35 = new A.DefaultRunProperties();

            paragraphProperties21.Append(lineSpacing15);
            paragraphProperties21.Append(spaceBefore13);
            paragraphProperties21.Append(spaceAfter13);
            paragraphProperties21.Append(bulletColorText13);
            paragraphProperties21.Append(bulletSizeText13);
            paragraphProperties21.Append(bulletFontText14);
            paragraphProperties21.Append(noBullet11);
            paragraphProperties21.Append(tabStopList12);
            paragraphProperties21.Append(defaultRunProperties35);

            A.EndParagraphRunProperties endParagraphRunProperties20 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline18 = new A.Outline();
            A.NoFill noFill111 = new A.NoFill();

            outline18.Append(noFill111);

            A.SolidFill solidFill62 = new A.SolidFill();
            A.PresetColor presetColor8 = new A.PresetColor() { Val = A.PresetColorValues.Black };

            solidFill62.Append(presetColor8);
            A.EffectList effectList21 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText11 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText11 = new A.UnderlineFillText();
            A.LatinFont latinFont45 = new A.LatinFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont37 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont45 = new A.ComplexScriptFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties20.Append(outline18);
            endParagraphRunProperties20.Append(solidFill62);
            endParagraphRunProperties20.Append(effectList21);
            endParagraphRunProperties20.Append(underlineFollowsText11);
            endParagraphRunProperties20.Append(underlineFillText11);
            endParagraphRunProperties20.Append(latinFont45);
            endParagraphRunProperties20.Append(eastAsianFont37);
            endParagraphRunProperties20.Append(complexScriptFont45);

            paragraph30.Append(paragraphProperties21);
            paragraph30.Append(endParagraphRunProperties20);

            textBody26.Append(bodyProperties27);
            textBody26.Append(listStyle27);
            textBody26.Append(paragraph30);

            A.TableCellProperties tableCellProperties17 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties17 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill112 = new A.NoFill();

            leftBorderLineProperties17.Append(noFill112);

            A.RightBorderLineProperties rightBorderLineProperties17 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill113 = new A.NoFill();

            rightBorderLineProperties17.Append(noFill113);

            A.TopBorderLineProperties topBorderLineProperties17 = new A.TopBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill114 = new A.NoFill();
            A.PresetDash presetDash65 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round30 = new A.Round();
            A.HeadEnd headEnd30 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd30 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties17.Append(noFill114);
            topBorderLineProperties17.Append(presetDash65);
            topBorderLineProperties17.Append(round30);
            topBorderLineProperties17.Append(headEnd30);
            topBorderLineProperties17.Append(tailEnd30);

            A.BottomBorderLineProperties bottomBorderLineProperties17 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill115 = new A.NoFill();
            A.PresetDash presetDash66 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round31 = new A.Round();
            A.HeadEnd headEnd31 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd31 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties17.Append(noFill115);
            bottomBorderLineProperties17.Append(presetDash66);
            bottomBorderLineProperties17.Append(round31);
            bottomBorderLineProperties17.Append(headEnd31);
            bottomBorderLineProperties17.Append(tailEnd31);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties17 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill116 = new A.NoFill();
            A.PresetDash presetDash67 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties17.Append(noFill116);
            topLeftToBottomRightBorderLineProperties17.Append(presetDash67);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties17 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill117 = new A.NoFill();
            A.PresetDash presetDash68 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties17.Append(noFill117);
            bottomLeftToTopRightBorderLineProperties17.Append(presetDash68);

            A.SolidFill solidFill63 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex22 = new A.RgbColorModelHex() { Val = "F5F5F5" };

            solidFill63.Append(rgbColorModelHex22);

            tableCellProperties17.Append(leftBorderLineProperties17);
            tableCellProperties17.Append(rightBorderLineProperties17);
            tableCellProperties17.Append(topBorderLineProperties17);
            tableCellProperties17.Append(bottomBorderLineProperties17);
            tableCellProperties17.Append(topLeftToBottomRightBorderLineProperties17);
            tableCellProperties17.Append(bottomLeftToTopRightBorderLineProperties17);
            tableCellProperties17.Append(solidFill63);

            tableCell17.Append(textBody26);
            tableCell17.Append(tableCellProperties17);

            A.TableCell tableCell18 = new A.TableCell();

            A.TextBody textBody27 = new A.TextBody();
            A.BodyProperties bodyProperties28 = new A.BodyProperties();
            A.ListStyle listStyle28 = new A.ListStyle();

            A.Paragraph paragraph31 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties22 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

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
            A.DefaultRunProperties defaultRunProperties36 = new A.DefaultRunProperties();

            paragraphProperties22.Append(lineSpacing16);
            paragraphProperties22.Append(spaceBefore14);
            paragraphProperties22.Append(spaceAfter14);
            paragraphProperties22.Append(bulletColorText14);
            paragraphProperties22.Append(bulletSizeText14);
            paragraphProperties22.Append(bulletFontText15);
            paragraphProperties22.Append(noBullet12);
            paragraphProperties22.Append(tabStopList13);
            paragraphProperties22.Append(defaultRunProperties36);

            A.EndParagraphRunProperties endParagraphRunProperties21 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline19 = new A.Outline();
            A.NoFill noFill118 = new A.NoFill();

            outline19.Append(noFill118);

            A.SolidFill solidFill64 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex23 = new A.RgbColorModelHex() { Val = "C50F1F" };

            solidFill64.Append(rgbColorModelHex23);
            A.EffectList effectList22 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText12 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText12 = new A.UnderlineFillText();
            A.LatinFont latinFont46 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont38 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont46 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties21.Append(outline19);
            endParagraphRunProperties21.Append(solidFill64);
            endParagraphRunProperties21.Append(effectList22);
            endParagraphRunProperties21.Append(underlineFollowsText12);
            endParagraphRunProperties21.Append(underlineFillText12);
            endParagraphRunProperties21.Append(latinFont46);
            endParagraphRunProperties21.Append(eastAsianFont38);
            endParagraphRunProperties21.Append(complexScriptFont46);

            paragraph31.Append(paragraphProperties22);
            paragraph31.Append(endParagraphRunProperties21);

            textBody27.Append(bodyProperties28);
            textBody27.Append(listStyle28);
            textBody27.Append(paragraph31);

            A.TableCellProperties tableCellProperties18 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties18 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill119 = new A.NoFill();

            leftBorderLineProperties18.Append(noFill119);

            A.RightBorderLineProperties rightBorderLineProperties18 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill120 = new A.NoFill();

            rightBorderLineProperties18.Append(noFill120);

            A.TopBorderLineProperties topBorderLineProperties18 = new A.TopBorderLineProperties() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill121 = new A.NoFill();
            A.PresetDash presetDash69 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round32 = new A.Round();
            A.HeadEnd headEnd32 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd32 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties18.Append(noFill121);
            topBorderLineProperties18.Append(presetDash69);
            topBorderLineProperties18.Append(round32);
            topBorderLineProperties18.Append(headEnd32);
            topBorderLineProperties18.Append(tailEnd32);

            A.BottomBorderLineProperties bottomBorderLineProperties18 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill122 = new A.NoFill();
            A.PresetDash presetDash70 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round33 = new A.Round();
            A.HeadEnd headEnd33 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd33 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties18.Append(noFill122);
            bottomBorderLineProperties18.Append(presetDash70);
            bottomBorderLineProperties18.Append(round33);
            bottomBorderLineProperties18.Append(headEnd33);
            bottomBorderLineProperties18.Append(tailEnd33);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties18 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill123 = new A.NoFill();
            A.PresetDash presetDash71 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties18.Append(noFill123);
            topLeftToBottomRightBorderLineProperties18.Append(presetDash71);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties18 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill124 = new A.NoFill();
            A.PresetDash presetDash72 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties18.Append(noFill124);
            bottomLeftToTopRightBorderLineProperties18.Append(presetDash72);

            A.SolidFill solidFill65 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex24 = new A.RgbColorModelHex() { Val = "F5F5F5" };

            solidFill65.Append(rgbColorModelHex24);

            tableCellProperties18.Append(leftBorderLineProperties18);
            tableCellProperties18.Append(rightBorderLineProperties18);
            tableCellProperties18.Append(topBorderLineProperties18);
            tableCellProperties18.Append(bottomBorderLineProperties18);
            tableCellProperties18.Append(topLeftToBottomRightBorderLineProperties18);
            tableCellProperties18.Append(bottomLeftToTopRightBorderLineProperties18);
            tableCellProperties18.Append(solidFill65);

            tableCell18.Append(textBody27);
            tableCell18.Append(tableCellProperties18);

            A.ExtensionList extensionList9 = new A.ExtensionList();

            A.Extension extension9 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement16 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"3811565737\" />");

            // extension9.Append(openXmlUnknownElement16);

            extensionList9.Append(extension9);

            tableRow6.Append(tableCell16);
            tableRow6.Append(tableCell17);
            tableRow6.Append(tableCell18);
            tableRow6.Append(extensionList9);

            A.TableRow tableRow7 = new A.TableRow() { Height = 891827L };

            A.TableCell tableCell19 = new A.TableCell();

            A.TextBody textBody28 = new A.TextBody();
            A.BodyProperties bodyProperties29 = new A.BodyProperties();
            A.ListStyle listStyle29 = new A.ListStyle();

            A.Paragraph paragraph32 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties23 = new A.ParagraphProperties() { LeftMargin = 171450, Indent = -171450 };
            A.BulletFontText bulletFontText16 = new A.BulletFontText();

            A.PictureBullet pictureBullet4 = new A.PictureBullet();

            A.Blip blip6 = new A.Blip() { Embed = "rId5" };

            A.BlipExtensionList blipExtensionList4 = new A.BlipExtensionList();

            A.BlipExtension blipExtension4 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement17 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId6\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension4.Append(openXmlUnknownElement17);

            blipExtensionList4.Append(blipExtension4);

            blip6.Append(blipExtensionList4);

            pictureBullet4.Append(blip6);

            paragraphProperties23.Append(bulletFontText16);
            paragraphProperties23.Append(pictureBullet4);

            A.Run run19 = new A.Run();

            A.RunProperties runProperties21 = new A.RunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false };
            A.EffectList effectList23 = new A.EffectList();
            A.LatinFont latinFont47 = new A.LatinFont() { Typeface = "Segoe UI" };
            A.ComplexScriptFont complexScriptFont47 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            runProperties21.Append(effectList23);
            runProperties21.Append(latinFont47);
            runProperties21.Append(complexScriptFont47);
            A.Text text21 = new A.Text();

            // ********************************* Custom Code Start *********************************
            if (goalsArrayForSlide.Count() > 5)
                text21.Text = goalsArrayForSlide.ElementAt(5).goalName.ToString() ?? "";
            else
                text21.Text = "NO DATA";
            // ********************************* Custom Code End *********************************

            run19.Append(runProperties21);
            run19.Append(text21);

            A.EndParagraphRunProperties endParagraphRunProperties22 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Dirty = false };
            A.LatinFont latinFont48 = new A.LatinFont() { Typeface = "Segoe UI" };
            A.ComplexScriptFont complexScriptFont48 = new A.ComplexScriptFont() { Typeface = "Segoe UI" };

            endParagraphRunProperties22.Append(latinFont48);
            endParagraphRunProperties22.Append(complexScriptFont48);

            paragraph32.Append(paragraphProperties23);
            paragraph32.Append(run19);
            paragraph32.Append(endParagraphRunProperties22);

            textBody28.Append(bodyProperties29);
            textBody28.Append(listStyle29);
            textBody28.Append(paragraph32);

            A.TableCellProperties tableCellProperties19 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 365760, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties19 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill125 = new A.NoFill();

            leftBorderLineProperties19.Append(noFill125);

            A.RightBorderLineProperties rightBorderLineProperties19 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill126 = new A.NoFill();

            rightBorderLineProperties19.Append(noFill126);

            A.TopBorderLineProperties topBorderLineProperties19 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill127 = new A.NoFill();
            A.PresetDash presetDash73 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round34 = new A.Round();
            A.HeadEnd headEnd34 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd34 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties19.Append(noFill127);
            topBorderLineProperties19.Append(presetDash73);
            topBorderLineProperties19.Append(round34);
            topBorderLineProperties19.Append(headEnd34);
            topBorderLineProperties19.Append(tailEnd34);

            A.BottomBorderLineProperties bottomBorderLineProperties19 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill128 = new A.NoFill();
            A.PresetDash presetDash74 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round35 = new A.Round();
            A.HeadEnd headEnd35 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd35 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties19.Append(noFill128);
            bottomBorderLineProperties19.Append(presetDash74);
            bottomBorderLineProperties19.Append(round35);
            bottomBorderLineProperties19.Append(headEnd35);
            bottomBorderLineProperties19.Append(tailEnd35);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties19 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill129 = new A.NoFill();
            A.PresetDash presetDash75 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties19.Append(noFill129);
            topLeftToBottomRightBorderLineProperties19.Append(presetDash75);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties19 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill130 = new A.NoFill();
            A.PresetDash presetDash76 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties19.Append(noFill130);
            bottomLeftToTopRightBorderLineProperties19.Append(presetDash76);

            A.SolidFill solidFill66 = new A.SolidFill();
            A.SchemeColor schemeColor60 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill66.Append(schemeColor60);

            tableCellProperties19.Append(leftBorderLineProperties19);
            tableCellProperties19.Append(rightBorderLineProperties19);
            tableCellProperties19.Append(topBorderLineProperties19);
            tableCellProperties19.Append(bottomBorderLineProperties19);
            tableCellProperties19.Append(topLeftToBottomRightBorderLineProperties19);
            tableCellProperties19.Append(bottomLeftToTopRightBorderLineProperties19);
            tableCellProperties19.Append(solidFill66);

            tableCell19.Append(textBody28);
            tableCell19.Append(tableCellProperties19);

            A.TableCell tableCell20 = new A.TableCell();

            A.TextBody textBody29 = new A.TextBody();
            A.BodyProperties bodyProperties30 = new A.BodyProperties();
            A.ListStyle listStyle30 = new A.ListStyle();

            A.Paragraph paragraph33 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties24 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

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
            A.DefaultRunProperties defaultRunProperties37 = new A.DefaultRunProperties();

            paragraphProperties24.Append(lineSpacing17);
            paragraphProperties24.Append(spaceBefore15);
            paragraphProperties24.Append(spaceAfter15);
            paragraphProperties24.Append(bulletColorText15);
            paragraphProperties24.Append(bulletSizeText15);
            paragraphProperties24.Append(bulletFontText17);
            paragraphProperties24.Append(noBullet13);
            paragraphProperties24.Append(tabStopList14);
            paragraphProperties24.Append(defaultRunProperties37);

            A.EndParagraphRunProperties endParagraphRunProperties23 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = false, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false };

            A.Outline outline20 = new A.Outline();
            A.NoFill noFill131 = new A.NoFill();

            outline20.Append(noFill131);

            A.SolidFill solidFill67 = new A.SolidFill();
            A.PresetColor presetColor9 = new A.PresetColor() { Val = A.PresetColorValues.Black };

            solidFill67.Append(presetColor9);
            A.EffectList effectList24 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText13 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText13 = new A.UnderlineFillText();
            A.LatinFont latinFont49 = new A.LatinFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont39 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont49 = new A.ComplexScriptFont() { Typeface = "Segoe UI", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties23.Append(outline20);
            endParagraphRunProperties23.Append(solidFill67);
            endParagraphRunProperties23.Append(effectList24);
            endParagraphRunProperties23.Append(underlineFollowsText13);
            endParagraphRunProperties23.Append(underlineFillText13);
            endParagraphRunProperties23.Append(latinFont49);
            endParagraphRunProperties23.Append(eastAsianFont39);
            endParagraphRunProperties23.Append(complexScriptFont49);

            paragraph33.Append(paragraphProperties24);
            paragraph33.Append(endParagraphRunProperties23);

            textBody29.Append(bodyProperties30);
            textBody29.Append(listStyle30);
            textBody29.Append(paragraph33);

            A.TableCellProperties tableCellProperties20 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties20 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill132 = new A.NoFill();

            leftBorderLineProperties20.Append(noFill132);

            A.RightBorderLineProperties rightBorderLineProperties20 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill133 = new A.NoFill();

            rightBorderLineProperties20.Append(noFill133);

            A.TopBorderLineProperties topBorderLineProperties20 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill134 = new A.NoFill();
            A.PresetDash presetDash77 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round36 = new A.Round();
            A.HeadEnd headEnd36 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd36 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties20.Append(noFill134);
            topBorderLineProperties20.Append(presetDash77);
            topBorderLineProperties20.Append(round36);
            topBorderLineProperties20.Append(headEnd36);
            topBorderLineProperties20.Append(tailEnd36);

            A.BottomBorderLineProperties bottomBorderLineProperties20 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill135 = new A.NoFill();
            A.PresetDash presetDash78 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round37 = new A.Round();
            A.HeadEnd headEnd37 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd37 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties20.Append(noFill135);
            bottomBorderLineProperties20.Append(presetDash78);
            bottomBorderLineProperties20.Append(round37);
            bottomBorderLineProperties20.Append(headEnd37);
            bottomBorderLineProperties20.Append(tailEnd37);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties20 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill136 = new A.NoFill();
            A.PresetDash presetDash79 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties20.Append(noFill136);
            topLeftToBottomRightBorderLineProperties20.Append(presetDash79);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties20 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill137 = new A.NoFill();
            A.PresetDash presetDash80 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties20.Append(noFill137);
            bottomLeftToTopRightBorderLineProperties20.Append(presetDash80);

            A.SolidFill solidFill68 = new A.SolidFill();
            A.SchemeColor schemeColor61 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill68.Append(schemeColor61);

            tableCellProperties20.Append(leftBorderLineProperties20);
            tableCellProperties20.Append(rightBorderLineProperties20);
            tableCellProperties20.Append(topBorderLineProperties20);
            tableCellProperties20.Append(bottomBorderLineProperties20);
            tableCellProperties20.Append(topLeftToBottomRightBorderLineProperties20);
            tableCellProperties20.Append(bottomLeftToTopRightBorderLineProperties20);
            tableCellProperties20.Append(solidFill68);

            tableCell20.Append(textBody29);
            tableCell20.Append(tableCellProperties20);

            A.TableCell tableCell21 = new A.TableCell();

            A.TextBody textBody30 = new A.TextBody();
            A.BodyProperties bodyProperties31 = new A.BodyProperties();
            A.ListStyle listStyle31 = new A.ListStyle();

            A.Paragraph paragraph34 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties25 = new A.ParagraphProperties() { LeftMargin = 0, RightMargin = 0, Level = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, FontAlignment = A.TextFontAlignmentValues.Automatic, LatinLineBreak = false, Height = true };

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
            A.DefaultRunProperties defaultRunProperties38 = new A.DefaultRunProperties();

            paragraphProperties25.Append(lineSpacing18);
            paragraphProperties25.Append(spaceBefore16);
            paragraphProperties25.Append(spaceAfter16);
            paragraphProperties25.Append(bulletColorText16);
            paragraphProperties25.Append(bulletSizeText16);
            paragraphProperties25.Append(bulletFontText18);
            paragraphProperties25.Append(noBullet14);
            paragraphProperties25.Append(tabStopList15);
            paragraphProperties25.Append(defaultRunProperties38);

            A.EndParagraphRunProperties endParagraphRunProperties24 = new A.EndParagraphRunProperties() { Kumimoji = false, Language = "en-US", FontSize = 1200, Bold = true, Italic = false, Underline = A.TextUnderlineValues.None, Strike = A.TextStrikeValues.NoStrike, Kerning = 1200, Capital = A.TextCapsValues.None, Spacing = 0, NormalizeHeight = false, Baseline = 0, NoProof = false, Dirty = false };

            A.Outline outline21 = new A.Outline();
            A.NoFill noFill138 = new A.NoFill();

            outline21.Append(noFill138);

            A.SolidFill solidFill69 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex25 = new A.RgbColorModelHex() { Val = "5EC65A" };

            solidFill69.Append(rgbColorModelHex25);
            A.EffectList effectList25 = new A.EffectList();
            A.UnderlineFollowsText underlineFollowsText14 = new A.UnderlineFollowsText();
            A.UnderlineFillText underlineFillText14 = new A.UnderlineFillText();
            A.LatinFont latinFont50 = new A.LatinFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont40 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont50 = new A.ComplexScriptFont() { Typeface = "Segoe UI Semibold", Panose = "020B0502040204020203", PitchFamily = 34, CharacterSet = 0 };

            endParagraphRunProperties24.Append(outline21);
            endParagraphRunProperties24.Append(solidFill69);
            endParagraphRunProperties24.Append(effectList25);
            endParagraphRunProperties24.Append(underlineFollowsText14);
            endParagraphRunProperties24.Append(underlineFillText14);
            endParagraphRunProperties24.Append(latinFont50);
            endParagraphRunProperties24.Append(eastAsianFont40);
            endParagraphRunProperties24.Append(complexScriptFont50);

            paragraph34.Append(paragraphProperties25);
            paragraph34.Append(endParagraphRunProperties24);

            textBody30.Append(bodyProperties31);
            textBody30.Append(listStyle31);
            textBody30.Append(paragraph34);

            A.TableCellProperties tableCellProperties21 = new A.TableCellProperties() { LeftMargin = 162477, RightMargin = 162477, TopMargin = 81239, BottomMargin = 81239, Anchor = A.TextAnchoringTypeValues.Center };

            A.LeftBorderLineProperties leftBorderLineProperties21 = new A.LeftBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill139 = new A.NoFill();

            leftBorderLineProperties21.Append(noFill139);

            A.RightBorderLineProperties rightBorderLineProperties21 = new A.RightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill140 = new A.NoFill();

            rightBorderLineProperties21.Append(noFill140);

            A.TopBorderLineProperties topBorderLineProperties21 = new A.TopBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill141 = new A.NoFill();
            A.PresetDash presetDash81 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round38 = new A.Round();
            A.HeadEnd headEnd38 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd38 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            topBorderLineProperties21.Append(noFill141);
            topBorderLineProperties21.Append(presetDash81);
            topBorderLineProperties21.Append(round38);
            topBorderLineProperties21.Append(headEnd38);
            topBorderLineProperties21.Append(tailEnd38);

            A.BottomBorderLineProperties bottomBorderLineProperties21 = new A.BottomBorderLineProperties() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            A.NoFill noFill142 = new A.NoFill();
            A.PresetDash presetDash82 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Round round39 = new A.Round();
            A.HeadEnd headEnd39 = new A.HeadEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };
            A.TailEnd tailEnd39 = new A.TailEnd() { Type = A.LineEndValues.None, Width = A.LineEndWidthValues.Medium, Length = A.LineEndLengthValues.Medium };

            bottomBorderLineProperties21.Append(noFill142);
            bottomBorderLineProperties21.Append(presetDash82);
            bottomBorderLineProperties21.Append(round39);
            bottomBorderLineProperties21.Append(headEnd39);
            bottomBorderLineProperties21.Append(tailEnd39);

            A.TopLeftToBottomRightBorderLineProperties topLeftToBottomRightBorderLineProperties21 = new A.TopLeftToBottomRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill143 = new A.NoFill();
            A.PresetDash presetDash83 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            topLeftToBottomRightBorderLineProperties21.Append(noFill143);
            topLeftToBottomRightBorderLineProperties21.Append(presetDash83);

            A.BottomLeftToTopRightBorderLineProperties bottomLeftToTopRightBorderLineProperties21 = new A.BottomLeftToTopRightBorderLineProperties() { Width = 12700, CompoundLineType = A.CompoundLineValues.Single };
            A.NoFill noFill144 = new A.NoFill();
            A.PresetDash presetDash84 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            bottomLeftToTopRightBorderLineProperties21.Append(noFill144);
            bottomLeftToTopRightBorderLineProperties21.Append(presetDash84);

            A.SolidFill solidFill70 = new A.SolidFill();
            A.SchemeColor schemeColor62 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill70.Append(schemeColor62);

            tableCellProperties21.Append(leftBorderLineProperties21);
            tableCellProperties21.Append(rightBorderLineProperties21);
            tableCellProperties21.Append(topBorderLineProperties21);
            tableCellProperties21.Append(bottomBorderLineProperties21);
            tableCellProperties21.Append(topLeftToBottomRightBorderLineProperties21);
            tableCellProperties21.Append(bottomLeftToTopRightBorderLineProperties21);
            tableCellProperties21.Append(solidFill70);

            tableCell21.Append(textBody30);
            tableCell21.Append(tableCellProperties21);

            A.ExtensionList extensionList10 = new A.ExtensionList();

            A.Extension extension10 = new A.Extension() { Uri = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement18 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:rowId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" val=\"21430747\" />");

            // extension10.Append(openXmlUnknownElement18);

            extensionList10.Append(extension10);

            tableRow7.Append(tableCell19);
            tableRow7.Append(tableCell20);
            tableRow7.Append(tableCell21);
            tableRow7.Append(extensionList10);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);

            // ********************************* Custom Code Start *********************************

            A.TableRow[] tableRows = [tableRow2, tableRow3, tableRow4, tableRow5, tableRow6, tableRow7];

            for (int i = 0; i < goalsArrayForSlide.Count(); i++)
            {
                table1.Append(tableRows[i]);
            }

            // ********************************* Custom Code End *********************************

            graphicData1.Append(table1);

            graphic1.Append(graphicData1);

            graphicFrame1.Append(nonVisualGraphicFrameProperties1);
            graphicFrame1.Append(transform1);
            graphicFrame1.Append(graphic1);

            GroupShape groupShape1 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties3 = new NonVisualGroupShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties15 = new NonVisualDrawingProperties() { Id = (UInt32Value)11U, Name = "Group 10" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList5 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension5 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement19 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{95C71C4A-B6AD-141F-909D-38B2CFB8CE06}\" />");

            // nonVisualDrawingPropertiesExtension5.Append(openXmlUnknownElement19);

            nonVisualDrawingPropertiesExtensionList5.Append(nonVisualDrawingPropertiesExtension5);

            nonVisualDrawingProperties15.Append(nonVisualDrawingPropertiesExtensionList5);
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties3 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties15 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties3.Append(nonVisualDrawingProperties15);
            nonVisualGroupShapeProperties3.Append(nonVisualGroupShapeDrawingProperties3);
            nonVisualGroupShapeProperties3.Append(applicationNonVisualDrawingProperties15);

            GroupShapeProperties groupShapeProperties3 = new GroupShapeProperties();

            A.TransformGroup transformGroup3 = new A.TransformGroup();
            A.Offset offset15 = new A.Offset() { X = 8344061L, Y = 2525794L };
            A.Extents extents15 = new A.Extents() { Cx = 1669738L, Cy = 274285L };
            A.ChildOffset childOffset3 = new A.ChildOffset() { X = 8234323L, Y = 2640488L };
            A.ChildExtents childExtents3 = new A.ChildExtents() { Cx = 1669738L, Cy = 274285L };

            transformGroup3.Append(offset15);
            transformGroup3.Append(extents15);
            transformGroup3.Append(childOffset3);
            transformGroup3.Append(childExtents3);

            groupShapeProperties3.Append(transformGroup3);

            Picture picture3 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties3 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties16 = new NonVisualDrawingProperties() { Id = (UInt32Value)12U, Name = "Image 20", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList6 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension6 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement20 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{F2274644-3DA8-477E-516F-6CC1F6C1D36F}\" />");

            // nonVisualDrawingPropertiesExtension6.Append(openXmlUnknownElement20);

            nonVisualDrawingPropertiesExtensionList6.Append(nonVisualDrawingPropertiesExtension6);

            nonVisualDrawingProperties16.Append(nonVisualDrawingPropertiesExtensionList6);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties3 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks3 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties3.Append(pictureLocks3);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties16 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties3.Append(nonVisualDrawingProperties16);
            nonVisualPictureProperties3.Append(nonVisualPictureDrawingProperties3);
            nonVisualPictureProperties3.Append(applicationNonVisualDrawingProperties16);

            BlipFill blipFill3 = new BlipFill();
            A.Blip blip7 = new A.Blip() { Embed = "rId7" };

            A.Stretch stretch3 = new A.Stretch();
            A.FillRectangle fillRectangle3 = new A.FillRectangle();

            stretch3.Append(fillRectangle3);

            blipFill3.Append(blip7);
            blipFill3.Append(stretch3);

            ShapeProperties shapeProperties13 = new ShapeProperties();

            A.Transform2D transform2D12 = new A.Transform2D();
            A.Offset offset16 = new A.Offset() { X = 8234323L, Y = 2640488L };
            A.Extents extents16 = new A.Extents() { Cx = 1669738L, Cy = 274285L };

            transform2D12.Append(offset16);
            transform2D12.Append(extents16);

            A.PresetGeometry presetGeometry12 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList12 = new A.AdjustValueList();

            presetGeometry12.Append(adjustValueList12);

            A.SolidFill solidFill1000 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex103 = new A.RgbColorModelHex() { Val = "000000" };

            shapeProperties13.Append(transform2D12);
            shapeProperties13.Append(presetGeometry12);

            picture3.Append(nonVisualPictureProperties3);
            picture3.Append(blipFill3);
            picture3.Append(shapeProperties13);

            Picture picture4 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties4 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties17 = new NonVisualDrawingProperties() { Id = (UInt32Value)13U, Name = "Image 21", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList7 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension7 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement21 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{9CB0E207-5528-68A1-8A8A-A0596753E75E}\" />");

            // nonVisualDrawingPropertiesExtension7.Append(openXmlUnknownElement21);

            nonVisualDrawingPropertiesExtensionList7.Append(nonVisualDrawingPropertiesExtension7);

            nonVisualDrawingProperties17.Append(nonVisualDrawingPropertiesExtensionList7);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties4 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks4 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties4.Append(pictureLocks4);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties17 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties4.Append(nonVisualDrawingProperties17);
            nonVisualPictureProperties4.Append(nonVisualPictureDrawingProperties4);
            nonVisualPictureProperties4.Append(applicationNonVisualDrawingProperties17);

            BlipFill blipFill4 = new BlipFill();

            A.Blip blip8 = new A.Blip() { Embed = "rId8" };

            A.BlipExtensionList blipExtensionList5 = new A.BlipExtensionList();

            A.BlipExtension blipExtension5 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement22 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId9\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension5.Append(openXmlUnknownElement22);

            blipExtensionList5.Append(blipExtension5);

            blip8.Append(blipExtensionList5);

            A.Stretch stretch4 = new A.Stretch();
            A.FillRectangle fillRectangle4 = new A.FillRectangle();

            stretch4.Append(fillRectangle4);

            blipFill4.Append(blip8);
            blipFill4.Append(stretch4);

            ShapeProperties shapeProperties14 = new ShapeProperties();

            A.Transform2D transform2D13 = new A.Transform2D();
            A.Offset offset17 = new A.Offset() { X = 8257196L, Y = 2663344L };
            A.Extents extents17 = new A.Extents() { Cx = 228731L, Cy = 228571L };

            transform2D13.Append(offset17);
            transform2D13.Append(extents17);

            A.PresetGeometry presetGeometry13 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList13 = new A.AdjustValueList();

            presetGeometry13.Append(adjustValueList13);

            shapeProperties14.Append(transform2D13);
            shapeProperties14.Append(presetGeometry13);

            picture4.Append(nonVisualPictureProperties4);
            picture4.Append(blipFill4);
            picture4.Append(shapeProperties14);

            Shape shape10 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties10 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties18 = new NonVisualDrawingProperties() { Id = (UInt32Value)14U, Name = "Text 46" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList8 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension8 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement23 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{327398F6-EA13-F12E-89CC-697D5E66D280}\" />");

            // nonVisualDrawingPropertiesExtension8.Append(openXmlUnknownElement23);

            nonVisualDrawingPropertiesExtensionList8.Append(nonVisualDrawingPropertiesExtension8);

            nonVisualDrawingProperties18.Append(nonVisualDrawingPropertiesExtensionList8);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties10 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties18 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties10.Append(nonVisualDrawingProperties18);
            nonVisualShapeProperties10.Append(nonVisualShapeDrawingProperties10);
            nonVisualShapeProperties10.Append(applicationNonVisualDrawingProperties18);

            ShapeProperties shapeProperties15 = new ShapeProperties();

            A.Transform2D transform2D14 = new A.Transform2D();
            A.Offset offset18 = new A.Offset() { X = 8577419L, Y = 2651915L };
            A.Extents extents18 = new A.Extents() { Cx = 1280895L, Cy = 228571L };

            transform2D14.Append(offset18);
            transform2D14.Append(extents18);

            A.PresetGeometry presetGeometry14 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList14 = new A.AdjustValueList();

            presetGeometry14.Append(adjustValueList14);
            A.NoFill noFill145 = new A.NoFill();
            A.Outline outline22 = new A.Outline();

            shapeProperties15.Append(transform2D14);
            shapeProperties15.Append(presetGeometry14);
            shapeProperties15.Append(noFill145);
            shapeProperties15.Append(outline22);

            TextBody textBody31 = new TextBody();
            A.BodyProperties bodyProperties32 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle32 = new A.ListStyle();

            A.Paragraph paragraph35 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties26 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing19 = new A.LineSpacing();
            A.SpacingPercent spacingPercent19 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing19.Append(spacingPercent19);

            paragraphProperties26.Append(lineSpacing19);

            A.Run run20 = new A.Run();

            A.RunProperties runProperties22 = new A.RunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            A.SolidFill solidFill71 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex26 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha5 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex26.Append(alpha5);

            solidFill71.Append(rgbColorModelHex26);
            A.LatinFont latinFont51 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont41 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont51 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties22.Append(solidFill71);
            runProperties22.Append(latinFont51);
            runProperties22.Append(eastAsianFont41);
            runProperties22.Append(complexScriptFont51);
            A.Text text22 = new A.Text();

            // ********************************* Custom Code Start *********************************
            if (goalsArrayForSlide.Count > 1)

                text22.Text = goalsArrayForSlide.ElementAt(1).goalOwner.ToString() ?? "";
            else
                text22.Text = "NO DATA";
            // ********************************* Custom Code End ***********************************

            run20.Append(runProperties22);
            run20.Append(text22);
            A.EndParagraphRunProperties endParagraphRunProperties25 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            paragraph35.Append(paragraphProperties26);
            paragraph35.Append(run20);
            paragraph35.Append(endParagraphRunProperties25);

            textBody31.Append(bodyProperties32);
            textBody31.Append(listStyle32);
            textBody31.Append(paragraph35);

            shape10.Append(nonVisualShapeProperties10);
            shape10.Append(shapeProperties15);
            shape10.Append(textBody31);

            groupShape1.Append(nonVisualGroupShapeProperties3);
            groupShape1.Append(groupShapeProperties3);
            groupShape1.Append(picture3);
            groupShape1.Append(picture4);
            groupShape1.Append(shape10);

            // ************************************************* Owner Code Start *************************************************

            GroupShape groupShape2 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties4 = new NonVisualGroupShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties19 = new NonVisualDrawingProperties() { Id = (UInt32Value)15U, Name = "Group 14" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList9 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension9 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement24 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{28AD22F3-1239-938D-378E-75C0C45E56A9}\" />");

            // nonVisualDrawingPropertiesExtension9.Append(openXmlUnknownElement24);

            nonVisualDrawingPropertiesExtensionList9.Append(nonVisualDrawingPropertiesExtension9);

            nonVisualDrawingProperties19.Append(nonVisualDrawingPropertiesExtensionList9);
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties4 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties19 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties4.Append(nonVisualDrawingProperties19);
            nonVisualGroupShapeProperties4.Append(nonVisualGroupShapeDrawingProperties4);
            nonVisualGroupShapeProperties4.Append(applicationNonVisualDrawingProperties19);

            GroupShapeProperties groupShapeProperties4 = new GroupShapeProperties();

            A.TransformGroup transformGroup4 = new A.TransformGroup();
            A.Offset offset19 = new A.Offset() { X = 8344061L, Y = 3294606L };
            A.Extents extents19 = new A.Extents() { Cx = 1669738L, Cy = 274285L };
            A.ChildOffset childOffset4 = new A.ChildOffset() { X = 8234323L, Y = 2640488L };
            A.ChildExtents childExtents4 = new A.ChildExtents() { Cx = 1669738L, Cy = 274285L };

            transformGroup4.Append(offset19);
            transformGroup4.Append(extents19);
            transformGroup4.Append(childOffset4);
            transformGroup4.Append(childExtents4);

            groupShapeProperties4.Append(transformGroup4);

            Picture picture5 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties5 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties20 = new NonVisualDrawingProperties() { Id = (UInt32Value)16U, Name = "Image 20", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList10 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension10 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement25 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{90F76FED-8261-4D46-7AC7-F0217CBE186C}\" />");

            // nonVisualDrawingPropertiesExtension10.Append(openXmlUnknownElement25);

            nonVisualDrawingPropertiesExtensionList10.Append(nonVisualDrawingPropertiesExtension10);

            nonVisualDrawingProperties20.Append(nonVisualDrawingPropertiesExtensionList10);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties5 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks5 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties5.Append(pictureLocks5);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties20 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties5.Append(nonVisualDrawingProperties20);
            nonVisualPictureProperties5.Append(nonVisualPictureDrawingProperties5);
            nonVisualPictureProperties5.Append(applicationNonVisualDrawingProperties20);

            BlipFill blipFill5 = new BlipFill();
            A.Blip blip9 = new A.Blip() { Embed = "rId7" };

            A.Stretch stretch5 = new A.Stretch();
            A.FillRectangle fillRectangle5 = new A.FillRectangle();

            stretch5.Append(fillRectangle5);

            blipFill5.Append(blip9);
            blipFill5.Append(stretch5);

            ShapeProperties shapeProperties16 = new ShapeProperties();

            A.Transform2D transform2D15 = new A.Transform2D();
            A.Offset offset20 = new A.Offset() { X = 8234323L, Y = 2640488L };
            A.Extents extents20 = new A.Extents() { Cx = 1669738L, Cy = 274285L };

            transform2D15.Append(offset20);
            transform2D15.Append(extents20);

            A.PresetGeometry presetGeometry15 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList15 = new A.AdjustValueList();

            presetGeometry15.Append(adjustValueList15);

            shapeProperties16.Append(transform2D15);
            shapeProperties16.Append(presetGeometry15);

            picture5.Append(nonVisualPictureProperties5);
            picture5.Append(blipFill5);
            picture5.Append(shapeProperties16);

            Picture picture6 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties6 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties21 = new NonVisualDrawingProperties() { Id = (UInt32Value)17U, Name = "Image 21", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList11 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension11 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement26 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{E0932B29-60BF-CAA5-69B2-E98C4379B1C6}\" />");

            // nonVisualDrawingPropertiesExtension11.Append(openXmlUnknownElement26);

            nonVisualDrawingPropertiesExtensionList11.Append(nonVisualDrawingPropertiesExtension11);

            nonVisualDrawingProperties21.Append(nonVisualDrawingPropertiesExtensionList11);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties6 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks6 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties6.Append(pictureLocks6);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties21 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties6.Append(nonVisualDrawingProperties21);
            nonVisualPictureProperties6.Append(nonVisualPictureDrawingProperties6);
            nonVisualPictureProperties6.Append(applicationNonVisualDrawingProperties21);

            BlipFill blipFill6 = new BlipFill();

            A.Blip blip10 = new A.Blip() { Embed = "rId8" };

            A.BlipExtensionList blipExtensionList6 = new A.BlipExtensionList();

            A.BlipExtension blipExtension6 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement27 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId9\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension6.Append(openXmlUnknownElement27);

            blipExtensionList6.Append(blipExtension6);

            blip10.Append(blipExtensionList6);

            A.Stretch stretch6 = new A.Stretch();
            A.FillRectangle fillRectangle6 = new A.FillRectangle();

            stretch6.Append(fillRectangle6);

            blipFill6.Append(blip10);
            blipFill6.Append(stretch6);

            ShapeProperties shapeProperties17 = new ShapeProperties();

            A.Transform2D transform2D16 = new A.Transform2D();
            A.Offset offset21 = new A.Offset() { X = 8257196L, Y = 2663344L };
            A.Extents extents21 = new A.Extents() { Cx = 228731L, Cy = 228571L };

            transform2D16.Append(offset21);
            transform2D16.Append(extents21);

            A.PresetGeometry presetGeometry16 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList16 = new A.AdjustValueList();

            presetGeometry16.Append(adjustValueList16);

            shapeProperties17.Append(transform2D16);
            shapeProperties17.Append(presetGeometry16);

            picture6.Append(nonVisualPictureProperties6);
            picture6.Append(blipFill6);
            picture6.Append(shapeProperties17);

            Shape shape11 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties11 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties22 = new NonVisualDrawingProperties() { Id = (UInt32Value)18U, Name = "Text 46" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList12 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension12 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement28 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{2AE16487-3A6F-5CC6-5890-23B6EE77DD04}\" />");

            // nonVisualDrawingPropertiesExtension12.Append(openXmlUnknownElement28);

            nonVisualDrawingPropertiesExtensionList12.Append(nonVisualDrawingPropertiesExtension12);

            nonVisualDrawingProperties22.Append(nonVisualDrawingPropertiesExtensionList12);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties11 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties22 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties11.Append(nonVisualDrawingProperties22);
            nonVisualShapeProperties11.Append(nonVisualShapeDrawingProperties11);
            nonVisualShapeProperties11.Append(applicationNonVisualDrawingProperties22);

            ShapeProperties shapeProperties18 = new ShapeProperties();

            A.Transform2D transform2D17 = new A.Transform2D();
            A.Offset offset22 = new A.Offset() { X = 8577419L, Y = 2651915L };
            A.Extents extents22 = new A.Extents() { Cx = 1280895L, Cy = 228571L };

            transform2D17.Append(offset22);
            transform2D17.Append(extents22);

            A.PresetGeometry presetGeometry17 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList17 = new A.AdjustValueList();

            presetGeometry17.Append(adjustValueList17);
            A.NoFill noFill146 = new A.NoFill();
            A.Outline outline23 = new A.Outline();

            shapeProperties18.Append(transform2D17);
            shapeProperties18.Append(presetGeometry17);
            shapeProperties18.Append(noFill146);
            shapeProperties18.Append(outline23);

            TextBody textBody32 = new TextBody();
            A.BodyProperties bodyProperties33 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle33 = new A.ListStyle();

            A.Paragraph paragraph36 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties27 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing20 = new A.LineSpacing();
            A.SpacingPercent spacingPercent20 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing20.Append(spacingPercent20);

            paragraphProperties27.Append(lineSpacing20);

            A.Run run21 = new A.Run();

            A.RunProperties runProperties23 = new A.RunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            A.SolidFill solidFill72 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex27 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha6 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex27.Append(alpha6);

            solidFill72.Append(rgbColorModelHex27);
            A.LatinFont latinFont52 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont42 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont52 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties23.Append(solidFill72);
            runProperties23.Append(latinFont52);
            runProperties23.Append(eastAsianFont42);
            runProperties23.Append(complexScriptFont52);
            A.Text text23 = new A.Text();

            // ********************************* Custom Code Start *********************************
            if (goalsArrayForSlide.Count > 2)
                text23.Text = goalsArrayForSlide.ElementAt(2).goalOwner.ToString() ?? "";
            else
                text23.Text = "NO DATA";
            // ********************************* Custom Code End ***********************************

            run21.Append(runProperties23);
            run21.Append(text23);
            A.EndParagraphRunProperties endParagraphRunProperties26 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            paragraph36.Append(paragraphProperties27);
            paragraph36.Append(run21);
            paragraph36.Append(endParagraphRunProperties26);

            textBody32.Append(bodyProperties33);
            textBody32.Append(listStyle33);
            textBody32.Append(paragraph36);

            shape11.Append(nonVisualShapeProperties11);
            shape11.Append(shapeProperties18);
            shape11.Append(textBody32);

            groupShape2.Append(nonVisualGroupShapeProperties4);
            groupShape2.Append(groupShapeProperties4);
            groupShape2.Append(picture5);
            groupShape2.Append(picture6);
            groupShape2.Append(shape11);

            GroupShape groupShape3 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties5 = new NonVisualGroupShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties23 = new NonVisualDrawingProperties() { Id = (UInt32Value)19U, Name = "Group 18" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList13 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension13 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement29 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{90D6CC8D-5C42-1C42-00B8-DCC17D9C981A}\" />");

            // nonVisualDrawingPropertiesExtension13.Append(openXmlUnknownElement29);

            nonVisualDrawingPropertiesExtensionList13.Append(nonVisualDrawingPropertiesExtension13);

            nonVisualDrawingProperties23.Append(nonVisualDrawingPropertiesExtensionList13);
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties5 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties23 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties5.Append(nonVisualDrawingProperties23);
            nonVisualGroupShapeProperties5.Append(nonVisualGroupShapeDrawingProperties5);
            nonVisualGroupShapeProperties5.Append(applicationNonVisualDrawingProperties23);

            GroupShapeProperties groupShapeProperties5 = new GroupShapeProperties();

            A.TransformGroup transformGroup5 = new A.TransformGroup();
            A.Offset offset23 = new A.Offset() { X = 8345338L, Y = 4009404L };
            A.Extents extents23 = new A.Extents() { Cx = 1669738L, Cy = 274285L };
            A.ChildOffset childOffset5 = new A.ChildOffset() { X = 8234323L, Y = 2640488L };
            A.ChildExtents childExtents5 = new A.ChildExtents() { Cx = 1669738L, Cy = 274285L };

            transformGroup5.Append(offset23);
            transformGroup5.Append(extents23);
            transformGroup5.Append(childOffset5);
            transformGroup5.Append(childExtents5);

            groupShapeProperties5.Append(transformGroup5);

            Picture picture7 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties7 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties24 = new NonVisualDrawingProperties() { Id = (UInt32Value)20U, Name = "Image 20", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList14 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension14 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement30 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{12262BAA-EA05-6775-C85A-A4DDCDAA83F4}\" />");

            // nonVisualDrawingPropertiesExtension14.Append(openXmlUnknownElement30);

            nonVisualDrawingPropertiesExtensionList14.Append(nonVisualDrawingPropertiesExtension14);

            nonVisualDrawingProperties24.Append(nonVisualDrawingPropertiesExtensionList14);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties7 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks7 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties7.Append(pictureLocks7);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties24 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties7.Append(nonVisualDrawingProperties24);
            nonVisualPictureProperties7.Append(nonVisualPictureDrawingProperties7);
            nonVisualPictureProperties7.Append(applicationNonVisualDrawingProperties24);

            BlipFill blipFill7 = new BlipFill();
            A.Blip blip11 = new A.Blip() { Embed = "rId7" };

            A.Stretch stretch7 = new A.Stretch();
            A.FillRectangle fillRectangle7 = new A.FillRectangle();

            stretch7.Append(fillRectangle7);

            blipFill7.Append(blip11);
            blipFill7.Append(stretch7);

            ShapeProperties shapeProperties19 = new ShapeProperties();

            A.Transform2D transform2D18 = new A.Transform2D();
            A.Offset offset24 = new A.Offset() { X = 8234323L, Y = 2640488L };
            A.Extents extents24 = new A.Extents() { Cx = 1669738L, Cy = 274285L };

            transform2D18.Append(offset24);
            transform2D18.Append(extents24);

            A.PresetGeometry presetGeometry18 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList18 = new A.AdjustValueList();

            presetGeometry18.Append(adjustValueList18);

            shapeProperties19.Append(transform2D18);
            shapeProperties19.Append(presetGeometry18);

            picture7.Append(nonVisualPictureProperties7);
            picture7.Append(blipFill7);
            picture7.Append(shapeProperties19);

            Picture picture8 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties8 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties25 = new NonVisualDrawingProperties() { Id = (UInt32Value)21U, Name = "Image 21", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList15 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension15 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement31 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{669DFA83-4368-0988-521F-37A545B68EFC}\" />");

            // nonVisualDrawingPropertiesExtension15.Append(openXmlUnknownElement31);

            nonVisualDrawingPropertiesExtensionList15.Append(nonVisualDrawingPropertiesExtension15);

            nonVisualDrawingProperties25.Append(nonVisualDrawingPropertiesExtensionList15);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties8 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks8 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties8.Append(pictureLocks8);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties25 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties8.Append(nonVisualDrawingProperties25);
            nonVisualPictureProperties8.Append(nonVisualPictureDrawingProperties8);
            nonVisualPictureProperties8.Append(applicationNonVisualDrawingProperties25);

            BlipFill blipFill8 = new BlipFill();

            A.Blip blip12 = new A.Blip() { Embed = "rId8" };

            A.BlipExtensionList blipExtensionList7 = new A.BlipExtensionList();

            A.BlipExtension blipExtension7 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement32 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId9\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension7.Append(openXmlUnknownElement32);

            blipExtensionList7.Append(blipExtension7);

            blip12.Append(blipExtensionList7);

            A.Stretch stretch8 = new A.Stretch();
            A.FillRectangle fillRectangle8 = new A.FillRectangle();

            stretch8.Append(fillRectangle8);

            blipFill8.Append(blip12);
            blipFill8.Append(stretch8);

            ShapeProperties shapeProperties20 = new ShapeProperties();

            A.Transform2D transform2D19 = new A.Transform2D();
            A.Offset offset25 = new A.Offset() { X = 8257196L, Y = 2663344L };
            A.Extents extents25 = new A.Extents() { Cx = 228731L, Cy = 228571L };

            transform2D19.Append(offset25);
            transform2D19.Append(extents25);

            A.PresetGeometry presetGeometry19 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList19 = new A.AdjustValueList();

            presetGeometry19.Append(adjustValueList19);

            shapeProperties20.Append(transform2D19);
            shapeProperties20.Append(presetGeometry19);

            picture8.Append(nonVisualPictureProperties8);
            picture8.Append(blipFill8);
            picture8.Append(shapeProperties20);

            Shape shape12 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties12 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties26 = new NonVisualDrawingProperties() { Id = (UInt32Value)22U, Name = "Text 46" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList16 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension16 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement33 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{03A76668-1C58-779C-4ED3-E2549AC71F50}\" />");

            // nonVisualDrawingPropertiesExtension16.Append(openXmlUnknownElement33);

            nonVisualDrawingPropertiesExtensionList16.Append(nonVisualDrawingPropertiesExtension16);

            nonVisualDrawingProperties26.Append(nonVisualDrawingPropertiesExtensionList16);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties12 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties26 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties12.Append(nonVisualDrawingProperties26);
            nonVisualShapeProperties12.Append(nonVisualShapeDrawingProperties12);
            nonVisualShapeProperties12.Append(applicationNonVisualDrawingProperties26);

            ShapeProperties shapeProperties21 = new ShapeProperties();

            A.Transform2D transform2D20 = new A.Transform2D();
            A.Offset offset26 = new A.Offset() { X = 8577419L, Y = 2651915L };
            A.Extents extents26 = new A.Extents() { Cx = 1280895L, Cy = 228571L };

            transform2D20.Append(offset26);
            transform2D20.Append(extents26);

            A.PresetGeometry presetGeometry20 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList20 = new A.AdjustValueList();

            presetGeometry20.Append(adjustValueList20);
            A.NoFill noFill147 = new A.NoFill();
            A.Outline outline24 = new A.Outline();

            shapeProperties21.Append(transform2D20);
            shapeProperties21.Append(presetGeometry20);
            shapeProperties21.Append(noFill147);
            shapeProperties21.Append(outline24);

            // ************************************************* Owner Code End *************************************************

            TextBody textBody33 = new TextBody();
            A.BodyProperties bodyProperties34 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle34 = new A.ListStyle();

            A.Paragraph paragraph37 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties28 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing21 = new A.LineSpacing();
            A.SpacingPercent spacingPercent21 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing21.Append(spacingPercent21);

            paragraphProperties28.Append(lineSpacing21);

            A.Run run22 = new A.Run();

            A.RunProperties runProperties24 = new A.RunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            A.SolidFill solidFill73 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex28 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha7 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex28.Append(alpha7);

            solidFill73.Append(rgbColorModelHex28);
            A.LatinFont latinFont53 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont43 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont53 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties24.Append(solidFill73);
            runProperties24.Append(latinFont53);
            runProperties24.Append(eastAsianFont43);
            runProperties24.Append(complexScriptFont53);
            A.Text text24 = new A.Text();

            // ********************************* Custom Code Start *********************************
            if (goalsArrayForSlide.Count > 3)
                text24.Text = goalsArrayForSlide.ElementAt(3).goalOwner.ToString() ?? "";
            else
                text24.Text = "NO DATA";
            // ********************************* Custom Code End ***********************************

            run22.Append(runProperties24);
            run22.Append(text24);
            A.EndParagraphRunProperties endParagraphRunProperties27 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            paragraph37.Append(paragraphProperties28);
            paragraph37.Append(run22);
            paragraph37.Append(endParagraphRunProperties27);

            textBody33.Append(bodyProperties34);
            textBody33.Append(listStyle34);
            textBody33.Append(paragraph37);

            shape12.Append(nonVisualShapeProperties12);
            shape12.Append(shapeProperties21);
            shape12.Append(textBody33);

            groupShape3.Append(nonVisualGroupShapeProperties5);
            groupShape3.Append(groupShapeProperties5);
            groupShape3.Append(picture7);
            groupShape3.Append(picture8);
            groupShape3.Append(shape12);

            GroupShape groupShape4 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties6 = new NonVisualGroupShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties27 = new NonVisualDrawingProperties() { Id = (UInt32Value)23U, Name = "Group 22" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList17 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension17 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement34 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{CFBD9566-8447-13A6-9E30-1CA71DD236DB}\" />");

            // nonVisualDrawingPropertiesExtension17.Append(openXmlUnknownElement34);

            nonVisualDrawingPropertiesExtensionList17.Append(nonVisualDrawingPropertiesExtension17);

            nonVisualDrawingProperties27.Append(nonVisualDrawingPropertiesExtensionList17);
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties6 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties27 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties6.Append(nonVisualDrawingProperties27);
            nonVisualGroupShapeProperties6.Append(nonVisualGroupShapeDrawingProperties6);
            nonVisualGroupShapeProperties6.Append(applicationNonVisualDrawingProperties27);

            GroupShapeProperties groupShapeProperties6 = new GroupShapeProperties();

            A.TransformGroup transformGroup6 = new A.TransformGroup();
            A.Offset offset27 = new A.Offset() { X = 8357150L, Y = 5665490L };
            A.Extents extents27 = new A.Extents() { Cx = 1669738L, Cy = 274285L };
            A.ChildOffset childOffset6 = new A.ChildOffset() { X = 8234323L, Y = 2640488L };
            A.ChildExtents childExtents6 = new A.ChildExtents() { Cx = 1669738L, Cy = 274285L };

            transformGroup6.Append(offset27);
            transformGroup6.Append(extents27);
            transformGroup6.Append(childOffset6);
            transformGroup6.Append(childExtents6);

            groupShapeProperties6.Append(transformGroup6);

            Picture picture9 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties9 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties28 = new NonVisualDrawingProperties() { Id = (UInt32Value)24U, Name = "Image 20", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList18 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension18 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement35 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{7A123E97-308F-9173-6E4D-8B0C599055E3}\" />");

            // nonVisualDrawingPropertiesExtension18.Append(openXmlUnknownElement35);

            nonVisualDrawingPropertiesExtensionList18.Append(nonVisualDrawingPropertiesExtension18);

            nonVisualDrawingProperties28.Append(nonVisualDrawingPropertiesExtensionList18);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties9 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks9 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties9.Append(pictureLocks9);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties28 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties9.Append(nonVisualDrawingProperties28);
            nonVisualPictureProperties9.Append(nonVisualPictureDrawingProperties9);
            nonVisualPictureProperties9.Append(applicationNonVisualDrawingProperties28);

            BlipFill blipFill9 = new BlipFill();
            A.Blip blip13 = new A.Blip() { Embed = "rId7" };

            A.Stretch stretch9 = new A.Stretch();
            A.FillRectangle fillRectangle9 = new A.FillRectangle();

            stretch9.Append(fillRectangle9);

            blipFill9.Append(blip13);
            blipFill9.Append(stretch9);

            ShapeProperties shapeProperties22 = new ShapeProperties();

            A.Transform2D transform2D21 = new A.Transform2D();
            A.Offset offset28 = new A.Offset() { X = 8234323L, Y = 2640488L };
            A.Extents extents28 = new A.Extents() { Cx = 1669738L, Cy = 274285L };

            transform2D21.Append(offset28);
            transform2D21.Append(extents28);

            A.PresetGeometry presetGeometry21 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList21 = new A.AdjustValueList();

            presetGeometry21.Append(adjustValueList21);

            shapeProperties22.Append(transform2D21);
            shapeProperties22.Append(presetGeometry21);

            picture9.Append(nonVisualPictureProperties9);
            picture9.Append(blipFill9);
            picture9.Append(shapeProperties22);

            Picture picture10 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties10 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties29 = new NonVisualDrawingProperties() { Id = (UInt32Value)25U, Name = "Image 21", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList19 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension19 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement36 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{C721E648-785A-2B15-4551-E1E9C4837981}\" />");

            // nonVisualDrawingPropertiesExtension19.Append(openXmlUnknownElement36);

            nonVisualDrawingPropertiesExtensionList19.Append(nonVisualDrawingPropertiesExtension19);

            nonVisualDrawingProperties29.Append(nonVisualDrawingPropertiesExtensionList19);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties10 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks10 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties10.Append(pictureLocks10);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties29 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties10.Append(nonVisualDrawingProperties29);
            nonVisualPictureProperties10.Append(nonVisualPictureDrawingProperties10);
            nonVisualPictureProperties10.Append(applicationNonVisualDrawingProperties29);

            BlipFill blipFill10 = new BlipFill();

            A.Blip blip14 = new A.Blip() { Embed = "rId8" };

            A.BlipExtensionList blipExtensionList8 = new A.BlipExtensionList();

            A.BlipExtension blipExtension8 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement37 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId9\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension8.Append(openXmlUnknownElement37);

            blipExtensionList8.Append(blipExtension8);

            blip14.Append(blipExtensionList8);

            A.Stretch stretch10 = new A.Stretch();
            A.FillRectangle fillRectangle10 = new A.FillRectangle();

            stretch10.Append(fillRectangle10);

            blipFill10.Append(blip14);
            blipFill10.Append(stretch10);

            ShapeProperties shapeProperties23 = new ShapeProperties();

            A.Transform2D transform2D22 = new A.Transform2D();
            A.Offset offset29 = new A.Offset() { X = 8257196L, Y = 2663344L };
            A.Extents extents29 = new A.Extents() { Cx = 228731L, Cy = 228571L };

            transform2D22.Append(offset29);
            transform2D22.Append(extents29);

            A.PresetGeometry presetGeometry22 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList22 = new A.AdjustValueList();

            presetGeometry22.Append(adjustValueList22);

            shapeProperties23.Append(transform2D22);
            shapeProperties23.Append(presetGeometry22);

            picture10.Append(nonVisualPictureProperties10);
            picture10.Append(blipFill10);
            picture10.Append(shapeProperties23);

            Shape shape13 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties13 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties30 = new NonVisualDrawingProperties() { Id = (UInt32Value)26U, Name = "Text 46" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList20 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension20 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement38 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{C49C8FAF-EDB1-74BF-B358-842F96B27376}\" />");

            // nonVisualDrawingPropertiesExtension20.Append(openXmlUnknownElement38);

            nonVisualDrawingPropertiesExtensionList20.Append(nonVisualDrawingPropertiesExtension20);

            nonVisualDrawingProperties30.Append(nonVisualDrawingPropertiesExtensionList20);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties13 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties30 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties13.Append(nonVisualDrawingProperties30);
            nonVisualShapeProperties13.Append(nonVisualShapeDrawingProperties13);
            nonVisualShapeProperties13.Append(applicationNonVisualDrawingProperties30);

            ShapeProperties shapeProperties24 = new ShapeProperties();

            A.Transform2D transform2D23 = new A.Transform2D();
            A.Offset offset30 = new A.Offset() { X = 8577419L, Y = 2651915L };
            A.Extents extents30 = new A.Extents() { Cx = 1280895L, Cy = 228571L };

            transform2D23.Append(offset30);
            transform2D23.Append(extents30);

            A.PresetGeometry presetGeometry23 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList23 = new A.AdjustValueList();

            presetGeometry23.Append(adjustValueList23);
            A.NoFill noFill148 = new A.NoFill();
            A.Outline outline25 = new A.Outline();

            shapeProperties24.Append(transform2D23);
            shapeProperties24.Append(presetGeometry23);
            shapeProperties24.Append(noFill148);
            shapeProperties24.Append(outline25);

            TextBody textBody34 = new TextBody();
            A.BodyProperties bodyProperties35 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle35 = new A.ListStyle();

            A.Paragraph paragraph38 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties29 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing22 = new A.LineSpacing();
            A.SpacingPercent spacingPercent22 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing22.Append(spacingPercent22);

            paragraphProperties29.Append(lineSpacing22);

            A.Run run23 = new A.Run();

            A.RunProperties runProperties25 = new A.RunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            A.SolidFill solidFill74 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex29 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha8 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex29.Append(alpha8);

            solidFill74.Append(rgbColorModelHex29);
            A.LatinFont latinFont54 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont44 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont54 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties25.Append(solidFill74);
            runProperties25.Append(latinFont54);
            runProperties25.Append(eastAsianFont44);
            runProperties25.Append(complexScriptFont54);
            A.Text text25 = new A.Text();

            // ********************************* Custom Code Start *********************************
            if (goalsArrayForSlide.Count > 5)
                text25.Text = goalsArrayForSlide.ElementAt(5).goalOwner.ToString() ?? "";
            else
                text25.Text = "NO DATA";
            // ********************************* Custom Code End ***********************************

            run23.Append(runProperties25);
            run23.Append(text25);
            A.EndParagraphRunProperties endParagraphRunProperties28 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            paragraph38.Append(paragraphProperties29);
            paragraph38.Append(run23);
            paragraph38.Append(endParagraphRunProperties28);

            textBody34.Append(bodyProperties35);
            textBody34.Append(listStyle35);
            textBody34.Append(paragraph38);

            shape13.Append(nonVisualShapeProperties13);
            shape13.Append(shapeProperties24);
            shape13.Append(textBody34);

            groupShape4.Append(nonVisualGroupShapeProperties6);
            groupShape4.Append(groupShapeProperties6);
            groupShape4.Append(picture9);
            groupShape4.Append(picture10);
            groupShape4.Append(shape13);

            GroupShape groupShape5 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties7 = new NonVisualGroupShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties31 = new NonVisualDrawingProperties() { Id = (UInt32Value)27U, Name = "Group 26" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList21 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension21 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement39 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{8FC11F99-23E7-5169-B092-4BF9476DA4F6}\" />");

            // nonVisualDrawingPropertiesExtension21.Append(openXmlUnknownElement39);

            nonVisualDrawingPropertiesExtensionList21.Append(nonVisualDrawingPropertiesExtension21);

            nonVisualDrawingProperties31.Append(nonVisualDrawingPropertiesExtensionList21);
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties7 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties31 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties7.Append(nonVisualDrawingProperties31);
            nonVisualGroupShapeProperties7.Append(nonVisualGroupShapeDrawingProperties7);
            nonVisualGroupShapeProperties7.Append(applicationNonVisualDrawingProperties31);

            GroupShapeProperties groupShapeProperties7 = new GroupShapeProperties();

            A.TransformGroup transformGroup7 = new A.TransformGroup();
            A.Offset offset31 = new A.Offset() { X = 10252766L, Y = 1771646L };
            A.Extents extents31 = new A.Extents() { Cx = 869179L, Cy = 274285L };
            A.ChildOffset childOffset7 = new A.ChildOffset() { X = 10361522L, Y = 2640488L };
            A.ChildExtents childExtents7 = new A.ChildExtents() { Cx = 869179L, Cy = 274285L };

            transformGroup7.Append(offset31);
            transformGroup7.Append(extents31);
            transformGroup7.Append(childOffset7);
            transformGroup7.Append(childExtents7);

            groupShapeProperties7.Append(transformGroup7);

            Picture picture11 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties11 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties32 = new NonVisualDrawingProperties() { Id = (UInt32Value)28U, Name = "Image 22", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList22 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension22 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement40 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{039F62CF-DB33-D247-84FD-C2756C7C5777}\" />");

            // nonVisualDrawingPropertiesExtension22.Append(openXmlUnknownElement40);

            nonVisualDrawingPropertiesExtensionList22.Append(nonVisualDrawingPropertiesExtension22);

            nonVisualDrawingProperties32.Append(nonVisualDrawingPropertiesExtensionList22);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties11 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks11 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties11.Append(pictureLocks11);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties32 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties11.Append(nonVisualDrawingProperties32);
            nonVisualPictureProperties11.Append(nonVisualPictureDrawingProperties11);
            nonVisualPictureProperties11.Append(applicationNonVisualDrawingProperties32);

            BlipFill blipFill11 = new BlipFill();
            A.Blip blip15 = new A.Blip() { Embed = "rId10" };

            A.Stretch stretch11 = new A.Stretch();
            A.FillRectangle fillRectangle11 = new A.FillRectangle();

            stretch11.Append(fillRectangle11);

            blipFill11.Append(blip15);
            blipFill11.Append(stretch11);

            ShapeProperties shapeProperties25 = new ShapeProperties();

            A.Transform2D transform2D24 = new A.Transform2D();
            A.Offset offset32 = new A.Offset() { X = 10361522L, Y = 2640488L };
            A.Extents extents32 = new A.Extents() { Cx = 869179L, Cy = 274285L };

            transform2D24.Append(offset32);
            transform2D24.Append(extents32);

            A.PresetGeometry presetGeometry24 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList24 = new A.AdjustValueList();

            presetGeometry24.Append(adjustValueList24);

            shapeProperties25.Append(transform2D24);
            shapeProperties25.Append(presetGeometry24);

            picture11.Append(nonVisualPictureProperties11);
            picture11.Append(blipFill11);
            picture11.Append(shapeProperties25);

            Picture picture12 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties12 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties33 = new NonVisualDrawingProperties() { Id = (UInt32Value)29U, Name = "Image 24", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList23 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension23 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement41 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{49A1CBA4-06BD-C726-63E7-C39634592388}\" />");

            // nonVisualDrawingPropertiesExtension23.Append(openXmlUnknownElement41);

            nonVisualDrawingPropertiesExtensionList23.Append(nonVisualDrawingPropertiesExtension23);

            nonVisualDrawingProperties33.Append(nonVisualDrawingPropertiesExtensionList23);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties12 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks12 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties12.Append(pictureLocks12);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties33 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties12.Append(nonVisualDrawingProperties33);
            nonVisualPictureProperties12.Append(nonVisualPictureDrawingProperties12);
            nonVisualPictureProperties12.Append(applicationNonVisualDrawingProperties33);

            BlipFill blipFill12 = new BlipFill();

            A.Blip blip16 = new A.Blip() { Embed = "rId11" };

            A.BlipExtensionList blipExtensionList9 = new A.BlipExtensionList();

            A.BlipExtension blipExtension9 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement42 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId12\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension9.Append(openXmlUnknownElement42);

            blipExtensionList9.Append(blipExtension9);

            blip16.Append(blipExtensionList9);

            A.Stretch stretch12 = new A.Stretch();
            A.FillRectangle fillRectangle12 = new A.FillRectangle();

            stretch12.Append(fillRectangle12);

            blipFill12.Append(blip16);
            blipFill12.Append(stretch12);

            ShapeProperties shapeProperties26 = new ShapeProperties();

            A.Transform2D transform2D25 = new A.Transform2D();
            A.Offset offset33 = new A.Offset() { X = 10453015L, Y = 2709059L };
            A.Extents extents33 = new A.Extents() { Cx = 137250L, Cy = 137157L };

            transform2D25.Append(offset33);
            transform2D25.Append(extents33);

            A.PresetGeometry presetGeometry25 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList25 = new A.AdjustValueList();

            presetGeometry25.Append(adjustValueList25);

            shapeProperties26.Append(transform2D25);
            shapeProperties26.Append(presetGeometry25);

            picture12.Append(nonVisualPictureProperties12);
            picture12.Append(blipFill12);
            picture12.Append(shapeProperties26);

            Shape shape14 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties14 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties34 = new NonVisualDrawingProperties() { Id = (UInt32Value)30U, Name = "Text 48" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList24 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension24 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement43 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{2B183206-CF90-918D-ADEA-A89D30DC4208}\" />");

            // nonVisualDrawingPropertiesExtension24.Append(openXmlUnknownElement43);

            nonVisualDrawingPropertiesExtensionList24.Append(nonVisualDrawingPropertiesExtension24);

            nonVisualDrawingProperties34.Append(nonVisualDrawingPropertiesExtensionList24);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties14 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties34 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties14.Append(nonVisualDrawingProperties34);
            nonVisualShapeProperties14.Append(nonVisualShapeDrawingProperties14);
            nonVisualShapeProperties14.Append(applicationNonVisualDrawingProperties34);

            ShapeProperties shapeProperties27 = new ShapeProperties();

            A.Transform2D transform2D26 = new A.Transform2D();
            A.Offset offset34 = new A.Offset() { X = 10658873L, Y = 2651915L };
            A.Extents extents34 = new A.Extents() { Cx = 526081L, Cy = 228571L };

            transform2D26.Append(offset34);
            transform2D26.Append(extents34);

            A.PresetGeometry presetGeometry26 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList26 = new A.AdjustValueList();

            presetGeometry26.Append(adjustValueList26);
            A.NoFill noFill149 = new A.NoFill();
            A.Outline outline26 = new A.Outline();

            shapeProperties27.Append(transform2D26);
            shapeProperties27.Append(presetGeometry26);
            shapeProperties27.Append(noFill149);
            shapeProperties27.Append(outline26);

            TextBody textBody35 = new TextBody();
            A.BodyProperties bodyProperties36 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle36 = new A.ListStyle();

            A.Paragraph paragraph39 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties30 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing23 = new A.LineSpacing();
            A.SpacingPercent spacingPercent23 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing23.Append(spacingPercent23);

            paragraphProperties30.Append(lineSpacing23);

            A.Run run24 = new A.Run();

            A.RunProperties runProperties26 = new A.RunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            A.SolidFill solidFill75 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex30 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha9 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex30.Append(alpha9);

            solidFill75.Append(rgbColorModelHex30);
            A.LatinFont latinFont55 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont45 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont55 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties26.Append(solidFill75);
            runProperties26.Append(latinFont55);
            runProperties26.Append(eastAsianFont45);
            runProperties26.Append(complexScriptFont55);
            A.Text text26 = new A.Text();
            text26.Text = goalsArrayForSlide.ElementAt(0).target.ToString() ?? "";

            run24.Append(runProperties26);
            run24.Append(text26);
            A.EndParagraphRunProperties endParagraphRunProperties29 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            paragraph39.Append(paragraphProperties30);
            paragraph39.Append(run24);
            paragraph39.Append(endParagraphRunProperties29);

            textBody35.Append(bodyProperties36);
            textBody35.Append(listStyle36);
            textBody35.Append(paragraph39);

            shape14.Append(nonVisualShapeProperties14);
            shape14.Append(shapeProperties27);
            shape14.Append(textBody35);

            groupShape5.Append(nonVisualGroupShapeProperties7);
            groupShape5.Append(groupShapeProperties7);
            groupShape5.Append(picture11);
            groupShape5.Append(picture12);
            groupShape5.Append(shape14);

            GroupShape groupShape6 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties8 = new NonVisualGroupShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties35 = new NonVisualDrawingProperties() { Id = (UInt32Value)39U, Name = "Group 38" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList25 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension25 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement44 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{28061249-CE2B-26C4-A55F-A2884A3F8817}\" />");

            // nonVisualDrawingPropertiesExtension25.Append(openXmlUnknownElement44);

            nonVisualDrawingPropertiesExtensionList25.Append(nonVisualDrawingPropertiesExtension25);

            nonVisualDrawingProperties35.Append(nonVisualDrawingPropertiesExtensionList25);
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties8 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties35 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties8.Append(nonVisualDrawingProperties35);
            nonVisualGroupShapeProperties8.Append(nonVisualGroupShapeDrawingProperties8);
            nonVisualGroupShapeProperties8.Append(applicationNonVisualDrawingProperties35);

            GroupShapeProperties groupShapeProperties8 = new GroupShapeProperties();

            A.TransformGroup transformGroup8 = new A.TransformGroup();
            A.Offset offset35 = new A.Offset() { X = 10252766L, Y = 2525794L };
            A.Extents extents35 = new A.Extents() { Cx = 869179L, Cy = 274285L };
            A.ChildOffset childOffset8 = new A.ChildOffset() { X = 10361522L, Y = 2640488L };
            A.ChildExtents childExtents8 = new A.ChildExtents() { Cx = 869179L, Cy = 274285L };

            transformGroup8.Append(offset35);
            transformGroup8.Append(extents35);
            transformGroup8.Append(childOffset8);
            transformGroup8.Append(childExtents8);

            groupShapeProperties8.Append(transformGroup8);

            Picture picture13 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties13 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties36 = new NonVisualDrawingProperties() { Id = (UInt32Value)40U, Name = "Image 22", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList26 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension26 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement45 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{30D6F5FA-420D-2786-E0F0-27783A3515B8}\" />");

            // nonVisualDrawingPropertiesExtension26.Append(openXmlUnknownElement45);

            nonVisualDrawingPropertiesExtensionList26.Append(nonVisualDrawingPropertiesExtension26);

            nonVisualDrawingProperties36.Append(nonVisualDrawingPropertiesExtensionList26);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties13 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks13 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties13.Append(pictureLocks13);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties36 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties13.Append(nonVisualDrawingProperties36);
            nonVisualPictureProperties13.Append(nonVisualPictureDrawingProperties13);
            nonVisualPictureProperties13.Append(applicationNonVisualDrawingProperties36);

            BlipFill blipFill13 = new BlipFill();
            A.Blip blip17 = new A.Blip() { Embed = "rId10" };

            A.Stretch stretch13 = new A.Stretch();
            A.FillRectangle fillRectangle13 = new A.FillRectangle();

            stretch13.Append(fillRectangle13);

            blipFill13.Append(blip17);
            blipFill13.Append(stretch13);

            ShapeProperties shapeProperties28 = new ShapeProperties();

            A.Transform2D transform2D27 = new A.Transform2D();
            A.Offset offset36 = new A.Offset() { X = 10361522L, Y = 2640488L };
            A.Extents extents36 = new A.Extents() { Cx = 869179L, Cy = 274285L };

            transform2D27.Append(offset36);
            transform2D27.Append(extents36);

            A.PresetGeometry presetGeometry27 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList27 = new A.AdjustValueList();

            presetGeometry27.Append(adjustValueList27);

            shapeProperties28.Append(transform2D27);
            shapeProperties28.Append(presetGeometry27);

            picture13.Append(nonVisualPictureProperties13);
            picture13.Append(blipFill13);
            picture13.Append(shapeProperties28);

            Picture picture14 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties14 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties37 = new NonVisualDrawingProperties() { Id = (UInt32Value)41U, Name = "Image 24", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList27 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension27 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement46 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{3C1E58C7-C3E7-D26D-2987-BEBE3DCA0E8A}\" />");

            // nonVisualDrawingPropertiesExtension27.Append(openXmlUnknownElement46);

            nonVisualDrawingPropertiesExtensionList27.Append(nonVisualDrawingPropertiesExtension27);

            nonVisualDrawingProperties37.Append(nonVisualDrawingPropertiesExtensionList27);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties14 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks14 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties14.Append(pictureLocks14);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties37 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties14.Append(nonVisualDrawingProperties37);
            nonVisualPictureProperties14.Append(nonVisualPictureDrawingProperties14);
            nonVisualPictureProperties14.Append(applicationNonVisualDrawingProperties37);

            BlipFill blipFill14 = new BlipFill();

            A.Blip blip18 = new A.Blip() { Embed = "rId11" };

            A.BlipExtensionList blipExtensionList10 = new A.BlipExtensionList();

            A.BlipExtension blipExtension10 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement47 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId12\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension10.Append(openXmlUnknownElement47);

            blipExtensionList10.Append(blipExtension10);

            blip18.Append(blipExtensionList10);

            A.Stretch stretch14 = new A.Stretch();
            A.FillRectangle fillRectangle14 = new A.FillRectangle();

            stretch14.Append(fillRectangle14);

            blipFill14.Append(blip18);
            blipFill14.Append(stretch14);

            ShapeProperties shapeProperties29 = new ShapeProperties();

            A.Transform2D transform2D28 = new A.Transform2D();
            A.Offset offset37 = new A.Offset() { X = 10453015L, Y = 2709059L };
            A.Extents extents37 = new A.Extents() { Cx = 137250L, Cy = 137157L };

            transform2D28.Append(offset37);
            transform2D28.Append(extents37);

            A.PresetGeometry presetGeometry28 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList28 = new A.AdjustValueList();

            presetGeometry28.Append(adjustValueList28);

            shapeProperties29.Append(transform2D28);
            shapeProperties29.Append(presetGeometry28);

            picture14.Append(nonVisualPictureProperties14);
            picture14.Append(blipFill14);
            picture14.Append(shapeProperties29);

            Shape shape15 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties15 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties38 = new NonVisualDrawingProperties() { Id = (UInt32Value)42U, Name = "Text 48" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList28 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension28 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement48 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{B7056AA0-38B7-D978-8CB0-F827C6B15B89}\" />");

            // nonVisualDrawingPropertiesExtension28.Append(openXmlUnknownElement48);

            nonVisualDrawingPropertiesExtensionList28.Append(nonVisualDrawingPropertiesExtension28);

            nonVisualDrawingProperties38.Append(nonVisualDrawingPropertiesExtensionList28);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties15 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties38 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties15.Append(nonVisualDrawingProperties38);
            nonVisualShapeProperties15.Append(nonVisualShapeDrawingProperties15);
            nonVisualShapeProperties15.Append(applicationNonVisualDrawingProperties38);

            ShapeProperties shapeProperties30 = new ShapeProperties();

            A.Transform2D transform2D29 = new A.Transform2D();
            A.Offset offset38 = new A.Offset() { X = 10658873L, Y = 2651915L };
            A.Extents extents38 = new A.Extents() { Cx = 526081L, Cy = 228571L };

            transform2D29.Append(offset38);
            transform2D29.Append(extents38);

            A.PresetGeometry presetGeometry29 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList29 = new A.AdjustValueList();

            presetGeometry29.Append(adjustValueList29);
            A.NoFill noFill150 = new A.NoFill();
            A.Outline outline27 = new A.Outline();

            shapeProperties30.Append(transform2D29);
            shapeProperties30.Append(presetGeometry29);
            shapeProperties30.Append(noFill150);
            shapeProperties30.Append(outline27);

            TextBody textBody36 = new TextBody();
            A.BodyProperties bodyProperties37 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle37 = new A.ListStyle();

            A.Paragraph paragraph40 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties31 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing24 = new A.LineSpacing();
            A.SpacingPercent spacingPercent24 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing24.Append(spacingPercent24);

            paragraphProperties31.Append(lineSpacing24);

            A.Run run25 = new A.Run();

            A.RunProperties runProperties27 = new A.RunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            A.SolidFill solidFill76 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex31 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha10 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex31.Append(alpha10);

            solidFill76.Append(rgbColorModelHex31);
            A.LatinFont latinFont56 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont46 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont56 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties27.Append(solidFill76);
            runProperties27.Append(latinFont56);
            runProperties27.Append(eastAsianFont46);
            runProperties27.Append(complexScriptFont56);
            A.Text text27 = new A.Text();

            // ********************************* Custom Code Start *********************************
            if (goalsArrayForSlide.Count > 1)
                text27.Text = goalsArrayForSlide.ElementAt(1).target.ToString() ?? "";
            else
                text27.Text = "NO DATA";
            // ********************************* Custom Code End *********************************

            run25.Append(runProperties27);
            run25.Append(text27);
            A.EndParagraphRunProperties endParagraphRunProperties30 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            paragraph40.Append(paragraphProperties31);
            paragraph40.Append(run25);
            paragraph40.Append(endParagraphRunProperties30);

            textBody36.Append(bodyProperties37);
            textBody36.Append(listStyle37);
            textBody36.Append(paragraph40);

            shape15.Append(nonVisualShapeProperties15);
            shape15.Append(shapeProperties30);
            shape15.Append(textBody36);

            groupShape6.Append(nonVisualGroupShapeProperties8);
            groupShape6.Append(groupShapeProperties8);
            groupShape6.Append(picture13);
            groupShape6.Append(picture14);
            groupShape6.Append(shape15);

            GroupShape groupShape7 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties9 = new NonVisualGroupShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties39 = new NonVisualDrawingProperties() { Id = (UInt32Value)43U, Name = "Group 42" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList29 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension29 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement49 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{977FA258-08FC-269F-9F2B-CEB8209948CA}\" />");

            // nonVisualDrawingPropertiesExtension29.Append(openXmlUnknownElement49);

            nonVisualDrawingPropertiesExtensionList29.Append(nonVisualDrawingPropertiesExtension29);

            nonVisualDrawingProperties39.Append(nonVisualDrawingPropertiesExtensionList29);
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties9 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties39 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties9.Append(nonVisualDrawingProperties39);
            nonVisualGroupShapeProperties9.Append(nonVisualGroupShapeDrawingProperties9);
            nonVisualGroupShapeProperties9.Append(applicationNonVisualDrawingProperties39);

            GroupShapeProperties groupShapeProperties9 = new GroupShapeProperties();

            A.TransformGroup transformGroup9 = new A.TransformGroup();
            A.Offset offset39 = new A.Offset() { X = 10252766L, Y = 3272174L };
            A.Extents extents39 = new A.Extents() { Cx = 869179L, Cy = 274285L };
            A.ChildOffset childOffset9 = new A.ChildOffset() { X = 10361522L, Y = 2640488L };
            A.ChildExtents childExtents9 = new A.ChildExtents() { Cx = 869179L, Cy = 274285L };

            transformGroup9.Append(offset39);
            transformGroup9.Append(extents39);
            transformGroup9.Append(childOffset9);
            transformGroup9.Append(childExtents9);

            groupShapeProperties9.Append(transformGroup9);

            Picture picture15 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties15 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties40 = new NonVisualDrawingProperties() { Id = (UInt32Value)44U, Name = "Image 22", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList30 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension30 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement50 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{30B3EE5D-EE6A-44FD-AF36-1865E494C377}\" />");

            // nonVisualDrawingPropertiesExtension30.Append(openXmlUnknownElement50);

            nonVisualDrawingPropertiesExtensionList30.Append(nonVisualDrawingPropertiesExtension30);

            nonVisualDrawingProperties40.Append(nonVisualDrawingPropertiesExtensionList30);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties15 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks15 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties15.Append(pictureLocks15);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties40 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties15.Append(nonVisualDrawingProperties40);
            nonVisualPictureProperties15.Append(nonVisualPictureDrawingProperties15);
            nonVisualPictureProperties15.Append(applicationNonVisualDrawingProperties40);

            BlipFill blipFill15 = new BlipFill();
            A.Blip blip19 = new A.Blip() { Embed = "rId10" };

            A.Stretch stretch15 = new A.Stretch();
            A.FillRectangle fillRectangle15 = new A.FillRectangle();

            stretch15.Append(fillRectangle15);

            blipFill15.Append(blip19);
            blipFill15.Append(stretch15);

            ShapeProperties shapeProperties31 = new ShapeProperties();

            A.Transform2D transform2D30 = new A.Transform2D();
            A.Offset offset40 = new A.Offset() { X = 10361522L, Y = 2640488L };
            A.Extents extents40 = new A.Extents() { Cx = 869179L, Cy = 274285L };

            transform2D30.Append(offset40);
            transform2D30.Append(extents40);

            A.PresetGeometry presetGeometry30 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList30 = new A.AdjustValueList();

            presetGeometry30.Append(adjustValueList30);

            shapeProperties31.Append(transform2D30);
            shapeProperties31.Append(presetGeometry30);

            picture15.Append(nonVisualPictureProperties15);
            picture15.Append(blipFill15);
            picture15.Append(shapeProperties31);

            Picture picture16 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties16 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties41 = new NonVisualDrawingProperties() { Id = (UInt32Value)45U, Name = "Image 24", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList31 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension31 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement51 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{6F087B49-FE6C-21CD-377F-9C4C3B9A55A5}\" />");

            // nonVisualDrawingPropertiesExtension31.Append(openXmlUnknownElement51);

            nonVisualDrawingPropertiesExtensionList31.Append(nonVisualDrawingPropertiesExtension31);

            nonVisualDrawingProperties41.Append(nonVisualDrawingPropertiesExtensionList31);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties16 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks16 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties16.Append(pictureLocks16);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties41 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties16.Append(nonVisualDrawingProperties41);
            nonVisualPictureProperties16.Append(nonVisualPictureDrawingProperties16);
            nonVisualPictureProperties16.Append(applicationNonVisualDrawingProperties41);

            BlipFill blipFill16 = new BlipFill();

            A.Blip blip20 = new A.Blip() { Embed = "rId11" };

            A.BlipExtensionList blipExtensionList11 = new A.BlipExtensionList();

            A.BlipExtension blipExtension11 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement52 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId12\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension11.Append(openXmlUnknownElement52);

            blipExtensionList11.Append(blipExtension11);

            blip20.Append(blipExtensionList11);

            A.Stretch stretch16 = new A.Stretch();
            A.FillRectangle fillRectangle16 = new A.FillRectangle();

            stretch16.Append(fillRectangle16);

            blipFill16.Append(blip20);
            blipFill16.Append(stretch16);

            ShapeProperties shapeProperties32 = new ShapeProperties();

            A.Transform2D transform2D31 = new A.Transform2D();
            A.Offset offset41 = new A.Offset() { X = 10453015L, Y = 2709059L };
            A.Extents extents41 = new A.Extents() { Cx = 137250L, Cy = 137157L };

            transform2D31.Append(offset41);
            transform2D31.Append(extents41);

            A.PresetGeometry presetGeometry31 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList31 = new A.AdjustValueList();

            presetGeometry31.Append(adjustValueList31);

            shapeProperties32.Append(transform2D31);
            shapeProperties32.Append(presetGeometry31);

            picture16.Append(nonVisualPictureProperties16);
            picture16.Append(blipFill16);
            picture16.Append(shapeProperties32);

            Shape shape16 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties16 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties42 = new NonVisualDrawingProperties() { Id = (UInt32Value)46U, Name = "Text 48" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList32 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension32 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement53 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{C21CD9FA-DAFE-4BA8-B147-F196936996E9}\" />");

            // nonVisualDrawingPropertiesExtension32.Append(openXmlUnknownElement53);

            nonVisualDrawingPropertiesExtensionList32.Append(nonVisualDrawingPropertiesExtension32);

            nonVisualDrawingProperties42.Append(nonVisualDrawingPropertiesExtensionList32);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties16 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties42 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties16.Append(nonVisualDrawingProperties42);
            nonVisualShapeProperties16.Append(nonVisualShapeDrawingProperties16);
            nonVisualShapeProperties16.Append(applicationNonVisualDrawingProperties42);

            ShapeProperties shapeProperties33 = new ShapeProperties();

            A.Transform2D transform2D32 = new A.Transform2D();
            A.Offset offset42 = new A.Offset() { X = 10658873L, Y = 2651915L };
            A.Extents extents42 = new A.Extents() { Cx = 526081L, Cy = 228571L };

            transform2D32.Append(offset42);
            transform2D32.Append(extents42);

            A.PresetGeometry presetGeometry32 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList32 = new A.AdjustValueList();

            presetGeometry32.Append(adjustValueList32);
            A.NoFill noFill151 = new A.NoFill();
            A.Outline outline28 = new A.Outline();

            shapeProperties33.Append(transform2D32);
            shapeProperties33.Append(presetGeometry32);
            shapeProperties33.Append(noFill151);
            shapeProperties33.Append(outline28);

            TextBody textBody37 = new TextBody();
            A.BodyProperties bodyProperties38 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle38 = new A.ListStyle();

            A.Paragraph paragraph41 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties32 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing25 = new A.LineSpacing();
            A.SpacingPercent spacingPercent25 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing25.Append(spacingPercent25);

            paragraphProperties32.Append(lineSpacing25);

            A.Run run26 = new A.Run();

            A.RunProperties runProperties28 = new A.RunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            A.SolidFill solidFill77 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex32 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha11 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex32.Append(alpha11);

            solidFill77.Append(rgbColorModelHex32);
            A.LatinFont latinFont57 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont47 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont57 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties28.Append(solidFill77);
            runProperties28.Append(latinFont57);
            runProperties28.Append(eastAsianFont47);
            runProperties28.Append(complexScriptFont57);
            A.Text text28 = new A.Text();

            // ********************************* Custom Code Start *********************************
            if (goalsArrayForSlide.Count > 2)
                text28.Text = goalsArrayForSlide.ElementAt(2).target.ToString() ?? "";
            else
                text28.Text = "NO DATA";
            // ********************************* Custom Code End *********************************

            run26.Append(runProperties28);
            run26.Append(text28);
            A.EndParagraphRunProperties endParagraphRunProperties31 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            paragraph41.Append(paragraphProperties32);
            paragraph41.Append(run26);
            paragraph41.Append(endParagraphRunProperties31);

            textBody37.Append(bodyProperties38);
            textBody37.Append(listStyle38);
            textBody37.Append(paragraph41);

            shape16.Append(nonVisualShapeProperties16);
            shape16.Append(shapeProperties33);
            shape16.Append(textBody37);

            groupShape7.Append(nonVisualGroupShapeProperties9);
            groupShape7.Append(groupShapeProperties9);
            groupShape7.Append(picture15);
            groupShape7.Append(picture16);
            groupShape7.Append(shape16);

            GroupShape groupShape8 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties10 = new NonVisualGroupShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties43 = new NonVisualDrawingProperties() { Id = (UInt32Value)47U, Name = "Group 46" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList33 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension33 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement54 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{5CBD79B3-E4A8-5A80-BCDB-AD9E788CF3A7}\" />");

            // nonVisualDrawingPropertiesExtension33.Append(openXmlUnknownElement54);

            nonVisualDrawingPropertiesExtensionList33.Append(nonVisualDrawingPropertiesExtension33);

            nonVisualDrawingProperties43.Append(nonVisualDrawingPropertiesExtensionList33);
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties10 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties43 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties10.Append(nonVisualDrawingProperties43);
            nonVisualGroupShapeProperties10.Append(nonVisualGroupShapeDrawingProperties10);
            nonVisualGroupShapeProperties10.Append(applicationNonVisualDrawingProperties43);

            GroupShapeProperties groupShapeProperties10 = new GroupShapeProperties();

            A.TransformGroup transformGroup10 = new A.TransformGroup();
            A.Offset offset43 = new A.Offset() { X = 10252766L, Y = 3997973L };
            A.Extents extents43 = new A.Extents() { Cx = 869179L, Cy = 274285L };
            A.ChildOffset childOffset10 = new A.ChildOffset() { X = 10361522L, Y = 2640488L };
            A.ChildExtents childExtents10 = new A.ChildExtents() { Cx = 869179L, Cy = 274285L };

            transformGroup10.Append(offset43);
            transformGroup10.Append(extents43);
            transformGroup10.Append(childOffset10);
            transformGroup10.Append(childExtents10);

            groupShapeProperties10.Append(transformGroup10);

            Picture picture17 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties17 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties44 = new NonVisualDrawingProperties() { Id = (UInt32Value)48U, Name = "Image 22", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList34 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension34 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement55 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{14FDEAD0-0B58-01D7-9455-D5F5FAA232EE}\" />");

            // nonVisualDrawingPropertiesExtension34.Append(openXmlUnknownElement55);

            nonVisualDrawingPropertiesExtensionList34.Append(nonVisualDrawingPropertiesExtension34);

            nonVisualDrawingProperties44.Append(nonVisualDrawingPropertiesExtensionList34);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties17 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks17 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties17.Append(pictureLocks17);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties44 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties17.Append(nonVisualDrawingProperties44);
            nonVisualPictureProperties17.Append(nonVisualPictureDrawingProperties17);
            nonVisualPictureProperties17.Append(applicationNonVisualDrawingProperties44);

            BlipFill blipFill17 = new BlipFill();
            A.Blip blip21 = new A.Blip() { Embed = "rId10" };

            A.Stretch stretch17 = new A.Stretch();
            A.FillRectangle fillRectangle17 = new A.FillRectangle();

            stretch17.Append(fillRectangle17);

            blipFill17.Append(blip21);
            blipFill17.Append(stretch17);

            ShapeProperties shapeProperties34 = new ShapeProperties();

            A.Transform2D transform2D33 = new A.Transform2D();
            A.Offset offset44 = new A.Offset() { X = 10361522L, Y = 2640488L };
            A.Extents extents44 = new A.Extents() { Cx = 869179L, Cy = 274285L };

            transform2D33.Append(offset44);
            transform2D33.Append(extents44);

            A.PresetGeometry presetGeometry33 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList33 = new A.AdjustValueList();

            presetGeometry33.Append(adjustValueList33);

            shapeProperties34.Append(transform2D33);
            shapeProperties34.Append(presetGeometry33);

            picture17.Append(nonVisualPictureProperties17);
            picture17.Append(blipFill17);
            picture17.Append(shapeProperties34);

            Picture picture18 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties18 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties45 = new NonVisualDrawingProperties() { Id = (UInt32Value)49U, Name = "Image 24", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList35 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension35 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement56 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{0E6B8C9F-6143-EEAE-75ED-885CE84D7192}\" />");

            // nonVisualDrawingPropertiesExtension35.Append(openXmlUnknownElement56);

            nonVisualDrawingPropertiesExtensionList35.Append(nonVisualDrawingPropertiesExtension35);

            nonVisualDrawingProperties45.Append(nonVisualDrawingPropertiesExtensionList35);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties18 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks18 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties18.Append(pictureLocks18);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties45 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties18.Append(nonVisualDrawingProperties45);
            nonVisualPictureProperties18.Append(nonVisualPictureDrawingProperties18);
            nonVisualPictureProperties18.Append(applicationNonVisualDrawingProperties45);

            BlipFill blipFill18 = new BlipFill();

            A.Blip blip22 = new A.Blip() { Embed = "rId11" };

            A.BlipExtensionList blipExtensionList12 = new A.BlipExtensionList();

            A.BlipExtension blipExtension12 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement57 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId12\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension12.Append(openXmlUnknownElement57);

            blipExtensionList12.Append(blipExtension12);

            blip22.Append(blipExtensionList12);

            A.Stretch stretch18 = new A.Stretch();
            A.FillRectangle fillRectangle18 = new A.FillRectangle();

            stretch18.Append(fillRectangle18);

            blipFill18.Append(blip22);
            blipFill18.Append(stretch18);

            ShapeProperties shapeProperties35 = new ShapeProperties();

            A.Transform2D transform2D34 = new A.Transform2D();
            A.Offset offset45 = new A.Offset() { X = 10453015L, Y = 2709059L };
            A.Extents extents45 = new A.Extents() { Cx = 137250L, Cy = 137157L };

            transform2D34.Append(offset45);
            transform2D34.Append(extents45);

            A.PresetGeometry presetGeometry34 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList34 = new A.AdjustValueList();

            presetGeometry34.Append(adjustValueList34);

            shapeProperties35.Append(transform2D34);
            shapeProperties35.Append(presetGeometry34);

            picture18.Append(nonVisualPictureProperties18);
            picture18.Append(blipFill18);
            picture18.Append(shapeProperties35);

            Shape shape17 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties17 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties46 = new NonVisualDrawingProperties() { Id = (UInt32Value)50U, Name = "Text 48" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList36 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension36 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement58 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{4AFB32C5-CF71-E3C0-5000-DF75FD9552EC}\" />");

            // nonVisualDrawingPropertiesExtension36.Append(openXmlUnknownElement58);

            nonVisualDrawingPropertiesExtensionList36.Append(nonVisualDrawingPropertiesExtension36);

            nonVisualDrawingProperties46.Append(nonVisualDrawingPropertiesExtensionList36);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties17 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties46 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties17.Append(nonVisualDrawingProperties46);
            nonVisualShapeProperties17.Append(nonVisualShapeDrawingProperties17);
            nonVisualShapeProperties17.Append(applicationNonVisualDrawingProperties46);

            ShapeProperties shapeProperties36 = new ShapeProperties();

            A.Transform2D transform2D35 = new A.Transform2D();
            A.Offset offset46 = new A.Offset() { X = 10658873L, Y = 2651915L };
            A.Extents extents46 = new A.Extents() { Cx = 526081L, Cy = 228571L };

            transform2D35.Append(offset46);
            transform2D35.Append(extents46);

            A.PresetGeometry presetGeometry35 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList35 = new A.AdjustValueList();

            presetGeometry35.Append(adjustValueList35);
            A.NoFill noFill152 = new A.NoFill();
            A.Outline outline29 = new A.Outline();

            shapeProperties36.Append(transform2D35);
            shapeProperties36.Append(presetGeometry35);
            shapeProperties36.Append(noFill152);
            shapeProperties36.Append(outline29);

            TextBody textBody38 = new TextBody();
            A.BodyProperties bodyProperties39 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle39 = new A.ListStyle();

            A.Paragraph paragraph42 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties33 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing26 = new A.LineSpacing();
            A.SpacingPercent spacingPercent26 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing26.Append(spacingPercent26);

            paragraphProperties33.Append(lineSpacing26);

            A.Run run27 = new A.Run();

            A.RunProperties runProperties29 = new A.RunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            A.SolidFill solidFill78 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex33 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha12 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex33.Append(alpha12);

            solidFill78.Append(rgbColorModelHex33);
            A.LatinFont latinFont58 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont48 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont58 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties29.Append(solidFill78);
            runProperties29.Append(latinFont58);
            runProperties29.Append(eastAsianFont48);
            runProperties29.Append(complexScriptFont58);
            A.Text text29 = new A.Text();

            // ********************************* Custom Code Start *********************************
            if (goalsArrayForSlide.Count > 3)
                text29.Text = goalsArrayForSlide.ElementAt(3).target.ToString() ?? "";
            else
                text29.Text = "NO DATA";
            // ********************************* Custom Code End *********************************

            run27.Append(runProperties29);
            run27.Append(text29);
            A.EndParagraphRunProperties endParagraphRunProperties32 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            paragraph42.Append(paragraphProperties33);
            paragraph42.Append(run27);
            paragraph42.Append(endParagraphRunProperties32);

            textBody38.Append(bodyProperties39);
            textBody38.Append(listStyle39);
            textBody38.Append(paragraph42);

            shape17.Append(nonVisualShapeProperties17);
            shape17.Append(shapeProperties36);
            shape17.Append(textBody38);

            groupShape8.Append(nonVisualGroupShapeProperties10);
            groupShape8.Append(groupShapeProperties10);
            groupShape8.Append(picture17);
            groupShape8.Append(picture18);
            groupShape8.Append(shape17);

            GroupShape groupShape9 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties11 = new NonVisualGroupShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties47 = new NonVisualDrawingProperties() { Id = (UInt32Value)51U, Name = "Group 50" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList37 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension37 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement59 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{9492CBBC-D2F7-C088-7352-EED3E971D212}\" />");

            // nonVisualDrawingPropertiesExtension37.Append(openXmlUnknownElement59);

            nonVisualDrawingPropertiesExtensionList37.Append(nonVisualDrawingPropertiesExtension37);

            nonVisualDrawingProperties47.Append(nonVisualDrawingPropertiesExtensionList37);
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties11 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties47 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties11.Append(nonVisualDrawingProperties47);
            nonVisualGroupShapeProperties11.Append(nonVisualGroupShapeDrawingProperties11);
            nonVisualGroupShapeProperties11.Append(applicationNonVisualDrawingProperties47);

            GroupShapeProperties groupShapeProperties11 = new GroupShapeProperties();

            A.TransformGroup transformGroup11 = new A.TransformGroup();
            A.Offset offset47 = new A.Offset() { X = 10252766L, Y = 5631203L };
            A.Extents extents47 = new A.Extents() { Cx = 869179L, Cy = 274285L };
            A.ChildOffset childOffset11 = new A.ChildOffset() { X = 10361522L, Y = 2640488L };
            A.ChildExtents childExtents11 = new A.ChildExtents() { Cx = 869179L, Cy = 274285L };

            transformGroup11.Append(offset47);
            transformGroup11.Append(extents47);
            transformGroup11.Append(childOffset11);
            transformGroup11.Append(childExtents11);

            groupShapeProperties11.Append(transformGroup11);

            Picture picture19 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties19 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties48 = new NonVisualDrawingProperties() { Id = (UInt32Value)52U, Name = "Image 22", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList38 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension38 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement60 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{45046161-276E-2474-BECF-B8527859591A}\" />");

            // nonVisualDrawingPropertiesExtension38.Append(openXmlUnknownElement60);

            nonVisualDrawingPropertiesExtensionList38.Append(nonVisualDrawingPropertiesExtension38);

            nonVisualDrawingProperties48.Append(nonVisualDrawingPropertiesExtensionList38);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties19 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks19 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties19.Append(pictureLocks19);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties48 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties19.Append(nonVisualDrawingProperties48);
            nonVisualPictureProperties19.Append(nonVisualPictureDrawingProperties19);
            nonVisualPictureProperties19.Append(applicationNonVisualDrawingProperties48);

            BlipFill blipFill19 = new BlipFill();
            A.Blip blip23 = new A.Blip() { Embed = "rId10" };

            A.Stretch stretch19 = new A.Stretch();
            A.FillRectangle fillRectangle19 = new A.FillRectangle();

            stretch19.Append(fillRectangle19);

            blipFill19.Append(blip23);
            blipFill19.Append(stretch19);

            ShapeProperties shapeProperties37 = new ShapeProperties();

            A.Transform2D transform2D36 = new A.Transform2D();
            A.Offset offset48 = new A.Offset() { X = 10361522L, Y = 2640488L };
            A.Extents extents48 = new A.Extents() { Cx = 869179L, Cy = 274285L };

            transform2D36.Append(offset48);
            transform2D36.Append(extents48);

            A.PresetGeometry presetGeometry36 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList36 = new A.AdjustValueList();

            presetGeometry36.Append(adjustValueList36);

            shapeProperties37.Append(transform2D36);
            shapeProperties37.Append(presetGeometry36);

            picture19.Append(nonVisualPictureProperties19);
            picture19.Append(blipFill19);
            picture19.Append(shapeProperties37);

            Picture picture20 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties20 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties49 = new NonVisualDrawingProperties() { Id = (UInt32Value)53U, Name = "Image 24", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList39 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension39 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement61 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{4E7EAFE1-51E7-72EF-75F6-C95A0FD3A494}\" />");

            // nonVisualDrawingPropertiesExtension39.Append(openXmlUnknownElement61);

            nonVisualDrawingPropertiesExtensionList39.Append(nonVisualDrawingPropertiesExtension39);

            nonVisualDrawingProperties49.Append(nonVisualDrawingPropertiesExtensionList39);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties20 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks20 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties20.Append(pictureLocks20);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties49 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties20.Append(nonVisualDrawingProperties49);
            nonVisualPictureProperties20.Append(nonVisualPictureDrawingProperties20);
            nonVisualPictureProperties20.Append(applicationNonVisualDrawingProperties49);

            BlipFill blipFill20 = new BlipFill();

            A.Blip blip24 = new A.Blip() { Embed = "rId11" };

            A.BlipExtensionList blipExtensionList13 = new A.BlipExtensionList();

            A.BlipExtension blipExtension13 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement62 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId12\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension13.Append(openXmlUnknownElement62);

            blipExtensionList13.Append(blipExtension13);

            blip24.Append(blipExtensionList13);

            A.Stretch stretch20 = new A.Stretch();
            A.FillRectangle fillRectangle20 = new A.FillRectangle();

            stretch20.Append(fillRectangle20);

            blipFill20.Append(blip24);
            blipFill20.Append(stretch20);

            ShapeProperties shapeProperties38 = new ShapeProperties();

            A.Transform2D transform2D37 = new A.Transform2D();
            A.Offset offset49 = new A.Offset() { X = 10453015L, Y = 2709059L };
            A.Extents extents49 = new A.Extents() { Cx = 137250L, Cy = 137157L };

            transform2D37.Append(offset49);
            transform2D37.Append(extents49);

            A.PresetGeometry presetGeometry37 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList37 = new A.AdjustValueList();

            presetGeometry37.Append(adjustValueList37);

            shapeProperties38.Append(transform2D37);
            shapeProperties38.Append(presetGeometry37);

            picture20.Append(nonVisualPictureProperties20);
            picture20.Append(blipFill20);
            picture20.Append(shapeProperties38);

            Shape shape18 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties18 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties50 = new NonVisualDrawingProperties() { Id = (UInt32Value)54U, Name = "Text 48" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList40 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension40 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement63 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{E3EBADEA-0378-BF13-FBD3-75904BEA2939}\" />");

            // nonVisualDrawingPropertiesExtension40.Append(openXmlUnknownElement63);

            nonVisualDrawingPropertiesExtensionList40.Append(nonVisualDrawingPropertiesExtension40);

            nonVisualDrawingProperties50.Append(nonVisualDrawingPropertiesExtensionList40);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties18 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties50 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties18.Append(nonVisualDrawingProperties50);
            nonVisualShapeProperties18.Append(nonVisualShapeDrawingProperties18);
            nonVisualShapeProperties18.Append(applicationNonVisualDrawingProperties50);

            ShapeProperties shapeProperties39 = new ShapeProperties();

            A.Transform2D transform2D38 = new A.Transform2D();
            A.Offset offset50 = new A.Offset() { X = 10658873L, Y = 2651915L };
            A.Extents extents50 = new A.Extents() { Cx = 526081L, Cy = 228571L };

            transform2D38.Append(offset50);
            transform2D38.Append(extents50);

            A.PresetGeometry presetGeometry38 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList38 = new A.AdjustValueList();

            presetGeometry38.Append(adjustValueList38);
            A.NoFill noFill153 = new A.NoFill();
            A.Outline outline30 = new A.Outline();

            shapeProperties39.Append(transform2D38);
            shapeProperties39.Append(presetGeometry38);
            shapeProperties39.Append(noFill153);
            shapeProperties39.Append(outline30);

            TextBody textBody39 = new TextBody();
            A.BodyProperties bodyProperties40 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle40 = new A.ListStyle();

            A.Paragraph paragraph43 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties34 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing27 = new A.LineSpacing();
            A.SpacingPercent spacingPercent27 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing27.Append(spacingPercent27);

            paragraphProperties34.Append(lineSpacing27);

            A.Run run28 = new A.Run();

            A.RunProperties runProperties30 = new A.RunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            A.SolidFill solidFill79 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex34 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha13 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex34.Append(alpha13);

            solidFill79.Append(rgbColorModelHex34);
            A.LatinFont latinFont59 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont49 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont59 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties30.Append(solidFill79);
            runProperties30.Append(latinFont59);
            runProperties30.Append(eastAsianFont49);
            runProperties30.Append(complexScriptFont59);
            A.Text text30 = new A.Text();

            // ********************************* Custom Code Start *********************************
            if (goalsArrayForSlide.Count > 5)
                text30.Text = goalsArrayForSlide.ElementAt(5).target.ToString() ?? "";
            else
                text30.Text = "NO DATA";
            // ********************************* Custom Code End *********************************


            run28.Append(runProperties30);
            run28.Append(text30);
            A.EndParagraphRunProperties endParagraphRunProperties33 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            paragraph43.Append(paragraphProperties34);
            paragraph43.Append(run28);
            paragraph43.Append(endParagraphRunProperties33);

            textBody39.Append(bodyProperties40);
            textBody39.Append(listStyle40);
            textBody39.Append(paragraph43);

            shape18.Append(nonVisualShapeProperties18);
            shape18.Append(shapeProperties39);
            shape18.Append(textBody39);

            groupShape9.Append(nonVisualGroupShapeProperties11);
            groupShape9.Append(groupShapeProperties11);
            groupShape9.Append(picture19);
            groupShape9.Append(picture20);
            groupShape9.Append(shape18);

            GroupShape groupShape10 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties12 = new NonVisualGroupShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties51 = new NonVisualDrawingProperties() { Id = (UInt32Value)31U, Name = "Group 30" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList41 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension41 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement64 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{EF434CEB-9403-0E42-6476-6C92A577B359}\" />");

            // nonVisualDrawingPropertiesExtension41.Append(openXmlUnknownElement64);

            nonVisualDrawingPropertiesExtensionList41.Append(nonVisualDrawingPropertiesExtension41);

            nonVisualDrawingProperties51.Append(nonVisualDrawingPropertiesExtensionList41);
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties12 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties51 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties12.Append(nonVisualDrawingProperties51);
            nonVisualGroupShapeProperties12.Append(nonVisualGroupShapeDrawingProperties12);
            nonVisualGroupShapeProperties12.Append(applicationNonVisualDrawingProperties51);

            GroupShapeProperties groupShapeProperties12 = new GroupShapeProperties();

            A.TransformGroup transformGroup12 = new A.TransformGroup();
            A.Offset offset51 = new A.Offset() { X = 10303409L, Y = 4786423L };
            A.Extents extents51 = new A.Extents() { Cx = 869179L, Cy = 274285L };
            A.ChildOffset childOffset12 = new A.ChildOffset() { X = 10361522L, Y = 2640488L };
            A.ChildExtents childExtents12 = new A.ChildExtents() { Cx = 869179L, Cy = 274285L };

            transformGroup12.Append(offset51);
            transformGroup12.Append(extents51);
            transformGroup12.Append(childOffset12);
            transformGroup12.Append(childExtents12);

            groupShapeProperties12.Append(transformGroup12);

            Picture picture21 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties21 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties52 = new NonVisualDrawingProperties() { Id = (UInt32Value)32U, Name = "Image 22", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList42 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension42 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement65 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{93337759-80A2-C2C6-4195-AECBF5ED8BC4}\" />");

            // nonVisualDrawingPropertiesExtension42.Append(openXmlUnknownElement65);

            nonVisualDrawingPropertiesExtensionList42.Append(nonVisualDrawingPropertiesExtension42);

            nonVisualDrawingProperties52.Append(nonVisualDrawingPropertiesExtensionList42);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties21 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks21 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties21.Append(pictureLocks21);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties52 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties21.Append(nonVisualDrawingProperties52);
            nonVisualPictureProperties21.Append(nonVisualPictureDrawingProperties21);
            nonVisualPictureProperties21.Append(applicationNonVisualDrawingProperties52);

            BlipFill blipFill21 = new BlipFill();
            A.Blip blip25 = new A.Blip() { Embed = "rId10" };

            A.Stretch stretch21 = new A.Stretch();
            A.FillRectangle fillRectangle21 = new A.FillRectangle();

            stretch21.Append(fillRectangle21);

            blipFill21.Append(blip25);
            blipFill21.Append(stretch21);

            ShapeProperties shapeProperties40 = new ShapeProperties();

            A.Transform2D transform2D39 = new A.Transform2D();
            A.Offset offset52 = new A.Offset() { X = 10361522L, Y = 2640488L };
            A.Extents extents52 = new A.Extents() { Cx = 869179L, Cy = 274285L };

            transform2D39.Append(offset52);
            transform2D39.Append(extents52);

            A.PresetGeometry presetGeometry39 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList39 = new A.AdjustValueList();

            presetGeometry39.Append(adjustValueList39);

            shapeProperties40.Append(transform2D39);
            shapeProperties40.Append(presetGeometry39);

            picture21.Append(nonVisualPictureProperties21);
            picture21.Append(blipFill21);
            picture21.Append(shapeProperties40);

            Picture picture22 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties22 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties53 = new NonVisualDrawingProperties() { Id = (UInt32Value)33U, Name = "Image 24", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList43 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension43 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement66 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{45D4776C-5AA0-2B14-591B-5262AEF375D4}\" />");

            // nonVisualDrawingPropertiesExtension43.Append(openXmlUnknownElement66);

            nonVisualDrawingPropertiesExtensionList43.Append(nonVisualDrawingPropertiesExtension43);

            nonVisualDrawingProperties53.Append(nonVisualDrawingPropertiesExtensionList43);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties22 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks22 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties22.Append(pictureLocks22);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties53 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties22.Append(nonVisualDrawingProperties53);
            nonVisualPictureProperties22.Append(nonVisualPictureDrawingProperties22);
            nonVisualPictureProperties22.Append(applicationNonVisualDrawingProperties53);

            BlipFill blipFill22 = new BlipFill();

            A.Blip blip26 = new A.Blip() { Embed = "rId11" };

            A.BlipExtensionList blipExtensionList14 = new A.BlipExtensionList();

            A.BlipExtension blipExtension14 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement67 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId12\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension14.Append(openXmlUnknownElement67);

            blipExtensionList14.Append(blipExtension14);

            blip26.Append(blipExtensionList14);

            A.Stretch stretch22 = new A.Stretch();
            A.FillRectangle fillRectangle22 = new A.FillRectangle();

            stretch22.Append(fillRectangle22);

            blipFill22.Append(blip26);
            blipFill22.Append(stretch22);

            ShapeProperties shapeProperties41 = new ShapeProperties();

            A.Transform2D transform2D40 = new A.Transform2D();
            A.Offset offset53 = new A.Offset() { X = 10453015L, Y = 2709059L };
            A.Extents extents53 = new A.Extents() { Cx = 137250L, Cy = 137157L };

            transform2D40.Append(offset53);
            transform2D40.Append(extents53);

            A.PresetGeometry presetGeometry40 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList40 = new A.AdjustValueList();

            presetGeometry40.Append(adjustValueList40);

            shapeProperties41.Append(transform2D40);
            shapeProperties41.Append(presetGeometry40);

            picture22.Append(nonVisualPictureProperties22);
            picture22.Append(blipFill22);
            picture22.Append(shapeProperties41);

            Shape shape19 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties19 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties54 = new NonVisualDrawingProperties() { Id = (UInt32Value)34U, Name = "Text 48" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList44 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension44 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement68 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{4287DBDB-975C-C67D-FEA6-D73C39A82935}\" />");

            // nonVisualDrawingPropertiesExtension44.Append(openXmlUnknownElement68);

            nonVisualDrawingPropertiesExtensionList44.Append(nonVisualDrawingPropertiesExtension44);

            nonVisualDrawingProperties54.Append(nonVisualDrawingPropertiesExtensionList44);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties19 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties54 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties19.Append(nonVisualDrawingProperties54);
            nonVisualShapeProperties19.Append(nonVisualShapeDrawingProperties19);
            nonVisualShapeProperties19.Append(applicationNonVisualDrawingProperties54);

            ShapeProperties shapeProperties42 = new ShapeProperties();

            A.Transform2D transform2D41 = new A.Transform2D();
            A.Offset offset54 = new A.Offset() { X = 10658873L, Y = 2651915L };
            A.Extents extents54 = new A.Extents() { Cx = 526081L, Cy = 228571L };

            transform2D41.Append(offset54);
            transform2D41.Append(extents54);

            A.PresetGeometry presetGeometry41 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList41 = new A.AdjustValueList();

            presetGeometry41.Append(adjustValueList41);
            A.NoFill noFill154 = new A.NoFill();
            A.Outline outline31 = new A.Outline();

            shapeProperties42.Append(transform2D41);
            shapeProperties42.Append(presetGeometry41);
            shapeProperties42.Append(noFill154);
            shapeProperties42.Append(outline31);

            TextBody textBody40 = new TextBody();
            A.BodyProperties bodyProperties41 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle41 = new A.ListStyle();

            A.Paragraph paragraph44 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties35 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing28 = new A.LineSpacing();
            A.SpacingPercent spacingPercent28 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing28.Append(spacingPercent28);

            paragraphProperties35.Append(lineSpacing28);

            A.Run run29 = new A.Run();

            A.RunProperties runProperties31 = new A.RunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            A.SolidFill solidFill80 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex35 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha14 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex35.Append(alpha14);

            solidFill80.Append(rgbColorModelHex35);
            A.LatinFont latinFont60 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont50 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont60 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties31.Append(solidFill80);
            runProperties31.Append(latinFont60);
            runProperties31.Append(eastAsianFont50);
            runProperties31.Append(complexScriptFont60);
            A.Text text31 = new A.Text();

            // ********************************* Custom Code Start *********************************
            if (goalsArrayForSlide.Count > 4)
                text31.Text = goalsArrayForSlide.ElementAt(4).target.ToString() ?? "";
            else
                text31.Text = "NO DATA";
            // ********************************* Custom Code End *********************************

            run29.Append(runProperties31);
            run29.Append(text31);
            A.EndParagraphRunProperties endParagraphRunProperties34 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            paragraph44.Append(paragraphProperties35);
            paragraph44.Append(run29);
            paragraph44.Append(endParagraphRunProperties34);

            textBody40.Append(bodyProperties41);
            textBody40.Append(listStyle41);
            textBody40.Append(paragraph44);

            shape19.Append(nonVisualShapeProperties19);
            shape19.Append(shapeProperties42);
            shape19.Append(textBody40);

            groupShape10.Append(nonVisualGroupShapeProperties12);
            groupShape10.Append(groupShapeProperties12);
            groupShape10.Append(picture21);
            groupShape10.Append(picture22);
            groupShape10.Append(shape19);

            Picture picture23 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties23 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties55 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Image 20", Description = " ", Hidden = true };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList45 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension45 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement69 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{2F6A9D2C-76F5-E8FE-40D3-E266E862A086}\" />");

            // nonVisualDrawingPropertiesExtension45.Append(openXmlUnknownElement69);

            nonVisualDrawingPropertiesExtensionList45.Append(nonVisualDrawingPropertiesExtension45);

            nonVisualDrawingProperties55.Append(nonVisualDrawingPropertiesExtensionList45);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties23 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks23 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties23.Append(pictureLocks23);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties55 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties23.Append(nonVisualDrawingProperties55);
            nonVisualPictureProperties23.Append(nonVisualPictureDrawingProperties23);
            nonVisualPictureProperties23.Append(applicationNonVisualDrawingProperties55);

            BlipFill blipFill23 = new BlipFill();
            A.Blip blip27 = new A.Blip() { Embed = "rId7" };

            A.Stretch stretch23 = new A.Stretch();
            A.FillRectangle fillRectangle23 = new A.FillRectangle();

            stretch23.Append(fillRectangle23);

            blipFill23.Append(blip27);
            blipFill23.Append(stretch23);

            ShapeProperties shapeProperties43 = new ShapeProperties();

            A.Transform2D transform2D42 = new A.Transform2D();
            A.Offset offset55 = new A.Offset() { X = 6592059L, Y = 1775663L };
            A.Extents extents55 = new A.Extents() { Cx = 1669738L, Cy = 274285L };

            transform2D42.Append(offset55);
            transform2D42.Append(extents55);

            A.PresetGeometry presetGeometry42 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList42 = new A.AdjustValueList();

            presetGeometry42.Append(adjustValueList42);

            shapeProperties43.Append(transform2D42);
            shapeProperties43.Append(presetGeometry42);

            picture23.Append(nonVisualPictureProperties23);
            picture23.Append(blipFill23);
            picture23.Append(shapeProperties43);

            GroupShape groupShape11 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties13 = new NonVisualGroupShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties56 = new NonVisualDrawingProperties() { Id = (UInt32Value)55U, Name = "Group 54" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList46 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension46 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement70 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{784CD544-CAE4-7C48-9305-766DD73A1AD1}\" />");

            // nonVisualDrawingPropertiesExtension46.Append(openXmlUnknownElement70);

            nonVisualDrawingPropertiesExtensionList46.Append(nonVisualDrawingPropertiesExtension46);

            nonVisualDrawingProperties56.Append(nonVisualDrawingPropertiesExtensionList46);
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties13 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties56 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties13.Append(nonVisualDrawingProperties56);
            nonVisualGroupShapeProperties13.Append(nonVisualGroupShapeDrawingProperties13);
            nonVisualGroupShapeProperties13.Append(applicationNonVisualDrawingProperties56);

            GroupShapeProperties groupShapeProperties13 = new GroupShapeProperties();

            A.TransformGroup transformGroup13 = new A.TransformGroup();
            A.Offset offset56 = new A.Offset() { X = 8355350L, Y = 1785546L };
            A.Extents extents56 = new A.Extents() { Cx = 1641229L, Cy = 290937L };
            A.ChildOffset childOffset13 = new A.ChildOffset() { X = 8344464L, Y = 1785546L };
            A.ChildExtents childExtents13 = new A.ChildExtents() { Cx = 1641229L, Cy = 290937L };

            transformGroup13.Append(offset56);
            transformGroup13.Append(extents56);
            transformGroup13.Append(childOffset13);
            transformGroup13.Append(childExtents13);

            groupShapeProperties13.Append(transformGroup13);

            Shape shape20 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties20 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties57 = new NonVisualDrawingProperties() { Id = (UInt32Value)35U, Name = "Rectangle: Rounded Corners 34" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList47 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension47 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement71 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{CA89EDB4-88F4-FE63-3EDC-F1AC979C1E7F}\" />");

            // nonVisualDrawingPropertiesExtension47.Append(openXmlUnknownElement71);

            nonVisualDrawingPropertiesExtensionList47.Append(nonVisualDrawingPropertiesExtension47);

            nonVisualDrawingProperties57.Append(nonVisualDrawingPropertiesExtensionList47);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties20 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties57 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties20.Append(nonVisualDrawingProperties57);
            nonVisualShapeProperties20.Append(nonVisualShapeDrawingProperties20);
            nonVisualShapeProperties20.Append(applicationNonVisualDrawingProperties57);

            ShapeProperties shapeProperties44 = new ShapeProperties();

            A.Transform2D transform2D43 = new A.Transform2D();
            A.Offset offset57 = new A.Offset() { X = 8344464L, Y = 1785546L };
            A.Extents extents57 = new A.Extents() { Cx = 1641229L, Cy = 290937L };

            transform2D43.Append(offset57);
            transform2D43.Append(extents57);

            A.PresetGeometry presetGeometry43 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.FlowChartTerminator };
            A.AdjustValueList adjustValueList43 = new A.AdjustValueList();

            presetGeometry43.Append(adjustValueList43);

            A.SolidFill solidFill81 = new A.SolidFill();
            A.SchemeColor schemeColor63 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill81.Append(schemeColor63);

            A.Outline outline32 = new A.Outline();
            A.NoFill noFill155 = new A.NoFill();

            outline32.Append(noFill155);

            shapeProperties44.Append(transform2D43);
            shapeProperties44.Append(presetGeometry43);
            shapeProperties44.Append(solidFill81);
            shapeProperties44.Append(outline32);

            ShapeStyle shapeStyle2 = new ShapeStyle();

            A.LineReference lineReference2 = new A.LineReference() { Index = (UInt32Value)2U };

            A.SchemeColor schemeColor64 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.Shade shade6 = new A.Shade() { Val = 15000 };

            schemeColor64.Append(shade6);

            lineReference2.Append(schemeColor64);

            A.FillReference fillReference2 = new A.FillReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor65 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference2.Append(schemeColor65);

            A.EffectReference effectReference2 = new A.EffectReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor66 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference2.Append(schemeColor66);

            A.FontReference fontReference2 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor67 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            fontReference2.Append(schemeColor67);

            shapeStyle2.Append(lineReference2);
            shapeStyle2.Append(fillReference2);
            shapeStyle2.Append(effectReference2);
            shapeStyle2.Append(fontReference2);

            TextBody textBody41 = new TextBody();
            A.BodyProperties bodyProperties42 = new A.BodyProperties() { RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle42 = new A.ListStyle();

            A.Paragraph paragraph45 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties36 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };
            A.EndParagraphRunProperties endParagraphRunProperties35 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph45.Append(paragraphProperties36);
            paragraph45.Append(endParagraphRunProperties35);

            textBody41.Append(bodyProperties42);
            textBody41.Append(listStyle42);
            textBody41.Append(paragraph45);

            shape20.Append(nonVisualShapeProperties20);
            shape20.Append(shapeProperties44);
            shape20.Append(shapeStyle2);
            shape20.Append(textBody41);

            Picture picture24 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties24 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties58 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Image 21", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList48 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension48 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement72 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{2C659360-BFDF-838B-45CA-9C706FF4435A}\" />");

            // nonVisualDrawingPropertiesExtension48.Append(openXmlUnknownElement72);

            nonVisualDrawingPropertiesExtensionList48.Append(nonVisualDrawingPropertiesExtension48);

            nonVisualDrawingProperties58.Append(nonVisualDrawingPropertiesExtensionList48);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties24 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks24 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties24.Append(pictureLocks24);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties58 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties24.Append(nonVisualDrawingProperties58);
            nonVisualPictureProperties24.Append(nonVisualPictureDrawingProperties24);
            nonVisualPictureProperties24.Append(applicationNonVisualDrawingProperties58);

            BlipFill blipFill24 = new BlipFill();

            A.Blip blip28 = new A.Blip() { Embed = "rId8" };

            A.BlipExtensionList blipExtensionList15 = new A.BlipExtensionList();

            A.BlipExtension blipExtension15 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement73 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId9\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension15.Append(openXmlUnknownElement73);

            blipExtensionList15.Append(blipExtension15);

            blip28.Append(blipExtensionList15);

            A.Stretch stretch24 = new A.Stretch();
            A.FillRectangle fillRectangle24 = new A.FillRectangle();

            stretch24.Append(fillRectangle24);

            blipFill24.Append(blip28);
            blipFill24.Append(stretch24);

            ShapeProperties shapeProperties45 = new ShapeProperties();

            A.Transform2D transform2D44 = new A.Transform2D();
            A.Offset offset58 = new A.Offset() { X = 8359057L, Y = 1819396L };
            A.Extents extents58 = new A.Extents() { Cx = 228731L, Cy = 228571L };

            transform2D44.Append(offset58);
            transform2D44.Append(extents58);

            A.PresetGeometry presetGeometry44 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList44 = new A.AdjustValueList();

            presetGeometry44.Append(adjustValueList44);

            shapeProperties45.Append(transform2D44);
            shapeProperties45.Append(presetGeometry44);

            picture24.Append(nonVisualPictureProperties24);
            picture24.Append(blipFill24);
            picture24.Append(shapeProperties45);

            Shape shape21 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties21 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties59 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Text 46" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList49 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension49 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement74 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{557508EE-992E-BB96-0D9E-9C77A9057FE2}\" />");

            // nonVisualDrawingPropertiesExtension49.Append(openXmlUnknownElement74);

            nonVisualDrawingPropertiesExtensionList49.Append(nonVisualDrawingPropertiesExtension49);

            nonVisualDrawingProperties59.Append(nonVisualDrawingPropertiesExtensionList49);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties21 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties59 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties21.Append(nonVisualDrawingProperties59);
            nonVisualShapeProperties21.Append(nonVisualShapeDrawingProperties21);
            nonVisualShapeProperties21.Append(applicationNonVisualDrawingProperties59);

            ShapeProperties shapeProperties46 = new ShapeProperties();

            A.Transform2D transform2D45 = new A.Transform2D();
            A.Offset offset59 = new A.Offset() { X = 8679280L, Y = 1807967L };
            A.Extents extents59 = new A.Extents() { Cx = 1280895L, Cy = 228571L };

            transform2D45.Append(offset59);
            transform2D45.Append(extents59);

            A.PresetGeometry presetGeometry45 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList45 = new A.AdjustValueList();

            presetGeometry45.Append(adjustValueList45);
            A.NoFill noFill156 = new A.NoFill();
            A.Outline outline33 = new A.Outline();

            shapeProperties46.Append(transform2D45);
            shapeProperties46.Append(presetGeometry45);
            shapeProperties46.Append(noFill156);
            shapeProperties46.Append(outline33);

            TextBody textBody42 = new TextBody();
            A.BodyProperties bodyProperties43 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle43 = new A.ListStyle();

            A.Paragraph paragraph46 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties37 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing29 = new A.LineSpacing();
            A.SpacingPercent spacingPercent29 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing29.Append(spacingPercent29);

            paragraphProperties37.Append(lineSpacing29);

            A.Run run30 = new A.Run();

            A.RunProperties runProperties32 = new A.RunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            A.SolidFill solidFill82 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex36 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha15 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex36.Append(alpha15);

            solidFill82.Append(rgbColorModelHex36);
            A.LatinFont latinFont61 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont51 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont61 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties32.Append(solidFill82);
            runProperties32.Append(latinFont61);
            runProperties32.Append(eastAsianFont51);
            runProperties32.Append(complexScriptFont61);
            A.Text text32 = new A.Text();
            text32.Text = goalsArrayForSlide.ElementAt(0).goalOwner.ToString() ?? "";

            run30.Append(runProperties32);
            run30.Append(text32);
            A.EndParagraphRunProperties endParagraphRunProperties36 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            paragraph46.Append(paragraphProperties37);
            paragraph46.Append(run30);
            paragraph46.Append(endParagraphRunProperties36);

            textBody42.Append(bodyProperties43);
            textBody42.Append(listStyle43);
            textBody42.Append(paragraph46);

            shape21.Append(nonVisualShapeProperties21);
            shape21.Append(shapeProperties46);
            shape21.Append(textBody42);

            groupShape11.Append(nonVisualGroupShapeProperties13);
            groupShape11.Append(groupShapeProperties13);
            groupShape11.Append(shape20);
            groupShape11.Append(picture24);
            groupShape11.Append(shape21);

            GroupShape groupShape12 = new GroupShape();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties14 = new NonVisualGroupShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties60 = new NonVisualDrawingProperties() { Id = (UInt32Value)56U, Name = "Group 55" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList50 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension50 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement75 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{13DB0D39-AB65-FE1D-9EC1-CA3E3E77515C}\" />");

            // nonVisualDrawingPropertiesExtension50.Append(openXmlUnknownElement75);

            nonVisualDrawingPropertiesExtensionList50.Append(nonVisualDrawingPropertiesExtension50);

            nonVisualDrawingProperties60.Append(nonVisualDrawingPropertiesExtensionList50);
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties14 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties60 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties14.Append(nonVisualDrawingProperties60);
            nonVisualGroupShapeProperties14.Append(nonVisualGroupShapeDrawingProperties14);
            nonVisualGroupShapeProperties14.Append(applicationNonVisualDrawingProperties60);

            GroupShapeProperties groupShapeProperties14 = new GroupShapeProperties();

            A.TransformGroup transformGroup14 = new A.TransformGroup();
            A.Offset offset60 = new A.Offset() { X = 8355350L, Y = 4833636L };
            A.Extents extents60 = new A.Extents() { Cx = 1641229L, Cy = 290937L };
            A.ChildOffset childOffset14 = new A.ChildOffset() { X = 8344464L, Y = 1785546L };
            A.ChildExtents childExtents14 = new A.ChildExtents() { Cx = 1641229L, Cy = 290937L };

            transformGroup14.Append(offset60);
            transformGroup14.Append(extents60);
            transformGroup14.Append(childOffset14);
            transformGroup14.Append(childExtents14);

            groupShapeProperties14.Append(transformGroup14);

            Shape shape22 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties22 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties61 = new NonVisualDrawingProperties() { Id = (UInt32Value)57U, Name = "Rectangle: Rounded Corners 34" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList51 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension51 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement76 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{9FE627CC-941E-C4BF-3E4E-B467CF96B43D}\" />");

            // nonVisualDrawingPropertiesExtension51.Append(openXmlUnknownElement76);

            nonVisualDrawingPropertiesExtensionList51.Append(nonVisualDrawingPropertiesExtension51);

            nonVisualDrawingProperties61.Append(nonVisualDrawingPropertiesExtensionList51);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties22 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties61 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties22.Append(nonVisualDrawingProperties61);
            nonVisualShapeProperties22.Append(nonVisualShapeDrawingProperties22);
            nonVisualShapeProperties22.Append(applicationNonVisualDrawingProperties61);

            ShapeProperties shapeProperties47 = new ShapeProperties();

            A.Transform2D transform2D46 = new A.Transform2D();
            A.Offset offset61 = new A.Offset() { X = 8344464L, Y = 1785546L };
            A.Extents extents61 = new A.Extents() { Cx = 1641229L, Cy = 290937L };

            transform2D46.Append(offset61);
            transform2D46.Append(extents61);

            A.PresetGeometry presetGeometry46 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.FlowChartTerminator };
            A.AdjustValueList adjustValueList46 = new A.AdjustValueList();

            presetGeometry46.Append(adjustValueList46);

            A.SolidFill solidFill83 = new A.SolidFill();
            A.SchemeColor schemeColor68 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            solidFill83.Append(schemeColor68);

            A.Outline outline34 = new A.Outline();
            A.NoFill noFill157 = new A.NoFill();

            outline34.Append(noFill157);

            shapeProperties47.Append(transform2D46);
            shapeProperties47.Append(presetGeometry46);
            shapeProperties47.Append(solidFill83);
            shapeProperties47.Append(outline34);

            ShapeStyle shapeStyle3 = new ShapeStyle();

            A.LineReference lineReference3 = new A.LineReference() { Index = (UInt32Value)2U };

            A.SchemeColor schemeColor69 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
            A.Shade shade7 = new A.Shade() { Val = 15000 };

            schemeColor69.Append(shade7);

            lineReference3.Append(schemeColor69);

            A.FillReference fillReference3 = new A.FillReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor70 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference3.Append(schemeColor70);

            A.EffectReference effectReference3 = new A.EffectReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor71 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference3.Append(schemeColor71);

            A.FontReference fontReference3 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor72 = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

            fontReference3.Append(schemeColor72);

            shapeStyle3.Append(lineReference3);
            shapeStyle3.Append(fillReference3);
            shapeStyle3.Append(effectReference3);
            shapeStyle3.Append(fontReference3);

            TextBody textBody43 = new TextBody();
            A.BodyProperties bodyProperties44 = new A.BodyProperties() { RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle44 = new A.ListStyle();

            A.Paragraph paragraph47 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties38 = new A.ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };
            A.EndParagraphRunProperties endParagraphRunProperties37 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph47.Append(paragraphProperties38);
            paragraph47.Append(endParagraphRunProperties37);

            textBody43.Append(bodyProperties44);
            textBody43.Append(listStyle44);
            textBody43.Append(paragraph47);

            shape22.Append(nonVisualShapeProperties22);
            shape22.Append(shapeProperties47);
            shape22.Append(shapeStyle3);
            shape22.Append(textBody43);

            Picture picture25 = new Picture();

            NonVisualPictureProperties nonVisualPictureProperties25 = new NonVisualPictureProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties62 = new NonVisualDrawingProperties() { Id = (UInt32Value)58U, Name = "Image 21", Description = " " };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList52 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension52 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement77 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{7CCDEB59-99FA-66F7-8515-03472E540D3C}\" />");

            // nonVisualDrawingPropertiesExtension52.Append(openXmlUnknownElement77);

            nonVisualDrawingPropertiesExtensionList52.Append(nonVisualDrawingPropertiesExtension52);

            nonVisualDrawingProperties62.Append(nonVisualDrawingPropertiesExtensionList52);

            NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties25 = new NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks25 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties25.Append(pictureLocks25);
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties62 = new ApplicationNonVisualDrawingProperties();

            nonVisualPictureProperties25.Append(nonVisualDrawingProperties62);
            nonVisualPictureProperties25.Append(nonVisualPictureDrawingProperties25);
            nonVisualPictureProperties25.Append(applicationNonVisualDrawingProperties62);

            BlipFill blipFill25 = new BlipFill();

            A.Blip blip29 = new A.Blip() { Embed = "rId8" };

            A.BlipExtensionList blipExtensionList16 = new A.BlipExtensionList();

            A.BlipExtension blipExtension16 = new A.BlipExtension() { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement78 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId9\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" />");

            // blipExtension16.Append(openXmlUnknownElement78);

            blipExtensionList16.Append(blipExtension16);

            blip29.Append(blipExtensionList16);

            A.Stretch stretch25 = new A.Stretch();
            A.FillRectangle fillRectangle25 = new A.FillRectangle();

            stretch25.Append(fillRectangle25);

            blipFill25.Append(blip29);
            blipFill25.Append(stretch25);

            ShapeProperties shapeProperties48 = new ShapeProperties();

            A.Transform2D transform2D47 = new A.Transform2D();
            A.Offset offset62 = new A.Offset() { X = 8359057L, Y = 1819396L };
            A.Extents extents62 = new A.Extents() { Cx = 228731L, Cy = 228571L };

            transform2D47.Append(offset62);
            transform2D47.Append(extents62);

            A.PresetGeometry presetGeometry47 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList47 = new A.AdjustValueList();

            presetGeometry47.Append(adjustValueList47);

            shapeProperties48.Append(transform2D47);
            shapeProperties48.Append(presetGeometry47);

            picture25.Append(nonVisualPictureProperties25);
            picture25.Append(blipFill25);
            picture25.Append(shapeProperties48);

            Shape shape23 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties23 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties63 = new NonVisualDrawingProperties() { Id = (UInt32Value)59U, Name = "Text 46" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList53 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension53 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement79 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{C1A2A4B9-2C91-5CDA-4882-B0470AE145A6}\" />");

            // nonVisualDrawingPropertiesExtension53.Append(openXmlUnknownElement79);

            nonVisualDrawingPropertiesExtensionList53.Append(nonVisualDrawingPropertiesExtension53);

            nonVisualDrawingProperties63.Append(nonVisualDrawingPropertiesExtensionList53);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties23 = new NonVisualShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties63 = new ApplicationNonVisualDrawingProperties();

            nonVisualShapeProperties23.Append(nonVisualDrawingProperties63);
            nonVisualShapeProperties23.Append(nonVisualShapeDrawingProperties23);
            nonVisualShapeProperties23.Append(applicationNonVisualDrawingProperties63);

            ShapeProperties shapeProperties49 = new ShapeProperties();

            A.Transform2D transform2D48 = new A.Transform2D();
            A.Offset offset63 = new A.Offset() { X = 8679280L, Y = 1807967L };
            A.Extents extents63 = new A.Extents() { Cx = 1280895L, Cy = 228571L };

            transform2D48.Append(offset63);
            transform2D48.Append(extents63);

            A.PresetGeometry presetGeometry48 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList48 = new A.AdjustValueList();

            presetGeometry48.Append(adjustValueList48);
            A.NoFill noFill158 = new A.NoFill();
            A.Outline outline35 = new A.Outline();

            shapeProperties49.Append(transform2D48);
            shapeProperties49.Append(presetGeometry48);
            shapeProperties49.Append(noFill158);
            shapeProperties49.Append(outline35);

            TextBody textBody44 = new TextBody();
            A.BodyProperties bodyProperties45 = new A.BodyProperties() { Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.ListStyle listStyle45 = new A.ListStyle();

            A.Paragraph paragraph48 = new A.Paragraph();

            A.ParagraphProperties paragraphProperties39 = new A.ParagraphProperties();

            A.LineSpacing lineSpacing30 = new A.LineSpacing();
            A.SpacingPercent spacingPercent30 = new A.SpacingPercent() { Val = 132505 };

            lineSpacing30.Append(spacingPercent30);

            paragraphProperties39.Append(lineSpacing30);

            A.Run run31 = new A.Run();

            A.RunProperties runProperties33 = new A.RunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            A.SolidFill solidFill84 = new A.SolidFill();

            A.RgbColorModelHex rgbColorModelHex37 = new A.RgbColorModelHex() { Val = "424242" };
            A.Alpha alpha16 = new A.Alpha() { Val = 100000 };

            rgbColorModelHex37.Append(alpha16);

            solidFill84.Append(rgbColorModelHex37);
            A.LatinFont latinFont62 = new A.LatinFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = 0 };
            A.EastAsianFont eastAsianFont52 = new A.EastAsianFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -122 };
            A.ComplexScriptFont complexScriptFont62 = new A.ComplexScriptFont() { Typeface = "Segoe UI Regular", PitchFamily = 34, CharacterSet = -120 };

            runProperties33.Append(solidFill84);
            runProperties33.Append(latinFont62);
            runProperties33.Append(eastAsianFont52);
            runProperties33.Append(complexScriptFont62);
            A.Text text33 = new A.Text();

            // *********************************** Custom Code Start ***********************************
            if (goalsArrayForSlide.Count > 4)
                text33.Text = goalsArrayForSlide.ElementAt(4).goalOwner.ToString() ?? "";
            else
                text33.Text = "NO DATA";
            // *********************************** Custom Code End ***********************************

            run31.Append(runProperties33);
            run31.Append(text33);
            A.EndParagraphRunProperties endParagraphRunProperties38 = new A.EndParagraphRunProperties() { Language = "en-US", FontSize = 1080, Dirty = false };

            paragraph48.Append(paragraphProperties39);
            paragraph48.Append(run31);
            paragraph48.Append(endParagraphRunProperties38);

            textBody44.Append(bodyProperties45);
            textBody44.Append(listStyle45);
            textBody44.Append(paragraph48);

            shape23.Append(nonVisualShapeProperties23);
            shape23.Append(shapeProperties49);
            shape23.Append(textBody44);

            groupShape12.Append(nonVisualGroupShapeProperties14);
            groupShape12.Append(groupShapeProperties14);
            groupShape12.Append(shape22);
            groupShape12.Append(picture25);
            groupShape12.Append(shape23);

            shapeTree2.Append(nonVisualGroupShapeProperties2);
            shapeTree2.Append(groupShapeProperties2);
            shapeTree2.Append(picture1);
            shapeTree2.Append(shape7);
            shapeTree2.Append(shape8);
            shapeTree2.Append(shape9);
            shapeTree2.Append(picture2);
            shapeTree2.Append(graphicFrame1);

            // // owner-1
            // shapeTree2.Append(groupShape11);

            // // owner-2
            // shapeTree2.Append(groupShape1);

            // // owner-3
            // shapeTree2.Append(groupShape2);

            // // owner-4
            // shapeTree2.Append(groupShape3);

            // // owner-5
            // shapeTree2.Append(groupShape12);

            // // owner-6
            // shapeTree2.Append(groupShape4);

            // // target-1
            // shapeTree2.Append(groupShape5);

            // // target-2
            // shapeTree2.Append(groupShape6);

            // // target-3
            // shapeTree2.Append(groupShape7);

            // // target-4
            // shapeTree2.Append(groupShape8);

            // // target-5
            // shapeTree2.Append(groupShape10);

            // // target-6
            // shapeTree2.Append(groupShape9);


            // *********************************** Custom Code Start ***********************************
            // creating an array for the owners and targets

            GroupShape[] ownerGroupShape = [groupShape11, groupShape1, groupShape2, groupShape3, groupShape12, groupShape4];
            GroupShape[] targetGroupShape = [groupShape5, groupShape6, groupShape7, groupShape8, groupShape10, groupShape9];

            for (int i = 0; i < goalsArrayForSlide.Count; i++)
            {
                shapeTree2.Append(ownerGroupShape[i]);
                shapeTree2.Append(targetGroupShape[i]);
            }

            // *********************************** Custom Code End ***********************************



            shapeTree2.Append(picture23);



            CommonSlideDataExtensionList commonSlideDataExtensionList2 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension2 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId2 = new P14.CreationId() { Val = (UInt32Value)427100226U };
            creationId2.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension2.Append(creationId2);

            commonSlideDataExtensionList2.Append(commonSlideDataExtension2);

            commonSlideData2.Append(shapeTree2);
            commonSlideData2.Append(commonSlideDataExtensionList2);

            ColorMapOverride colorMapOverride1 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping1 = new A.MasterColorMapping();

            colorMapOverride1.Append(masterColorMapping1);

            slide1.Append(commonSlideData2);
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

        // Generates content of imagePart4.
        private void GenerateImagePart4Content(ImagePart imagePart4)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart4Data);
            imagePart4.FeedData(data);
            data.Close();
        }

        // Generates content of notesSlidePart1.
        private void GenerateNotesSlidePartContent(NotesSlidePart notesSlidePart1)
        {
            NotesSlide notesSlide1 = new NotesSlide();
            notesSlide1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            notesSlide1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            notesSlide1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData3 = new CommonSlideData();

            ShapeTree shapeTree3 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties15 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties64 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties15 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties64 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties15.Append(nonVisualDrawingProperties64);
            nonVisualGroupShapeProperties15.Append(nonVisualGroupShapeDrawingProperties15);
            nonVisualGroupShapeProperties15.Append(applicationNonVisualDrawingProperties64);

            GroupShapeProperties groupShapeProperties15 = new GroupShapeProperties();

            A.TransformGroup transformGroup15 = new A.TransformGroup();
            A.Offset offset64 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents64 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset15 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents15 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup15.Append(offset64);
            transformGroup15.Append(extents64);
            transformGroup15.Append(childOffset15);
            transformGroup15.Append(childExtents15);

            groupShapeProperties15.Append(transformGroup15);

            Shape shape24 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties24 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties65 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Slide Image Placeholder 1" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties24 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks7 = new A.ShapeLocks() { NoGrouping = true, NoRotation = true, NoChangeAspect = true };

            nonVisualShapeDrawingProperties24.Append(shapeLocks7);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties65 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape7 = new PlaceholderShape() { Type = PlaceholderValues.SlideImage };

            applicationNonVisualDrawingProperties65.Append(placeholderShape7);

            nonVisualShapeProperties24.Append(nonVisualDrawingProperties65);
            nonVisualShapeProperties24.Append(nonVisualShapeDrawingProperties24);
            nonVisualShapeProperties24.Append(applicationNonVisualDrawingProperties65);
            ShapeProperties shapeProperties50 = new ShapeProperties();

            shape24.Append(nonVisualShapeProperties24);
            shape24.Append(shapeProperties50);

            Shape shape25 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties25 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties66 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Notes Placeholder 2" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties25 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks8 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties25.Append(shapeLocks8);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties66 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape8 = new PlaceholderShape() { Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties66.Append(placeholderShape8);

            nonVisualShapeProperties25.Append(nonVisualDrawingProperties66);
            nonVisualShapeProperties25.Append(nonVisualShapeDrawingProperties25);
            nonVisualShapeProperties25.Append(applicationNonVisualDrawingProperties66);
            ShapeProperties shapeProperties51 = new ShapeProperties();

            TextBody textBody45 = new TextBody();
            A.BodyProperties bodyProperties46 = new A.BodyProperties();
            A.ListStyle listStyle46 = new A.ListStyle();

            A.Paragraph paragraph49 = new A.Paragraph();

            A.EndParagraphRunProperties endParagraphRunProperties39 = new A.EndParagraphRunProperties() { Language = "en-US", Bold = true, Dirty = false };

            A.SolidFill solidFill85 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex38 = new A.RgbColorModelHex() { Val = "C00000" };

            solidFill85.Append(rgbColorModelHex38);

            endParagraphRunProperties39.Append(solidFill85);

            paragraph49.Append(endParagraphRunProperties39);

            textBody45.Append(bodyProperties46);
            textBody45.Append(listStyle46);
            textBody45.Append(paragraph49);

            shape25.Append(nonVisualShapeProperties25);
            shape25.Append(shapeProperties51);
            shape25.Append(textBody45);

            Shape shape26 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties26 = new NonVisualShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties67 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Slide Number Placeholder 3" };

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties26 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks9 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties26.Append(shapeLocks9);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties67 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape9 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties67.Append(placeholderShape9);

            nonVisualShapeProperties26.Append(nonVisualDrawingProperties67);
            nonVisualShapeProperties26.Append(nonVisualShapeDrawingProperties26);
            nonVisualShapeProperties26.Append(applicationNonVisualDrawingProperties67);
            ShapeProperties shapeProperties52 = new ShapeProperties();

            TextBody textBody46 = new TextBody();
            A.BodyProperties bodyProperties47 = new A.BodyProperties();
            A.ListStyle listStyle47 = new A.ListStyle();

            A.Paragraph paragraph50 = new A.Paragraph();

            A.Field field3 = new A.Field() { Id = "{F7021451-1387-4CA6-816F-3879F97B5CBC}", Type = "slidenum" };
            A.RunProperties runProperties34 = new A.RunProperties() { Language = "en-US" };
            A.Text text34 = new A.Text();
            text34.Text = "1";

            field3.Append(runProperties34);
            field3.Append(text34);
            A.EndParagraphRunProperties endParagraphRunProperties40 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph50.Append(field3);
            paragraph50.Append(endParagraphRunProperties40);

            textBody46.Append(bodyProperties47);
            textBody46.Append(listStyle47);
            textBody46.Append(paragraph50);

            shape26.Append(nonVisualShapeProperties26);
            shape26.Append(shapeProperties52);
            shape26.Append(textBody46);

            shapeTree3.Append(nonVisualGroupShapeProperties15);
            shapeTree3.Append(groupShapeProperties15);
            shapeTree3.Append(shape24);
            shapeTree3.Append(shape25);
            shapeTree3.Append(shape26);

            CommonSlideDataExtensionList commonSlideDataExtensionList3 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension3 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId3 = new P14.CreationId() { Val = (UInt32Value)1797079604U };
            creationId3.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension3.Append(creationId3);

            commonSlideDataExtensionList3.Append(commonSlideDataExtension3);

            commonSlideData3.Append(shapeTree3);
            commonSlideData3.Append(commonSlideDataExtensionList3);

            ColorMapOverride colorMapOverride2 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping2 = new A.MasterColorMapping();

            colorMapOverride2.Append(masterColorMapping2);

            notesSlide1.Append(commonSlideData3);
            notesSlide1.Append(colorMapOverride2);

            notesSlidePart1.NotesSlide = notesSlide1;
        }

        // Generates content of slideLayoutPart1.
        private void GenerateSlideLayoutPart1Content(SlideLayoutPart slideLayoutPart1)
        {
            SlideLayout slideLayout1 = new SlideLayout();
            slideLayout1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData4 = new CommonSlideData() { Name = "DEFAULT" };

            Background background2 = new Background();

            BackgroundStyleReference backgroundStyleReference2 = new BackgroundStyleReference() { Index = (UInt32Value)1001U };
            A.SchemeColor schemeColor73 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            backgroundStyleReference2.Append(schemeColor73);

            background2.Append(backgroundStyleReference2);

            ShapeTree shapeTree4 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties16 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties68 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties16 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties68 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties16.Append(nonVisualDrawingProperties68);
            nonVisualGroupShapeProperties16.Append(nonVisualGroupShapeDrawingProperties16);
            nonVisualGroupShapeProperties16.Append(applicationNonVisualDrawingProperties68);

            GroupShapeProperties groupShapeProperties16 = new GroupShapeProperties();

            A.TransformGroup transformGroup16 = new A.TransformGroup();
            A.Offset offset65 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents65 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset16 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents16 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup16.Append(offset65);
            transformGroup16.Append(extents65);
            transformGroup16.Append(childOffset16);
            transformGroup16.Append(childExtents16);

            groupShapeProperties16.Append(transformGroup16);

            shapeTree4.Append(nonVisualGroupShapeProperties16);
            shapeTree4.Append(groupShapeProperties16);

            CommonSlideDataExtensionList commonSlideDataExtensionList4 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension4 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId4 = new P14.CreationId() { Val = (UInt32Value)1918912108U };
            creationId4.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension4.Append(creationId4);

            commonSlideDataExtensionList4.Append(commonSlideDataExtension4);

            commonSlideData4.Append(background2);
            commonSlideData4.Append(shapeTree4);
            commonSlideData4.Append(commonSlideDataExtensionList4);

            ColorMapOverride colorMapOverride3 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping3 = new A.MasterColorMapping();

            colorMapOverride3.Append(masterColorMapping3);

            slideLayout1.Append(commonSlideData4);
            slideLayout1.Append(colorMapOverride3);

            slideLayoutPart1.SlideLayout = slideLayout1;
        }

        // Generates content of slideMasterPart1.
        private void GenerateSlideMasterPart1Content(SlideMasterPart slideMasterPart1)
        {
            SlideMaster slideMaster1 = new SlideMaster();
            slideMaster1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideMaster1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideMaster1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData5 = new CommonSlideData();

            Background background3 = new Background();

            BackgroundStyleReference backgroundStyleReference3 = new BackgroundStyleReference() { Index = (UInt32Value)1001U };
            A.SchemeColor schemeColor74 = new A.SchemeColor() { Val = A.SchemeColorValues.Background1 };

            backgroundStyleReference3.Append(schemeColor74);

            background3.Append(backgroundStyleReference3);

            ShapeTree shapeTree5 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties17 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties69 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties17 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties69 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties17.Append(nonVisualDrawingProperties69);
            nonVisualGroupShapeProperties17.Append(nonVisualGroupShapeDrawingProperties17);
            nonVisualGroupShapeProperties17.Append(applicationNonVisualDrawingProperties69);

            GroupShapeProperties groupShapeProperties17 = new GroupShapeProperties();

            A.TransformGroup transformGroup17 = new A.TransformGroup();
            A.Offset offset66 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents66 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset17 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents17 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup17.Append(offset66);
            transformGroup17.Append(extents66);
            transformGroup17.Append(childOffset17);
            transformGroup17.Append(childExtents17);

            groupShapeProperties17.Append(transformGroup17);

            Shape shape27 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties27 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties70 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList54 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension54 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement80 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{6DF95E68-CDFE-179F-B40B-2CE59531C89D}\" />");

            // nonVisualDrawingPropertiesExtension54.Append(openXmlUnknownElement80);

            nonVisualDrawingPropertiesExtensionList54.Append(nonVisualDrawingPropertiesExtension54);

            nonVisualDrawingProperties70.Append(nonVisualDrawingPropertiesExtensionList54);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties27 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks10 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties27.Append(shapeLocks10);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties70 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape10 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties70.Append(placeholderShape10);

            nonVisualShapeProperties27.Append(nonVisualDrawingProperties70);
            nonVisualShapeProperties27.Append(nonVisualShapeDrawingProperties27);
            nonVisualShapeProperties27.Append(applicationNonVisualDrawingProperties70);

            ShapeProperties shapeProperties53 = new ShapeProperties();

            A.Transform2D transform2D49 = new A.Transform2D();
            A.Offset offset67 = new A.Offset() { X = 838200L, Y = 365125L };
            A.Extents extents67 = new A.Extents() { Cx = 10515600L, Cy = 1325563L };

            transform2D49.Append(offset67);
            transform2D49.Append(extents67);

            A.PresetGeometry presetGeometry49 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList49 = new A.AdjustValueList();

            presetGeometry49.Append(adjustValueList49);

            shapeProperties53.Append(transform2D49);
            shapeProperties53.Append(presetGeometry49);

            TextBody textBody47 = new TextBody();

            A.BodyProperties bodyProperties48 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };
            A.NormalAutoFit normalAutoFit1 = new A.NormalAutoFit();

            bodyProperties48.Append(normalAutoFit1);
            A.ListStyle listStyle48 = new A.ListStyle();

            A.Paragraph paragraph51 = new A.Paragraph();

            A.Run run32 = new A.Run();
            A.RunProperties runProperties35 = new A.RunProperties() { Language = "en-US" };
            A.Text text35 = new A.Text();
            text35.Text = "Click to edit Master title style";

            run32.Append(runProperties35);
            run32.Append(text35);

            paragraph51.Append(run32);

            textBody47.Append(bodyProperties48);
            textBody47.Append(listStyle48);
            textBody47.Append(paragraph51);

            shape27.Append(nonVisualShapeProperties27);
            shape27.Append(shapeProperties53);
            shape27.Append(textBody47);

            Shape shape28 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties28 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties71 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Text Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList55 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension55 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement81 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{2B1C2CD1-5E64-671A-E456-46FCF72C199D}\" />");

            // nonVisualDrawingPropertiesExtension55.Append(openXmlUnknownElement81);

            nonVisualDrawingPropertiesExtensionList55.Append(nonVisualDrawingPropertiesExtension55);

            nonVisualDrawingProperties71.Append(nonVisualDrawingPropertiesExtensionList55);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties28 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks11 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties28.Append(shapeLocks11);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties71 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape11 = new PlaceholderShape() { Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties71.Append(placeholderShape11);

            nonVisualShapeProperties28.Append(nonVisualDrawingProperties71);
            nonVisualShapeProperties28.Append(nonVisualShapeDrawingProperties28);
            nonVisualShapeProperties28.Append(applicationNonVisualDrawingProperties71);

            ShapeProperties shapeProperties54 = new ShapeProperties();

            A.Transform2D transform2D50 = new A.Transform2D();
            A.Offset offset68 = new A.Offset() { X = 838200L, Y = 1825625L };
            A.Extents extents68 = new A.Extents() { Cx = 10515600L, Cy = 4351338L };

            transform2D50.Append(offset68);
            transform2D50.Append(extents68);

            A.PresetGeometry presetGeometry50 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList50 = new A.AdjustValueList();

            presetGeometry50.Append(adjustValueList50);

            shapeProperties54.Append(transform2D50);
            shapeProperties54.Append(presetGeometry50);

            TextBody textBody48 = new TextBody();

            A.BodyProperties bodyProperties49 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false };
            A.NormalAutoFit normalAutoFit2 = new A.NormalAutoFit();

            bodyProperties49.Append(normalAutoFit2);
            A.ListStyle listStyle49 = new A.ListStyle();

            A.Paragraph paragraph52 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties40 = new A.ParagraphProperties() { Level = 0 };

            A.Run run33 = new A.Run();
            A.RunProperties runProperties36 = new A.RunProperties() { Language = "en-US" };
            A.Text text36 = new A.Text();
            text36.Text = "Click to edit Master text styles";

            run33.Append(runProperties36);
            run33.Append(text36);

            paragraph52.Append(paragraphProperties40);
            paragraph52.Append(run33);

            A.Paragraph paragraph53 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties41 = new A.ParagraphProperties() { Level = 1 };

            A.Run run34 = new A.Run();
            A.RunProperties runProperties37 = new A.RunProperties() { Language = "en-US" };
            A.Text text37 = new A.Text();
            text37.Text = "Second level";

            run34.Append(runProperties37);
            run34.Append(text37);

            paragraph53.Append(paragraphProperties41);
            paragraph53.Append(run34);

            A.Paragraph paragraph54 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties42 = new A.ParagraphProperties() { Level = 2 };

            A.Run run35 = new A.Run();
            A.RunProperties runProperties38 = new A.RunProperties() { Language = "en-US" };
            A.Text text38 = new A.Text();
            text38.Text = "Third level";

            run35.Append(runProperties38);
            run35.Append(text38);

            paragraph54.Append(paragraphProperties42);
            paragraph54.Append(run35);

            A.Paragraph paragraph55 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties43 = new A.ParagraphProperties() { Level = 3 };

            A.Run run36 = new A.Run();
            A.RunProperties runProperties39 = new A.RunProperties() { Language = "en-US" };
            A.Text text39 = new A.Text();
            text39.Text = "Fourth level";

            run36.Append(runProperties39);
            run36.Append(text39);

            paragraph55.Append(paragraphProperties43);
            paragraph55.Append(run36);

            A.Paragraph paragraph56 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties44 = new A.ParagraphProperties() { Level = 4 };

            A.Run run37 = new A.Run();
            A.RunProperties runProperties40 = new A.RunProperties() { Language = "en-US" };
            A.Text text40 = new A.Text();
            text40.Text = "Fifth level";

            run37.Append(runProperties40);
            run37.Append(text40);

            paragraph56.Append(paragraphProperties44);
            paragraph56.Append(run37);

            textBody48.Append(bodyProperties49);
            textBody48.Append(listStyle49);
            textBody48.Append(paragraph52);
            textBody48.Append(paragraph53);
            textBody48.Append(paragraph54);
            textBody48.Append(paragraph55);
            textBody48.Append(paragraph56);

            shape28.Append(nonVisualShapeProperties28);
            shape28.Append(shapeProperties54);
            shape28.Append(textBody48);

            Shape shape29 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties29 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties72 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList56 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension56 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement82 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{C1310E64-FEBF-708C-0497-2FE119F00E07}\" />");

            // nonVisualDrawingPropertiesExtension56.Append(openXmlUnknownElement82);

            nonVisualDrawingPropertiesExtensionList56.Append(nonVisualDrawingPropertiesExtension56);

            nonVisualDrawingProperties72.Append(nonVisualDrawingPropertiesExtensionList56);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties29 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks12 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties29.Append(shapeLocks12);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties72 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape12 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties72.Append(placeholderShape12);

            nonVisualShapeProperties29.Append(nonVisualDrawingProperties72);
            nonVisualShapeProperties29.Append(nonVisualShapeDrawingProperties29);
            nonVisualShapeProperties29.Append(applicationNonVisualDrawingProperties72);

            ShapeProperties shapeProperties55 = new ShapeProperties();

            A.Transform2D transform2D51 = new A.Transform2D();
            A.Offset offset69 = new A.Offset() { X = 838200L, Y = 6356350L };
            A.Extents extents69 = new A.Extents() { Cx = 2743200L, Cy = 365125L };

            transform2D51.Append(offset69);
            transform2D51.Append(extents69);

            A.PresetGeometry presetGeometry51 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList51 = new A.AdjustValueList();

            presetGeometry51.Append(adjustValueList51);

            shapeProperties55.Append(transform2D51);
            shapeProperties55.Append(presetGeometry51);

            TextBody textBody49 = new TextBody();
            A.BodyProperties bodyProperties50 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle50 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties7 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Left };

            A.DefaultRunProperties defaultRunProperties39 = new A.DefaultRunProperties() { FontSize = 1200 };

            A.SolidFill solidFill86 = new A.SolidFill();

            A.SchemeColor schemeColor75 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint8 = new A.Tint() { Val = 82000 };

            schemeColor75.Append(tint8);

            solidFill86.Append(schemeColor75);

            defaultRunProperties39.Append(solidFill86);

            level1ParagraphProperties7.Append(defaultRunProperties39);

            listStyle50.Append(level1ParagraphProperties7);

            A.Paragraph paragraph57 = new A.Paragraph();

            A.Field field4 = new A.Field() { Id = "{B5CF329D-DBF6-4E15-9822-D0979B6928D0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties41 = new A.RunProperties() { Language = "en-US" };
            runProperties41.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text41 = new A.Text();
            text41.Text = "4/24/2024";

            field4.Append(runProperties41);
            field4.Append(text41);
            A.EndParagraphRunProperties endParagraphRunProperties41 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph57.Append(field4);
            paragraph57.Append(endParagraphRunProperties41);

            textBody49.Append(bodyProperties50);
            textBody49.Append(listStyle50);
            textBody49.Append(paragraph57);

            shape29.Append(nonVisualShapeProperties29);
            shape29.Append(shapeProperties55);
            shape29.Append(textBody49);

            Shape shape30 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties30 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties73 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList57 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension57 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement83 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{A58D1CDC-138A-8B40-7CAC-3579DC4A2883}\" />");

            // nonVisualDrawingPropertiesExtension57.Append(openXmlUnknownElement83);

            nonVisualDrawingPropertiesExtensionList57.Append(nonVisualDrawingPropertiesExtension57);

            nonVisualDrawingProperties73.Append(nonVisualDrawingPropertiesExtensionList57);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties30 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks13 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties30.Append(shapeLocks13);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties73 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape13 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)3U };

            applicationNonVisualDrawingProperties73.Append(placeholderShape13);

            nonVisualShapeProperties30.Append(nonVisualDrawingProperties73);
            nonVisualShapeProperties30.Append(nonVisualShapeDrawingProperties30);
            nonVisualShapeProperties30.Append(applicationNonVisualDrawingProperties73);

            ShapeProperties shapeProperties56 = new ShapeProperties();

            A.Transform2D transform2D52 = new A.Transform2D();
            A.Offset offset70 = new A.Offset() { X = 4038600L, Y = 6356350L };
            A.Extents extents70 = new A.Extents() { Cx = 4114800L, Cy = 365125L };

            transform2D52.Append(offset70);
            transform2D52.Append(extents70);

            A.PresetGeometry presetGeometry52 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList52 = new A.AdjustValueList();

            presetGeometry52.Append(adjustValueList52);

            shapeProperties56.Append(transform2D52);
            shapeProperties56.Append(presetGeometry52);

            TextBody textBody50 = new TextBody();
            A.BodyProperties bodyProperties51 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle51 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties8 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };

            A.DefaultRunProperties defaultRunProperties40 = new A.DefaultRunProperties() { FontSize = 1200 };

            A.SolidFill solidFill87 = new A.SolidFill();

            A.SchemeColor schemeColor76 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint9 = new A.Tint() { Val = 82000 };

            schemeColor76.Append(tint9);

            solidFill87.Append(schemeColor76);

            defaultRunProperties40.Append(solidFill87);

            level1ParagraphProperties8.Append(defaultRunProperties40);

            listStyle51.Append(level1ParagraphProperties8);

            A.Paragraph paragraph58 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties42 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph58.Append(endParagraphRunProperties42);

            textBody50.Append(bodyProperties51);
            textBody50.Append(listStyle51);
            textBody50.Append(paragraph58);

            shape30.Append(nonVisualShapeProperties30);
            shape30.Append(shapeProperties56);
            shape30.Append(textBody50);

            Shape shape31 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties31 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties74 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList58 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension58 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement84 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{9E270B75-6D5D-49F8-4462-82107F5738B7}\" />");

            // nonVisualDrawingPropertiesExtension58.Append(openXmlUnknownElement84);

            nonVisualDrawingPropertiesExtensionList58.Append(nonVisualDrawingPropertiesExtension58);

            nonVisualDrawingProperties74.Append(nonVisualDrawingPropertiesExtensionList58);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties31 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks14 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties31.Append(shapeLocks14);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties74 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape14 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)4U };

            applicationNonVisualDrawingProperties74.Append(placeholderShape14);

            nonVisualShapeProperties31.Append(nonVisualDrawingProperties74);
            nonVisualShapeProperties31.Append(nonVisualShapeDrawingProperties31);
            nonVisualShapeProperties31.Append(applicationNonVisualDrawingProperties74);

            ShapeProperties shapeProperties57 = new ShapeProperties();

            A.Transform2D transform2D53 = new A.Transform2D();
            A.Offset offset71 = new A.Offset() { X = 8610600L, Y = 6356350L };
            A.Extents extents71 = new A.Extents() { Cx = 2743200L, Cy = 365125L };

            transform2D53.Append(offset71);
            transform2D53.Append(extents71);

            A.PresetGeometry presetGeometry53 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList53 = new A.AdjustValueList();

            presetGeometry53.Append(adjustValueList53);

            shapeProperties57.Append(transform2D53);
            shapeProperties57.Append(presetGeometry53);

            TextBody textBody51 = new TextBody();
            A.BodyProperties bodyProperties52 = new A.BodyProperties() { Vertical = A.TextVerticalValues.Horizontal, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, RightToLeftColumns = false, Anchor = A.TextAnchoringTypeValues.Center };

            A.ListStyle listStyle52 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties9 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Right };

            A.DefaultRunProperties defaultRunProperties41 = new A.DefaultRunProperties() { FontSize = 1200 };

            A.SolidFill solidFill88 = new A.SolidFill();

            A.SchemeColor schemeColor77 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint10 = new A.Tint() { Val = 82000 };

            schemeColor77.Append(tint10);

            solidFill88.Append(schemeColor77);

            defaultRunProperties41.Append(solidFill88);

            level1ParagraphProperties9.Append(defaultRunProperties41);

            listStyle52.Append(level1ParagraphProperties9);

            A.Paragraph paragraph59 = new A.Paragraph();

            A.Field field5 = new A.Field() { Id = "{19E6023F-2333-4A80-AFF1-A7F70B52B230}", Type = "slidenum" };

            A.RunProperties runProperties42 = new A.RunProperties() { Language = "en-US" };
            runProperties42.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text42 = new A.Text();
            text42.Text = "‹#›";

            field5.Append(runProperties42);
            field5.Append(text42);
            A.EndParagraphRunProperties endParagraphRunProperties43 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph59.Append(field5);
            paragraph59.Append(endParagraphRunProperties43);

            textBody51.Append(bodyProperties52);
            textBody51.Append(listStyle52);
            textBody51.Append(paragraph59);

            shape31.Append(nonVisualShapeProperties31);
            shape31.Append(shapeProperties57);
            shape31.Append(textBody51);

            shapeTree5.Append(nonVisualGroupShapeProperties17);
            shapeTree5.Append(groupShapeProperties17);
            shapeTree5.Append(shape27);
            shapeTree5.Append(shape28);
            shapeTree5.Append(shape29);
            shapeTree5.Append(shape30);
            shapeTree5.Append(shape31);

            CommonSlideDataExtensionList commonSlideDataExtensionList5 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension5 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId5 = new P14.CreationId() { Val = (UInt32Value)3325504521U };
            creationId5.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension5.Append(creationId5);

            commonSlideDataExtensionList5.Append(commonSlideDataExtension5);

            commonSlideData5.Append(background3);
            commonSlideData5.Append(shapeTree5);
            commonSlideData5.Append(commonSlideDataExtensionList5);
            ColorMap colorMap2 = new ColorMap() { Background1 = A.ColorSchemeIndexValues.Light1, Text1 = A.ColorSchemeIndexValues.Dark1, Background2 = A.ColorSchemeIndexValues.Light2, Text2 = A.ColorSchemeIndexValues.Dark2, Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2, Accent3 = A.ColorSchemeIndexValues.Accent3, Accent4 = A.ColorSchemeIndexValues.Accent4, Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6, Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink };

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
            SlideLayoutId slideLayoutId12 = new SlideLayoutId() { Id = (UInt32Value)2147483660U, RelationshipId = "rId12" };

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
            slideLayoutIdList1.Append(slideLayoutId12);

            TextStyles textStyles1 = new TextStyles();

            TitleStyle titleStyle1 = new TitleStyle();

            A.Level1ParagraphProperties level1ParagraphProperties10 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing31 = new A.LineSpacing();
            A.SpacingPercent spacingPercent31 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing31.Append(spacingPercent31);

            A.SpaceBefore spaceBefore17 = new A.SpaceBefore();
            A.SpacingPercent spacingPercent32 = new A.SpacingPercent() { Val = 0 };

            spaceBefore17.Append(spacingPercent32);
            A.NoBullet noBullet15 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties42 = new A.DefaultRunProperties() { FontSize = 4400, Kerning = 1200 };

            A.SolidFill solidFill89 = new A.SolidFill();
            A.SchemeColor schemeColor78 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill89.Append(schemeColor78);
            A.LatinFont latinFont63 = new A.LatinFont() { Typeface = "+mj-lt" };
            A.EastAsianFont eastAsianFont53 = new A.EastAsianFont() { Typeface = "+mj-ea" };
            A.ComplexScriptFont complexScriptFont63 = new A.ComplexScriptFont() { Typeface = "+mj-cs" };

            defaultRunProperties42.Append(solidFill89);
            defaultRunProperties42.Append(latinFont63);
            defaultRunProperties42.Append(eastAsianFont53);
            defaultRunProperties42.Append(complexScriptFont63);

            level1ParagraphProperties10.Append(lineSpacing31);
            level1ParagraphProperties10.Append(spaceBefore17);
            level1ParagraphProperties10.Append(noBullet15);
            level1ParagraphProperties10.Append(defaultRunProperties42);

            titleStyle1.Append(level1ParagraphProperties10);

            BodyStyle bodyStyle1 = new BodyStyle();

            A.Level1ParagraphProperties level1ParagraphProperties11 = new A.Level1ParagraphProperties() { LeftMargin = 228600, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing32 = new A.LineSpacing();
            A.SpacingPercent spacingPercent33 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing32.Append(spacingPercent33);

            A.SpaceBefore spaceBefore18 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints33 = new A.SpacingPoints() { Val = 1000 };

            spaceBefore18.Append(spacingPoints33);
            A.BulletFont bulletFont1 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet1 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties43 = new A.DefaultRunProperties() { FontSize = 2800, Kerning = 1200 };

            A.SolidFill solidFill90 = new A.SolidFill();
            A.SchemeColor schemeColor79 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill90.Append(schemeColor79);
            A.LatinFont latinFont64 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont54 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont64 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties43.Append(solidFill90);
            defaultRunProperties43.Append(latinFont64);
            defaultRunProperties43.Append(eastAsianFont54);
            defaultRunProperties43.Append(complexScriptFont64);

            level1ParagraphProperties11.Append(lineSpacing32);
            level1ParagraphProperties11.Append(spaceBefore18);
            level1ParagraphProperties11.Append(bulletFont1);
            level1ParagraphProperties11.Append(characterBullet1);
            level1ParagraphProperties11.Append(defaultRunProperties43);

            A.Level2ParagraphProperties level2ParagraphProperties3 = new A.Level2ParagraphProperties() { LeftMargin = 685800, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing33 = new A.LineSpacing();
            A.SpacingPercent spacingPercent34 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing33.Append(spacingPercent34);

            A.SpaceBefore spaceBefore19 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints34 = new A.SpacingPoints() { Val = 500 };

            spaceBefore19.Append(spacingPoints34);
            A.BulletFont bulletFont2 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet2 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties44 = new A.DefaultRunProperties() { FontSize = 2400, Kerning = 1200 };

            A.SolidFill solidFill91 = new A.SolidFill();
            A.SchemeColor schemeColor80 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill91.Append(schemeColor80);
            A.LatinFont latinFont65 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont55 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont65 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties44.Append(solidFill91);
            defaultRunProperties44.Append(latinFont65);
            defaultRunProperties44.Append(eastAsianFont55);
            defaultRunProperties44.Append(complexScriptFont65);

            level2ParagraphProperties3.Append(lineSpacing33);
            level2ParagraphProperties3.Append(spaceBefore19);
            level2ParagraphProperties3.Append(bulletFont2);
            level2ParagraphProperties3.Append(characterBullet2);
            level2ParagraphProperties3.Append(defaultRunProperties44);

            A.Level3ParagraphProperties level3ParagraphProperties3 = new A.Level3ParagraphProperties() { LeftMargin = 1143000, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing34 = new A.LineSpacing();
            A.SpacingPercent spacingPercent35 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing34.Append(spacingPercent35);

            A.SpaceBefore spaceBefore20 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints35 = new A.SpacingPoints() { Val = 500 };

            spaceBefore20.Append(spacingPoints35);
            A.BulletFont bulletFont3 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet3 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties45 = new A.DefaultRunProperties() { FontSize = 2000, Kerning = 1200 };

            A.SolidFill solidFill92 = new A.SolidFill();
            A.SchemeColor schemeColor81 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill92.Append(schemeColor81);
            A.LatinFont latinFont66 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont56 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont66 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties45.Append(solidFill92);
            defaultRunProperties45.Append(latinFont66);
            defaultRunProperties45.Append(eastAsianFont56);
            defaultRunProperties45.Append(complexScriptFont66);

            level3ParagraphProperties3.Append(lineSpacing34);
            level3ParagraphProperties3.Append(spaceBefore20);
            level3ParagraphProperties3.Append(bulletFont3);
            level3ParagraphProperties3.Append(characterBullet3);
            level3ParagraphProperties3.Append(defaultRunProperties45);

            A.Level4ParagraphProperties level4ParagraphProperties3 = new A.Level4ParagraphProperties() { LeftMargin = 1600200, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing35 = new A.LineSpacing();
            A.SpacingPercent spacingPercent36 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing35.Append(spacingPercent36);

            A.SpaceBefore spaceBefore21 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints36 = new A.SpacingPoints() { Val = 500 };

            spaceBefore21.Append(spacingPoints36);
            A.BulletFont bulletFont4 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet4 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties46 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill93 = new A.SolidFill();
            A.SchemeColor schemeColor82 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill93.Append(schemeColor82);
            A.LatinFont latinFont67 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont57 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont67 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties46.Append(solidFill93);
            defaultRunProperties46.Append(latinFont67);
            defaultRunProperties46.Append(eastAsianFont57);
            defaultRunProperties46.Append(complexScriptFont67);

            level4ParagraphProperties3.Append(lineSpacing35);
            level4ParagraphProperties3.Append(spaceBefore21);
            level4ParagraphProperties3.Append(bulletFont4);
            level4ParagraphProperties3.Append(characterBullet4);
            level4ParagraphProperties3.Append(defaultRunProperties46);

            A.Level5ParagraphProperties level5ParagraphProperties3 = new A.Level5ParagraphProperties() { LeftMargin = 2057400, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing36 = new A.LineSpacing();
            A.SpacingPercent spacingPercent37 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing36.Append(spacingPercent37);

            A.SpaceBefore spaceBefore22 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints37 = new A.SpacingPoints() { Val = 500 };

            spaceBefore22.Append(spacingPoints37);
            A.BulletFont bulletFont5 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet5 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties47 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill94 = new A.SolidFill();
            A.SchemeColor schemeColor83 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill94.Append(schemeColor83);
            A.LatinFont latinFont68 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont58 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont68 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties47.Append(solidFill94);
            defaultRunProperties47.Append(latinFont68);
            defaultRunProperties47.Append(eastAsianFont58);
            defaultRunProperties47.Append(complexScriptFont68);

            level5ParagraphProperties3.Append(lineSpacing36);
            level5ParagraphProperties3.Append(spaceBefore22);
            level5ParagraphProperties3.Append(bulletFont5);
            level5ParagraphProperties3.Append(characterBullet5);
            level5ParagraphProperties3.Append(defaultRunProperties47);

            A.Level6ParagraphProperties level6ParagraphProperties3 = new A.Level6ParagraphProperties() { LeftMargin = 2514600, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing37 = new A.LineSpacing();
            A.SpacingPercent spacingPercent38 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing37.Append(spacingPercent38);

            A.SpaceBefore spaceBefore23 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints38 = new A.SpacingPoints() { Val = 500 };

            spaceBefore23.Append(spacingPoints38);
            A.BulletFont bulletFont6 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet6 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties48 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill95 = new A.SolidFill();
            A.SchemeColor schemeColor84 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill95.Append(schemeColor84);
            A.LatinFont latinFont69 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont59 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont69 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties48.Append(solidFill95);
            defaultRunProperties48.Append(latinFont69);
            defaultRunProperties48.Append(eastAsianFont59);
            defaultRunProperties48.Append(complexScriptFont69);

            level6ParagraphProperties3.Append(lineSpacing37);
            level6ParagraphProperties3.Append(spaceBefore23);
            level6ParagraphProperties3.Append(bulletFont6);
            level6ParagraphProperties3.Append(characterBullet6);
            level6ParagraphProperties3.Append(defaultRunProperties48);

            A.Level7ParagraphProperties level7ParagraphProperties3 = new A.Level7ParagraphProperties() { LeftMargin = 2971800, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing38 = new A.LineSpacing();
            A.SpacingPercent spacingPercent39 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing38.Append(spacingPercent39);

            A.SpaceBefore spaceBefore24 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints39 = new A.SpacingPoints() { Val = 500 };

            spaceBefore24.Append(spacingPoints39);
            A.BulletFont bulletFont7 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet7 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties49 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill96 = new A.SolidFill();
            A.SchemeColor schemeColor85 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill96.Append(schemeColor85);
            A.LatinFont latinFont70 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont60 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont70 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties49.Append(solidFill96);
            defaultRunProperties49.Append(latinFont70);
            defaultRunProperties49.Append(eastAsianFont60);
            defaultRunProperties49.Append(complexScriptFont70);

            level7ParagraphProperties3.Append(lineSpacing38);
            level7ParagraphProperties3.Append(spaceBefore24);
            level7ParagraphProperties3.Append(bulletFont7);
            level7ParagraphProperties3.Append(characterBullet7);
            level7ParagraphProperties3.Append(defaultRunProperties49);

            A.Level8ParagraphProperties level8ParagraphProperties3 = new A.Level8ParagraphProperties() { LeftMargin = 3429000, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing39 = new A.LineSpacing();
            A.SpacingPercent spacingPercent40 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing39.Append(spacingPercent40);

            A.SpaceBefore spaceBefore25 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints40 = new A.SpacingPoints() { Val = 500 };

            spaceBefore25.Append(spacingPoints40);
            A.BulletFont bulletFont8 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet8 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties50 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill97 = new A.SolidFill();
            A.SchemeColor schemeColor86 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill97.Append(schemeColor86);
            A.LatinFont latinFont71 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont61 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont71 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties50.Append(solidFill97);
            defaultRunProperties50.Append(latinFont71);
            defaultRunProperties50.Append(eastAsianFont61);
            defaultRunProperties50.Append(complexScriptFont71);

            level8ParagraphProperties3.Append(lineSpacing39);
            level8ParagraphProperties3.Append(spaceBefore25);
            level8ParagraphProperties3.Append(bulletFont8);
            level8ParagraphProperties3.Append(characterBullet8);
            level8ParagraphProperties3.Append(defaultRunProperties50);

            A.Level9ParagraphProperties level9ParagraphProperties3 = new A.Level9ParagraphProperties() { LeftMargin = 3886200, Indent = -228600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.LineSpacing lineSpacing40 = new A.LineSpacing();
            A.SpacingPercent spacingPercent41 = new A.SpacingPercent() { Val = 90000 };

            lineSpacing40.Append(spacingPercent41);

            A.SpaceBefore spaceBefore26 = new A.SpaceBefore();
            A.SpacingPoints spacingPoints41 = new A.SpacingPoints() { Val = 500 };

            spaceBefore26.Append(spacingPoints41);
            A.BulletFont bulletFont9 = new A.BulletFont() { Typeface = "Arial", Panose = "020B0604020202020204", PitchFamily = 34, CharacterSet = 0 };
            A.CharacterBullet characterBullet9 = new A.CharacterBullet() { Char = "•" };

            A.DefaultRunProperties defaultRunProperties51 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill98 = new A.SolidFill();
            A.SchemeColor schemeColor87 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill98.Append(schemeColor87);
            A.LatinFont latinFont72 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont62 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont72 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties51.Append(solidFill98);
            defaultRunProperties51.Append(latinFont72);
            defaultRunProperties51.Append(eastAsianFont62);
            defaultRunProperties51.Append(complexScriptFont72);

            level9ParagraphProperties3.Append(lineSpacing40);
            level9ParagraphProperties3.Append(spaceBefore26);
            level9ParagraphProperties3.Append(bulletFont9);
            level9ParagraphProperties3.Append(characterBullet9);
            level9ParagraphProperties3.Append(defaultRunProperties51);

            bodyStyle1.Append(level1ParagraphProperties11);
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
            A.DefaultRunProperties defaultRunProperties52 = new A.DefaultRunProperties() { Language = "en-US" };

            defaultParagraphProperties2.Append(defaultRunProperties52);

            A.Level1ParagraphProperties level1ParagraphProperties12 = new A.Level1ParagraphProperties() { LeftMargin = 0, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties53 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill99 = new A.SolidFill();
            A.SchemeColor schemeColor88 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill99.Append(schemeColor88);
            A.LatinFont latinFont73 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont63 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont73 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties53.Append(solidFill99);
            defaultRunProperties53.Append(latinFont73);
            defaultRunProperties53.Append(eastAsianFont63);
            defaultRunProperties53.Append(complexScriptFont73);

            level1ParagraphProperties12.Append(defaultRunProperties53);

            A.Level2ParagraphProperties level2ParagraphProperties4 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties54 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill100 = new A.SolidFill();
            A.SchemeColor schemeColor89 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill100.Append(schemeColor89);
            A.LatinFont latinFont74 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont64 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont74 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties54.Append(solidFill100);
            defaultRunProperties54.Append(latinFont74);
            defaultRunProperties54.Append(eastAsianFont64);
            defaultRunProperties54.Append(complexScriptFont74);

            level2ParagraphProperties4.Append(defaultRunProperties54);

            A.Level3ParagraphProperties level3ParagraphProperties4 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties55 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill101 = new A.SolidFill();
            A.SchemeColor schemeColor90 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill101.Append(schemeColor90);
            A.LatinFont latinFont75 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont65 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont75 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties55.Append(solidFill101);
            defaultRunProperties55.Append(latinFont75);
            defaultRunProperties55.Append(eastAsianFont65);
            defaultRunProperties55.Append(complexScriptFont75);

            level3ParagraphProperties4.Append(defaultRunProperties55);

            A.Level4ParagraphProperties level4ParagraphProperties4 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties56 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill102 = new A.SolidFill();
            A.SchemeColor schemeColor91 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill102.Append(schemeColor91);
            A.LatinFont latinFont76 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont66 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont76 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties56.Append(solidFill102);
            defaultRunProperties56.Append(latinFont76);
            defaultRunProperties56.Append(eastAsianFont66);
            defaultRunProperties56.Append(complexScriptFont76);

            level4ParagraphProperties4.Append(defaultRunProperties56);

            A.Level5ParagraphProperties level5ParagraphProperties4 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties57 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill103 = new A.SolidFill();
            A.SchemeColor schemeColor92 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill103.Append(schemeColor92);
            A.LatinFont latinFont77 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont67 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont77 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties57.Append(solidFill103);
            defaultRunProperties57.Append(latinFont77);
            defaultRunProperties57.Append(eastAsianFont67);
            defaultRunProperties57.Append(complexScriptFont77);

            level5ParagraphProperties4.Append(defaultRunProperties57);

            A.Level6ParagraphProperties level6ParagraphProperties4 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties58 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill104 = new A.SolidFill();
            A.SchemeColor schemeColor93 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill104.Append(schemeColor93);
            A.LatinFont latinFont78 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont68 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont78 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties58.Append(solidFill104);
            defaultRunProperties58.Append(latinFont78);
            defaultRunProperties58.Append(eastAsianFont68);
            defaultRunProperties58.Append(complexScriptFont78);

            level6ParagraphProperties4.Append(defaultRunProperties58);

            A.Level7ParagraphProperties level7ParagraphProperties4 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties59 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill105 = new A.SolidFill();
            A.SchemeColor schemeColor94 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill105.Append(schemeColor94);
            A.LatinFont latinFont79 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont69 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont79 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties59.Append(solidFill105);
            defaultRunProperties59.Append(latinFont79);
            defaultRunProperties59.Append(eastAsianFont69);
            defaultRunProperties59.Append(complexScriptFont79);

            level7ParagraphProperties4.Append(defaultRunProperties59);

            A.Level8ParagraphProperties level8ParagraphProperties4 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties60 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill106 = new A.SolidFill();
            A.SchemeColor schemeColor95 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill106.Append(schemeColor95);
            A.LatinFont latinFont80 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont70 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont80 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties60.Append(solidFill106);
            defaultRunProperties60.Append(latinFont80);
            defaultRunProperties60.Append(eastAsianFont70);
            defaultRunProperties60.Append(complexScriptFont80);

            level8ParagraphProperties4.Append(defaultRunProperties60);

            A.Level9ParagraphProperties level9ParagraphProperties4 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Alignment = A.TextAlignmentTypeValues.Left, DefaultTabSize = 914400, RightToLeft = false, EastAsianLineBreak = true, LatinLineBreak = false, Height = true };

            A.DefaultRunProperties defaultRunProperties61 = new A.DefaultRunProperties() { FontSize = 1800, Kerning = 1200 };

            A.SolidFill solidFill107 = new A.SolidFill();
            A.SchemeColor schemeColor96 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            solidFill107.Append(schemeColor96);
            A.LatinFont latinFont81 = new A.LatinFont() { Typeface = "+mn-lt" };
            A.EastAsianFont eastAsianFont71 = new A.EastAsianFont() { Typeface = "+mn-ea" };
            A.ComplexScriptFont complexScriptFont81 = new A.ComplexScriptFont() { Typeface = "+mn-cs" };

            defaultRunProperties61.Append(solidFill107);
            defaultRunProperties61.Append(latinFont81);
            defaultRunProperties61.Append(eastAsianFont71);
            defaultRunProperties61.Append(complexScriptFont81);

            level9ParagraphProperties4.Append(defaultRunProperties61);

            otherStyle1.Append(defaultParagraphProperties2);
            otherStyle1.Append(level1ParagraphProperties12);
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

            slideMaster1.Append(commonSlideData5);
            slideMaster1.Append(colorMap2);
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

            CommonSlideData commonSlideData6 = new CommonSlideData() { Name = "Content with Caption" };

            ShapeTree shapeTree6 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties18 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties75 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties18 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties75 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties18.Append(nonVisualDrawingProperties75);
            nonVisualGroupShapeProperties18.Append(nonVisualGroupShapeDrawingProperties18);
            nonVisualGroupShapeProperties18.Append(applicationNonVisualDrawingProperties75);

            GroupShapeProperties groupShapeProperties18 = new GroupShapeProperties();

            A.TransformGroup transformGroup18 = new A.TransformGroup();
            A.Offset offset72 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents72 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset18 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents18 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup18.Append(offset72);
            transformGroup18.Append(extents72);
            transformGroup18.Append(childOffset18);
            transformGroup18.Append(childExtents18);

            groupShapeProperties18.Append(transformGroup18);

            Shape shape32 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties32 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties76 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList59 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension59 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement85 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{AB2654FA-4ED6-9213-94F9-3CD4BFDC1635}\" />");

            // nonVisualDrawingPropertiesExtension59.Append(openXmlUnknownElement85);

            nonVisualDrawingPropertiesExtensionList59.Append(nonVisualDrawingPropertiesExtension59);

            nonVisualDrawingProperties76.Append(nonVisualDrawingPropertiesExtensionList59);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties32 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks15 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties32.Append(shapeLocks15);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties76 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape15 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties76.Append(placeholderShape15);

            nonVisualShapeProperties32.Append(nonVisualDrawingProperties76);
            nonVisualShapeProperties32.Append(nonVisualShapeDrawingProperties32);
            nonVisualShapeProperties32.Append(applicationNonVisualDrawingProperties76);

            ShapeProperties shapeProperties58 = new ShapeProperties();

            A.Transform2D transform2D54 = new A.Transform2D();
            A.Offset offset73 = new A.Offset() { X = 839788L, Y = 457200L };
            A.Extents extents73 = new A.Extents() { Cx = 3932237L, Cy = 1600200L };

            transform2D54.Append(offset73);
            transform2D54.Append(extents73);

            shapeProperties58.Append(transform2D54);

            TextBody textBody52 = new TextBody();
            A.BodyProperties bodyProperties53 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle53 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties13 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties62 = new A.DefaultRunProperties() { FontSize = 3200 };

            level1ParagraphProperties13.Append(defaultRunProperties62);

            listStyle53.Append(level1ParagraphProperties13);

            A.Paragraph paragraph60 = new A.Paragraph();

            A.Run run38 = new A.Run();
            A.RunProperties runProperties43 = new A.RunProperties() { Language = "en-US" };
            A.Text text43 = new A.Text();
            text43.Text = "Click to edit Master title style";

            run38.Append(runProperties43);
            run38.Append(text43);

            paragraph60.Append(run38);

            textBody52.Append(bodyProperties53);
            textBody52.Append(listStyle53);
            textBody52.Append(paragraph60);

            shape32.Append(nonVisualShapeProperties32);
            shape32.Append(shapeProperties58);
            shape32.Append(textBody52);

            Shape shape33 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties33 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties77 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Content Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList60 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension60 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement86 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{4D82C328-7919-C7A5-3098-95756631E119}\" />");

            // nonVisualDrawingPropertiesExtension60.Append(openXmlUnknownElement86);

            nonVisualDrawingPropertiesExtensionList60.Append(nonVisualDrawingPropertiesExtension60);

            nonVisualDrawingProperties77.Append(nonVisualDrawingPropertiesExtensionList60);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties33 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks16 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties33.Append(shapeLocks16);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties77 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape16 = new PlaceholderShape() { Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties77.Append(placeholderShape16);

            nonVisualShapeProperties33.Append(nonVisualDrawingProperties77);
            nonVisualShapeProperties33.Append(nonVisualShapeDrawingProperties33);
            nonVisualShapeProperties33.Append(applicationNonVisualDrawingProperties77);

            ShapeProperties shapeProperties59 = new ShapeProperties();

            A.Transform2D transform2D55 = new A.Transform2D();
            A.Offset offset74 = new A.Offset() { X = 5183188L, Y = 987425L };
            A.Extents extents74 = new A.Extents() { Cx = 6172200L, Cy = 4873625L };

            transform2D55.Append(offset74);
            transform2D55.Append(extents74);

            shapeProperties59.Append(transform2D55);

            TextBody textBody53 = new TextBody();
            A.BodyProperties bodyProperties54 = new A.BodyProperties();

            A.ListStyle listStyle54 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties14 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties63 = new A.DefaultRunProperties() { FontSize = 3200 };

            level1ParagraphProperties14.Append(defaultRunProperties63);

            A.Level2ParagraphProperties level2ParagraphProperties5 = new A.Level2ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties64 = new A.DefaultRunProperties() { FontSize = 2800 };

            level2ParagraphProperties5.Append(defaultRunProperties64);

            A.Level3ParagraphProperties level3ParagraphProperties5 = new A.Level3ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties65 = new A.DefaultRunProperties() { FontSize = 2400 };

            level3ParagraphProperties5.Append(defaultRunProperties65);

            A.Level4ParagraphProperties level4ParagraphProperties5 = new A.Level4ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties66 = new A.DefaultRunProperties() { FontSize = 2000 };

            level4ParagraphProperties5.Append(defaultRunProperties66);

            A.Level5ParagraphProperties level5ParagraphProperties5 = new A.Level5ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties67 = new A.DefaultRunProperties() { FontSize = 2000 };

            level5ParagraphProperties5.Append(defaultRunProperties67);

            A.Level6ParagraphProperties level6ParagraphProperties5 = new A.Level6ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties68 = new A.DefaultRunProperties() { FontSize = 2000 };

            level6ParagraphProperties5.Append(defaultRunProperties68);

            A.Level7ParagraphProperties level7ParagraphProperties5 = new A.Level7ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties69 = new A.DefaultRunProperties() { FontSize = 2000 };

            level7ParagraphProperties5.Append(defaultRunProperties69);

            A.Level8ParagraphProperties level8ParagraphProperties5 = new A.Level8ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties70 = new A.DefaultRunProperties() { FontSize = 2000 };

            level8ParagraphProperties5.Append(defaultRunProperties70);

            A.Level9ParagraphProperties level9ParagraphProperties5 = new A.Level9ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties71 = new A.DefaultRunProperties() { FontSize = 2000 };

            level9ParagraphProperties5.Append(defaultRunProperties71);

            listStyle54.Append(level1ParagraphProperties14);
            listStyle54.Append(level2ParagraphProperties5);
            listStyle54.Append(level3ParagraphProperties5);
            listStyle54.Append(level4ParagraphProperties5);
            listStyle54.Append(level5ParagraphProperties5);
            listStyle54.Append(level6ParagraphProperties5);
            listStyle54.Append(level7ParagraphProperties5);
            listStyle54.Append(level8ParagraphProperties5);
            listStyle54.Append(level9ParagraphProperties5);

            A.Paragraph paragraph61 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties45 = new A.ParagraphProperties() { Level = 0 };

            A.Run run39 = new A.Run();
            A.RunProperties runProperties44 = new A.RunProperties() { Language = "en-US" };
            A.Text text44 = new A.Text();
            text44.Text = "Click to edit Master text styles";

            run39.Append(runProperties44);
            run39.Append(text44);

            paragraph61.Append(paragraphProperties45);
            paragraph61.Append(run39);

            A.Paragraph paragraph62 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties46 = new A.ParagraphProperties() { Level = 1 };

            A.Run run40 = new A.Run();
            A.RunProperties runProperties45 = new A.RunProperties() { Language = "en-US" };
            A.Text text45 = new A.Text();
            text45.Text = "Second level";

            run40.Append(runProperties45);
            run40.Append(text45);

            paragraph62.Append(paragraphProperties46);
            paragraph62.Append(run40);

            A.Paragraph paragraph63 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties47 = new A.ParagraphProperties() { Level = 2 };

            A.Run run41 = new A.Run();
            A.RunProperties runProperties46 = new A.RunProperties() { Language = "en-US" };
            A.Text text46 = new A.Text();
            text46.Text = "Third level";

            run41.Append(runProperties46);
            run41.Append(text46);

            paragraph63.Append(paragraphProperties47);
            paragraph63.Append(run41);

            A.Paragraph paragraph64 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties48 = new A.ParagraphProperties() { Level = 3 };

            A.Run run42 = new A.Run();
            A.RunProperties runProperties47 = new A.RunProperties() { Language = "en-US" };
            A.Text text47 = new A.Text();
            text47.Text = "Fourth level";

            run42.Append(runProperties47);
            run42.Append(text47);

            paragraph64.Append(paragraphProperties48);
            paragraph64.Append(run42);

            A.Paragraph paragraph65 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties49 = new A.ParagraphProperties() { Level = 4 };

            A.Run run43 = new A.Run();
            A.RunProperties runProperties48 = new A.RunProperties() { Language = "en-US" };
            A.Text text48 = new A.Text();
            text48.Text = "Fifth level";

            run43.Append(runProperties48);
            run43.Append(text48);

            paragraph65.Append(paragraphProperties49);
            paragraph65.Append(run43);

            textBody53.Append(bodyProperties54);
            textBody53.Append(listStyle54);
            textBody53.Append(paragraph61);
            textBody53.Append(paragraph62);
            textBody53.Append(paragraph63);
            textBody53.Append(paragraph64);
            textBody53.Append(paragraph65);

            shape33.Append(nonVisualShapeProperties33);
            shape33.Append(shapeProperties59);
            shape33.Append(textBody53);

            Shape shape34 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties34 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties78 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Text Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList61 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension61 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement87 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{08F89E76-E41E-1085-0D60-A7AEF0019888}\" />");

            // nonVisualDrawingPropertiesExtension61.Append(openXmlUnknownElement87);

            nonVisualDrawingPropertiesExtensionList61.Append(nonVisualDrawingPropertiesExtension61);

            nonVisualDrawingProperties78.Append(nonVisualDrawingPropertiesExtensionList61);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties34 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks17 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties34.Append(shapeLocks17);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties78 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape17 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties78.Append(placeholderShape17);

            nonVisualShapeProperties34.Append(nonVisualDrawingProperties78);
            nonVisualShapeProperties34.Append(nonVisualShapeDrawingProperties34);
            nonVisualShapeProperties34.Append(applicationNonVisualDrawingProperties78);

            ShapeProperties shapeProperties60 = new ShapeProperties();

            A.Transform2D transform2D56 = new A.Transform2D();
            A.Offset offset75 = new A.Offset() { X = 839788L, Y = 2057400L };
            A.Extents extents75 = new A.Extents() { Cx = 3932237L, Cy = 3811588L };

            transform2D56.Append(offset75);
            transform2D56.Append(extents75);

            shapeProperties60.Append(transform2D56);

            TextBody textBody54 = new TextBody();
            A.BodyProperties bodyProperties55 = new A.BodyProperties();

            A.ListStyle listStyle55 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties15 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet16 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties72 = new A.DefaultRunProperties() { FontSize = 1600 };

            level1ParagraphProperties15.Append(noBullet16);
            level1ParagraphProperties15.Append(defaultRunProperties72);

            A.Level2ParagraphProperties level2ParagraphProperties6 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet17 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties73 = new A.DefaultRunProperties() { FontSize = 1400 };

            level2ParagraphProperties6.Append(noBullet17);
            level2ParagraphProperties6.Append(defaultRunProperties73);

            A.Level3ParagraphProperties level3ParagraphProperties6 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet18 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties74 = new A.DefaultRunProperties() { FontSize = 1200 };

            level3ParagraphProperties6.Append(noBullet18);
            level3ParagraphProperties6.Append(defaultRunProperties74);

            A.Level4ParagraphProperties level4ParagraphProperties6 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet19 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties75 = new A.DefaultRunProperties() { FontSize = 1000 };

            level4ParagraphProperties6.Append(noBullet19);
            level4ParagraphProperties6.Append(defaultRunProperties75);

            A.Level5ParagraphProperties level5ParagraphProperties6 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet20 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties76 = new A.DefaultRunProperties() { FontSize = 1000 };

            level5ParagraphProperties6.Append(noBullet20);
            level5ParagraphProperties6.Append(defaultRunProperties76);

            A.Level6ParagraphProperties level6ParagraphProperties6 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet21 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties77 = new A.DefaultRunProperties() { FontSize = 1000 };

            level6ParagraphProperties6.Append(noBullet21);
            level6ParagraphProperties6.Append(defaultRunProperties77);

            A.Level7ParagraphProperties level7ParagraphProperties6 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet22 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties78 = new A.DefaultRunProperties() { FontSize = 1000 };

            level7ParagraphProperties6.Append(noBullet22);
            level7ParagraphProperties6.Append(defaultRunProperties78);

            A.Level8ParagraphProperties level8ParagraphProperties6 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet23 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties79 = new A.DefaultRunProperties() { FontSize = 1000 };

            level8ParagraphProperties6.Append(noBullet23);
            level8ParagraphProperties6.Append(defaultRunProperties79);

            A.Level9ParagraphProperties level9ParagraphProperties6 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet24 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties80 = new A.DefaultRunProperties() { FontSize = 1000 };

            level9ParagraphProperties6.Append(noBullet24);
            level9ParagraphProperties6.Append(defaultRunProperties80);

            listStyle55.Append(level1ParagraphProperties15);
            listStyle55.Append(level2ParagraphProperties6);
            listStyle55.Append(level3ParagraphProperties6);
            listStyle55.Append(level4ParagraphProperties6);
            listStyle55.Append(level5ParagraphProperties6);
            listStyle55.Append(level6ParagraphProperties6);
            listStyle55.Append(level7ParagraphProperties6);
            listStyle55.Append(level8ParagraphProperties6);
            listStyle55.Append(level9ParagraphProperties6);

            A.Paragraph paragraph66 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties50 = new A.ParagraphProperties() { Level = 0 };

            A.Run run44 = new A.Run();
            A.RunProperties runProperties49 = new A.RunProperties() { Language = "en-US" };
            A.Text text49 = new A.Text();
            text49.Text = "Click to edit Master text styles";

            run44.Append(runProperties49);
            run44.Append(text49);

            paragraph66.Append(paragraphProperties50);
            paragraph66.Append(run44);

            textBody54.Append(bodyProperties55);
            textBody54.Append(listStyle55);
            textBody54.Append(paragraph66);

            shape34.Append(nonVisualShapeProperties34);
            shape34.Append(shapeProperties60);
            shape34.Append(textBody54);

            Shape shape35 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties35 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties79 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Date Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList62 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension62 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement88 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{D3694F81-93CF-EC5A-C1BB-9D466C79C98A}\" />");

            // nonVisualDrawingPropertiesExtension62.Append(openXmlUnknownElement88);

            nonVisualDrawingPropertiesExtensionList62.Append(nonVisualDrawingPropertiesExtension62);

            nonVisualDrawingProperties79.Append(nonVisualDrawingPropertiesExtensionList62);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties35 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks18 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties35.Append(shapeLocks18);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties79 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape18 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties79.Append(placeholderShape18);

            nonVisualShapeProperties35.Append(nonVisualDrawingProperties79);
            nonVisualShapeProperties35.Append(nonVisualShapeDrawingProperties35);
            nonVisualShapeProperties35.Append(applicationNonVisualDrawingProperties79);
            ShapeProperties shapeProperties61 = new ShapeProperties();

            TextBody textBody55 = new TextBody();
            A.BodyProperties bodyProperties56 = new A.BodyProperties();
            A.ListStyle listStyle56 = new A.ListStyle();

            A.Paragraph paragraph67 = new A.Paragraph();

            A.Field field6 = new A.Field() { Id = "{B5CF329D-DBF6-4E15-9822-D0979B6928D0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties50 = new A.RunProperties() { Language = "en-US" };
            runProperties50.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text50 = new A.Text();
            text50.Text = "4/24/2024";

            field6.Append(runProperties50);
            field6.Append(text50);
            A.EndParagraphRunProperties endParagraphRunProperties44 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph67.Append(field6);
            paragraph67.Append(endParagraphRunProperties44);

            textBody55.Append(bodyProperties56);
            textBody55.Append(listStyle56);
            textBody55.Append(paragraph67);

            shape35.Append(nonVisualShapeProperties35);
            shape35.Append(shapeProperties61);
            shape35.Append(textBody55);

            Shape shape36 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties36 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties80 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Footer Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList63 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension63 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement89 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{13E74D08-319D-343B-ABC3-21DCB541EB20}\" />");

            // nonVisualDrawingPropertiesExtension63.Append(openXmlUnknownElement89);

            nonVisualDrawingPropertiesExtensionList63.Append(nonVisualDrawingPropertiesExtension63);

            nonVisualDrawingProperties80.Append(nonVisualDrawingPropertiesExtensionList63);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties36 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks19 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties36.Append(shapeLocks19);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties80 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape19 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties80.Append(placeholderShape19);

            nonVisualShapeProperties36.Append(nonVisualDrawingProperties80);
            nonVisualShapeProperties36.Append(nonVisualShapeDrawingProperties36);
            nonVisualShapeProperties36.Append(applicationNonVisualDrawingProperties80);
            ShapeProperties shapeProperties62 = new ShapeProperties();

            TextBody textBody56 = new TextBody();
            A.BodyProperties bodyProperties57 = new A.BodyProperties();
            A.ListStyle listStyle57 = new A.ListStyle();

            A.Paragraph paragraph68 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties45 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph68.Append(endParagraphRunProperties45);

            textBody56.Append(bodyProperties57);
            textBody56.Append(listStyle57);
            textBody56.Append(paragraph68);

            shape36.Append(nonVisualShapeProperties36);
            shape36.Append(shapeProperties62);
            shape36.Append(textBody56);

            Shape shape37 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties37 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties81 = new NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Slide Number Placeholder 6" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList64 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension64 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement90 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{E4990AC2-5D21-2FC8-A84E-2FF389E6F2AC}\" />");

            // nonVisualDrawingPropertiesExtension64.Append(openXmlUnknownElement90);

            nonVisualDrawingPropertiesExtensionList64.Append(nonVisualDrawingPropertiesExtension64);

            nonVisualDrawingProperties81.Append(nonVisualDrawingPropertiesExtensionList64);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties37 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks20 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties37.Append(shapeLocks20);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties81 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape20 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties81.Append(placeholderShape20);

            nonVisualShapeProperties37.Append(nonVisualDrawingProperties81);
            nonVisualShapeProperties37.Append(nonVisualShapeDrawingProperties37);
            nonVisualShapeProperties37.Append(applicationNonVisualDrawingProperties81);
            ShapeProperties shapeProperties63 = new ShapeProperties();

            TextBody textBody57 = new TextBody();
            A.BodyProperties bodyProperties58 = new A.BodyProperties();
            A.ListStyle listStyle58 = new A.ListStyle();

            A.Paragraph paragraph69 = new A.Paragraph();

            A.Field field7 = new A.Field() { Id = "{19E6023F-2333-4A80-AFF1-A7F70B52B230}", Type = "slidenum" };

            A.RunProperties runProperties51 = new A.RunProperties() { Language = "en-US" };
            runProperties51.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text51 = new A.Text();
            text51.Text = "‹#›";

            field7.Append(runProperties51);
            field7.Append(text51);
            A.EndParagraphRunProperties endParagraphRunProperties46 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph69.Append(field7);
            paragraph69.Append(endParagraphRunProperties46);

            textBody57.Append(bodyProperties58);
            textBody57.Append(listStyle58);
            textBody57.Append(paragraph69);

            shape37.Append(nonVisualShapeProperties37);
            shape37.Append(shapeProperties63);
            shape37.Append(textBody57);

            shapeTree6.Append(nonVisualGroupShapeProperties18);
            shapeTree6.Append(groupShapeProperties18);
            shapeTree6.Append(shape32);
            shapeTree6.Append(shape33);
            shapeTree6.Append(shape34);
            shapeTree6.Append(shape35);
            shapeTree6.Append(shape36);
            shapeTree6.Append(shape37);

            CommonSlideDataExtensionList commonSlideDataExtensionList6 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension6 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId6 = new P14.CreationId() { Val = (UInt32Value)3077090653U };
            creationId6.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension6.Append(creationId6);

            commonSlideDataExtensionList6.Append(commonSlideDataExtension6);

            commonSlideData6.Append(shapeTree6);
            commonSlideData6.Append(commonSlideDataExtensionList6);

            ColorMapOverride colorMapOverride4 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping4 = new A.MasterColorMapping();

            colorMapOverride4.Append(masterColorMapping4);

            slideLayout2.Append(commonSlideData6);
            slideLayout2.Append(colorMapOverride4);

            slideLayoutPart2.SlideLayout = slideLayout2;
        }

        // Generates content of themePart2.
        private void GenerateThemePart2Content(ThemePart themePart2)
        {
            A.Theme theme2 = new A.Theme() { Name = "Office Theme" };
            theme2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements2 = new A.ThemeElements();

            A.ColorScheme colorScheme2 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color2 = new A.Dark1Color();
            A.SystemColor systemColor3 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color2.Append(systemColor3);

            A.Light1Color light1Color2 = new A.Light1Color();
            A.SystemColor systemColor4 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color2.Append(systemColor4);

            A.Dark2Color dark2Color2 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex39 = new A.RgbColorModelHex() { Val = "0E2841" };

            dark2Color2.Append(rgbColorModelHex39);

            A.Light2Color light2Color2 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex40 = new A.RgbColorModelHex() { Val = "E8E8E8" };

            light2Color2.Append(rgbColorModelHex40);

            A.Accent1Color accent1Color2 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex41 = new A.RgbColorModelHex() { Val = "156082" };

            accent1Color2.Append(rgbColorModelHex41);

            A.Accent2Color accent2Color2 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex42 = new A.RgbColorModelHex() { Val = "E97132" };

            accent2Color2.Append(rgbColorModelHex42);

            A.Accent3Color accent3Color2 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex43 = new A.RgbColorModelHex() { Val = "196B24" };

            accent3Color2.Append(rgbColorModelHex43);

            A.Accent4Color accent4Color2 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex44 = new A.RgbColorModelHex() { Val = "0F9ED5" };

            accent4Color2.Append(rgbColorModelHex44);

            A.Accent5Color accent5Color2 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex45 = new A.RgbColorModelHex() { Val = "A02B93" };

            accent5Color2.Append(rgbColorModelHex45);

            A.Accent6Color accent6Color2 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex46 = new A.RgbColorModelHex() { Val = "4EA72E" };

            accent6Color2.Append(rgbColorModelHex46);

            A.Hyperlink hyperlink2 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex47 = new A.RgbColorModelHex() { Val = "467886" };

            hyperlink2.Append(rgbColorModelHex47);

            A.FollowedHyperlinkColor followedHyperlinkColor2 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex48 = new A.RgbColorModelHex() { Val = "96607D" };

            followedHyperlinkColor2.Append(rgbColorModelHex48);

            colorScheme2.Append(dark1Color2);
            colorScheme2.Append(light1Color2);
            colorScheme2.Append(dark2Color2);
            colorScheme2.Append(light2Color2);
            colorScheme2.Append(accent1Color2);
            colorScheme2.Append(accent2Color2);
            colorScheme2.Append(accent3Color2);
            colorScheme2.Append(accent4Color2);
            colorScheme2.Append(accent5Color2);
            colorScheme2.Append(accent6Color2);
            colorScheme2.Append(hyperlink2);
            colorScheme2.Append(followedHyperlinkColor2);

            A.FontScheme fontScheme2 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont2 = new A.MajorFont();
            A.LatinFont latinFont82 = new A.LatinFont() { Typeface = "Aptos Display", Panose = "02110004020202020204" };
            A.EastAsianFont eastAsianFont72 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont82 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont95 = new A.SupplementalFont() { Script = "Jpan", Typeface = "游ゴシック Light" };
            A.SupplementalFont supplementalFont96 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont97 = new A.SupplementalFont() { Script = "Hans", Typeface = "等线 Light" };
            A.SupplementalFont supplementalFont98 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont99 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont100 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont101 = new A.SupplementalFont() { Script = "Thai", Typeface = "Angsana New" };
            A.SupplementalFont supplementalFont102 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont103 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont104 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont105 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont106 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont107 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont108 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont109 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont110 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont111 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont112 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont113 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont114 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont115 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont116 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont117 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont118 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont119 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont120 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont121 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont122 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont123 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont124 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };
            A.SupplementalFont supplementalFont125 = new A.SupplementalFont() { Script = "Armn", Typeface = "Arial" };
            A.SupplementalFont supplementalFont126 = new A.SupplementalFont() { Script = "Bugi", Typeface = "Leelawadee UI" };
            A.SupplementalFont supplementalFont127 = new A.SupplementalFont() { Script = "Bopo", Typeface = "Microsoft JhengHei" };
            A.SupplementalFont supplementalFont128 = new A.SupplementalFont() { Script = "Java", Typeface = "Javanese Text" };
            A.SupplementalFont supplementalFont129 = new A.SupplementalFont() { Script = "Lisu", Typeface = "Segoe UI" };
            A.SupplementalFont supplementalFont130 = new A.SupplementalFont() { Script = "Mymr", Typeface = "Myanmar Text" };
            A.SupplementalFont supplementalFont131 = new A.SupplementalFont() { Script = "Nkoo", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont132 = new A.SupplementalFont() { Script = "Olck", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont133 = new A.SupplementalFont() { Script = "Osma", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont134 = new A.SupplementalFont() { Script = "Phag", Typeface = "Phagspa" };
            A.SupplementalFont supplementalFont135 = new A.SupplementalFont() { Script = "Syrn", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont136 = new A.SupplementalFont() { Script = "Syrj", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont137 = new A.SupplementalFont() { Script = "Syre", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont138 = new A.SupplementalFont() { Script = "Sora", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont139 = new A.SupplementalFont() { Script = "Tale", Typeface = "Microsoft Tai Le" };
            A.SupplementalFont supplementalFont140 = new A.SupplementalFont() { Script = "Talu", Typeface = "Microsoft New Tai Lue" };
            A.SupplementalFont supplementalFont141 = new A.SupplementalFont() { Script = "Tfng", Typeface = "Ebrima" };

            majorFont2.Append(latinFont82);
            majorFont2.Append(eastAsianFont72);
            majorFont2.Append(complexScriptFont82);
            majorFont2.Append(supplementalFont95);
            majorFont2.Append(supplementalFont96);
            majorFont2.Append(supplementalFont97);
            majorFont2.Append(supplementalFont98);
            majorFont2.Append(supplementalFont99);
            majorFont2.Append(supplementalFont100);
            majorFont2.Append(supplementalFont101);
            majorFont2.Append(supplementalFont102);
            majorFont2.Append(supplementalFont103);
            majorFont2.Append(supplementalFont104);
            majorFont2.Append(supplementalFont105);
            majorFont2.Append(supplementalFont106);
            majorFont2.Append(supplementalFont107);
            majorFont2.Append(supplementalFont108);
            majorFont2.Append(supplementalFont109);
            majorFont2.Append(supplementalFont110);
            majorFont2.Append(supplementalFont111);
            majorFont2.Append(supplementalFont112);
            majorFont2.Append(supplementalFont113);
            majorFont2.Append(supplementalFont114);
            majorFont2.Append(supplementalFont115);
            majorFont2.Append(supplementalFont116);
            majorFont2.Append(supplementalFont117);
            majorFont2.Append(supplementalFont118);
            majorFont2.Append(supplementalFont119);
            majorFont2.Append(supplementalFont120);
            majorFont2.Append(supplementalFont121);
            majorFont2.Append(supplementalFont122);
            majorFont2.Append(supplementalFont123);
            majorFont2.Append(supplementalFont124);
            majorFont2.Append(supplementalFont125);
            majorFont2.Append(supplementalFont126);
            majorFont2.Append(supplementalFont127);
            majorFont2.Append(supplementalFont128);
            majorFont2.Append(supplementalFont129);
            majorFont2.Append(supplementalFont130);
            majorFont2.Append(supplementalFont131);
            majorFont2.Append(supplementalFont132);
            majorFont2.Append(supplementalFont133);
            majorFont2.Append(supplementalFont134);
            majorFont2.Append(supplementalFont135);
            majorFont2.Append(supplementalFont136);
            majorFont2.Append(supplementalFont137);
            majorFont2.Append(supplementalFont138);
            majorFont2.Append(supplementalFont139);
            majorFont2.Append(supplementalFont140);
            majorFont2.Append(supplementalFont141);

            A.MinorFont minorFont2 = new A.MinorFont();
            A.LatinFont latinFont83 = new A.LatinFont() { Typeface = "Aptos", Panose = "02110004020202020204" };
            A.EastAsianFont eastAsianFont73 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont83 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont142 = new A.SupplementalFont() { Script = "Jpan", Typeface = "游ゴシック" };
            A.SupplementalFont supplementalFont143 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont144 = new A.SupplementalFont() { Script = "Hans", Typeface = "等线" };
            A.SupplementalFont supplementalFont145 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont146 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont147 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont148 = new A.SupplementalFont() { Script = "Thai", Typeface = "Cordia New" };
            A.SupplementalFont supplementalFont149 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont150 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont151 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont152 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont153 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont154 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont155 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont156 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont157 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont158 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont159 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont160 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont161 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont162 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont163 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont164 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont165 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont166 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont167 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont168 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont169 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont170 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont171 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };
            A.SupplementalFont supplementalFont172 = new A.SupplementalFont() { Script = "Armn", Typeface = "Arial" };
            A.SupplementalFont supplementalFont173 = new A.SupplementalFont() { Script = "Bugi", Typeface = "Leelawadee UI" };
            A.SupplementalFont supplementalFont174 = new A.SupplementalFont() { Script = "Bopo", Typeface = "Microsoft JhengHei" };
            A.SupplementalFont supplementalFont175 = new A.SupplementalFont() { Script = "Java", Typeface = "Javanese Text" };
            A.SupplementalFont supplementalFont176 = new A.SupplementalFont() { Script = "Lisu", Typeface = "Segoe UI" };
            A.SupplementalFont supplementalFont177 = new A.SupplementalFont() { Script = "Mymr", Typeface = "Myanmar Text" };
            A.SupplementalFont supplementalFont178 = new A.SupplementalFont() { Script = "Nkoo", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont179 = new A.SupplementalFont() { Script = "Olck", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont180 = new A.SupplementalFont() { Script = "Osma", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont181 = new A.SupplementalFont() { Script = "Phag", Typeface = "Phagspa" };
            A.SupplementalFont supplementalFont182 = new A.SupplementalFont() { Script = "Syrn", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont183 = new A.SupplementalFont() { Script = "Syrj", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont184 = new A.SupplementalFont() { Script = "Syre", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont185 = new A.SupplementalFont() { Script = "Sora", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont186 = new A.SupplementalFont() { Script = "Tale", Typeface = "Microsoft Tai Le" };
            A.SupplementalFont supplementalFont187 = new A.SupplementalFont() { Script = "Talu", Typeface = "Microsoft New Tai Lue" };
            A.SupplementalFont supplementalFont188 = new A.SupplementalFont() { Script = "Tfng", Typeface = "Ebrima" };

            minorFont2.Append(latinFont83);
            minorFont2.Append(eastAsianFont73);
            minorFont2.Append(complexScriptFont83);
            minorFont2.Append(supplementalFont142);
            minorFont2.Append(supplementalFont143);
            minorFont2.Append(supplementalFont144);
            minorFont2.Append(supplementalFont145);
            minorFont2.Append(supplementalFont146);
            minorFont2.Append(supplementalFont147);
            minorFont2.Append(supplementalFont148);
            minorFont2.Append(supplementalFont149);
            minorFont2.Append(supplementalFont150);
            minorFont2.Append(supplementalFont151);
            minorFont2.Append(supplementalFont152);
            minorFont2.Append(supplementalFont153);
            minorFont2.Append(supplementalFont154);
            minorFont2.Append(supplementalFont155);
            minorFont2.Append(supplementalFont156);
            minorFont2.Append(supplementalFont157);
            minorFont2.Append(supplementalFont158);
            minorFont2.Append(supplementalFont159);
            minorFont2.Append(supplementalFont160);
            minorFont2.Append(supplementalFont161);
            minorFont2.Append(supplementalFont162);
            minorFont2.Append(supplementalFont163);
            minorFont2.Append(supplementalFont164);
            minorFont2.Append(supplementalFont165);
            minorFont2.Append(supplementalFont166);
            minorFont2.Append(supplementalFont167);
            minorFont2.Append(supplementalFont168);
            minorFont2.Append(supplementalFont169);
            minorFont2.Append(supplementalFont170);
            minorFont2.Append(supplementalFont171);
            minorFont2.Append(supplementalFont172);
            minorFont2.Append(supplementalFont173);
            minorFont2.Append(supplementalFont174);
            minorFont2.Append(supplementalFont175);
            minorFont2.Append(supplementalFont176);
            minorFont2.Append(supplementalFont177);
            minorFont2.Append(supplementalFont178);
            minorFont2.Append(supplementalFont179);
            minorFont2.Append(supplementalFont180);
            minorFont2.Append(supplementalFont181);
            minorFont2.Append(supplementalFont182);
            minorFont2.Append(supplementalFont183);
            minorFont2.Append(supplementalFont184);
            minorFont2.Append(supplementalFont185);
            minorFont2.Append(supplementalFont186);
            minorFont2.Append(supplementalFont187);
            minorFont2.Append(supplementalFont188);

            fontScheme2.Append(majorFont2);
            fontScheme2.Append(minorFont2);

            A.FormatScheme formatScheme2 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList2 = new A.FillStyleList();

            A.SolidFill solidFill108 = new A.SolidFill();
            A.SchemeColor schemeColor97 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill108.Append(schemeColor97);

            A.GradientFill gradientFill4 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop10 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor98 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation12 = new A.LuminanceModulation() { Val = 110000 };
            A.SaturationModulation saturationModulation11 = new A.SaturationModulation() { Val = 105000 };
            A.Tint tint11 = new A.Tint() { Val = 67000 };

            schemeColor98.Append(luminanceModulation12);
            schemeColor98.Append(saturationModulation11);
            schemeColor98.Append(tint11);

            gradientStop10.Append(schemeColor98);

            A.GradientStop gradientStop11 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor99 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation13 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation12 = new A.SaturationModulation() { Val = 103000 };
            A.Tint tint12 = new A.Tint() { Val = 73000 };

            schemeColor99.Append(luminanceModulation13);
            schemeColor99.Append(saturationModulation12);
            schemeColor99.Append(tint12);

            gradientStop11.Append(schemeColor99);

            A.GradientStop gradientStop12 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor100 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation14 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation13 = new A.SaturationModulation() { Val = 109000 };
            A.Tint tint13 = new A.Tint() { Val = 81000 };

            schemeColor100.Append(luminanceModulation14);
            schemeColor100.Append(saturationModulation13);
            schemeColor100.Append(tint13);

            gradientStop12.Append(schemeColor100);

            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);
            gradientStopList4.Append(gradientStop12);
            A.LinearGradientFill linearGradientFill4 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(linearGradientFill4);

            A.GradientFill gradientFill5 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList5 = new A.GradientStopList();

            A.GradientStop gradientStop13 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor101 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation14 = new A.SaturationModulation() { Val = 103000 };
            A.LuminanceModulation luminanceModulation15 = new A.LuminanceModulation() { Val = 102000 };
            A.Tint tint14 = new A.Tint() { Val = 94000 };

            schemeColor101.Append(saturationModulation14);
            schemeColor101.Append(luminanceModulation15);
            schemeColor101.Append(tint14);

            gradientStop13.Append(schemeColor101);

            A.GradientStop gradientStop14 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor102 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation15 = new A.SaturationModulation() { Val = 110000 };
            A.LuminanceModulation luminanceModulation16 = new A.LuminanceModulation() { Val = 100000 };
            A.Shade shade8 = new A.Shade() { Val = 100000 };

            schemeColor102.Append(saturationModulation15);
            schemeColor102.Append(luminanceModulation16);
            schemeColor102.Append(shade8);

            gradientStop14.Append(schemeColor102);

            A.GradientStop gradientStop15 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor103 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation17 = new A.LuminanceModulation() { Val = 99000 };
            A.SaturationModulation saturationModulation16 = new A.SaturationModulation() { Val = 120000 };
            A.Shade shade9 = new A.Shade() { Val = 78000 };

            schemeColor103.Append(luminanceModulation17);
            schemeColor103.Append(saturationModulation16);
            schemeColor103.Append(shade9);

            gradientStop15.Append(schemeColor103);

            gradientStopList5.Append(gradientStop13);
            gradientStopList5.Append(gradientStop14);
            gradientStopList5.Append(gradientStop15);
            A.LinearGradientFill linearGradientFill5 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill5.Append(gradientStopList5);
            gradientFill5.Append(linearGradientFill5);

            fillStyleList2.Append(solidFill108);
            fillStyleList2.Append(gradientFill4);
            fillStyleList2.Append(gradientFill5);

            A.LineStyleList lineStyleList2 = new A.LineStyleList();

            A.Outline outline36 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill109 = new A.SolidFill();
            A.SchemeColor schemeColor104 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill109.Append(schemeColor104);
            A.PresetDash presetDash85 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter4 = new A.Miter() { Limit = 800000 };

            outline36.Append(solidFill109);
            outline36.Append(presetDash85);
            outline36.Append(miter4);

            A.Outline outline37 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill110 = new A.SolidFill();
            A.SchemeColor schemeColor105 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill110.Append(schemeColor105);
            A.PresetDash presetDash86 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter5 = new A.Miter() { Limit = 800000 };

            outline37.Append(solidFill110);
            outline37.Append(presetDash86);
            outline37.Append(miter5);

            A.Outline outline38 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill111 = new A.SolidFill();
            A.SchemeColor schemeColor106 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill111.Append(schemeColor106);
            A.PresetDash presetDash87 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter6 = new A.Miter() { Limit = 800000 };

            outline38.Append(solidFill111);
            outline38.Append(presetDash87);
            outline38.Append(miter6);

            lineStyleList2.Append(outline36);
            lineStyleList2.Append(outline37);
            lineStyleList2.Append(outline38);

            A.EffectStyleList effectStyleList2 = new A.EffectStyleList();

            A.EffectStyle effectStyle4 = new A.EffectStyle();
            A.EffectList effectList26 = new A.EffectList();

            effectStyle4.Append(effectList26);

            A.EffectStyle effectStyle5 = new A.EffectStyle();
            A.EffectList effectList27 = new A.EffectList();

            effectStyle5.Append(effectList27);

            A.EffectStyle effectStyle6 = new A.EffectStyle();

            A.EffectList effectList28 = new A.EffectList();

            A.OuterShadow outerShadow3 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex49 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha17 = new A.Alpha() { Val = 63000 };

            rgbColorModelHex49.Append(alpha17);

            outerShadow3.Append(rgbColorModelHex49);

            effectList28.Append(outerShadow3);

            effectStyle6.Append(effectList28);

            effectStyleList2.Append(effectStyle4);
            effectStyleList2.Append(effectStyle5);
            effectStyleList2.Append(effectStyle6);

            A.BackgroundFillStyleList backgroundFillStyleList2 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill112 = new A.SolidFill();
            A.SchemeColor schemeColor107 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill112.Append(schemeColor107);

            A.SolidFill solidFill113 = new A.SolidFill();

            A.SchemeColor schemeColor108 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint15 = new A.Tint() { Val = 95000 };
            A.SaturationModulation saturationModulation17 = new A.SaturationModulation() { Val = 170000 };

            schemeColor108.Append(tint15);
            schemeColor108.Append(saturationModulation17);

            solidFill113.Append(schemeColor108);

            A.GradientFill gradientFill6 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList6 = new A.GradientStopList();

            A.GradientStop gradientStop16 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor109 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint16 = new A.Tint() { Val = 93000 };
            A.SaturationModulation saturationModulation18 = new A.SaturationModulation() { Val = 150000 };
            A.Shade shade10 = new A.Shade() { Val = 98000 };
            A.LuminanceModulation luminanceModulation18 = new A.LuminanceModulation() { Val = 102000 };

            schemeColor109.Append(tint16);
            schemeColor109.Append(saturationModulation18);
            schemeColor109.Append(shade10);
            schemeColor109.Append(luminanceModulation18);

            gradientStop16.Append(schemeColor109);

            A.GradientStop gradientStop17 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor110 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint17 = new A.Tint() { Val = 98000 };
            A.SaturationModulation saturationModulation19 = new A.SaturationModulation() { Val = 130000 };
            A.Shade shade11 = new A.Shade() { Val = 90000 };
            A.LuminanceModulation luminanceModulation19 = new A.LuminanceModulation() { Val = 103000 };

            schemeColor110.Append(tint17);
            schemeColor110.Append(saturationModulation19);
            schemeColor110.Append(shade11);
            schemeColor110.Append(luminanceModulation19);

            gradientStop17.Append(schemeColor110);

            A.GradientStop gradientStop18 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor111 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade12 = new A.Shade() { Val = 63000 };
            A.SaturationModulation saturationModulation20 = new A.SaturationModulation() { Val = 120000 };

            schemeColor111.Append(shade12);
            schemeColor111.Append(saturationModulation20);

            gradientStop18.Append(schemeColor111);

            gradientStopList6.Append(gradientStop16);
            gradientStopList6.Append(gradientStop17);
            gradientStopList6.Append(gradientStop18);
            A.LinearGradientFill linearGradientFill6 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill6.Append(gradientStopList6);
            gradientFill6.Append(linearGradientFill6);

            backgroundFillStyleList2.Append(solidFill112);
            backgroundFillStyleList2.Append(solidFill113);
            backgroundFillStyleList2.Append(gradientFill6);

            formatScheme2.Append(fillStyleList2);
            formatScheme2.Append(lineStyleList2);
            formatScheme2.Append(effectStyleList2);
            formatScheme2.Append(backgroundFillStyleList2);

            themeElements2.Append(colorScheme2);
            themeElements2.Append(fontScheme2);
            themeElements2.Append(formatScheme2);

            A.ObjectDefaults objectDefaults2 = new A.ObjectDefaults();

            A.LineDefault lineDefault2 = new A.LineDefault();
            A.ShapeProperties shapeProperties64 = new A.ShapeProperties();
            A.BodyProperties bodyProperties59 = new A.BodyProperties();
            A.ListStyle listStyle59 = new A.ListStyle();

            A.ShapeStyle shapeStyle4 = new A.ShapeStyle();

            A.LineReference lineReference4 = new A.LineReference() { Index = (UInt32Value)2U };
            A.SchemeColor schemeColor112 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            lineReference4.Append(schemeColor112);

            A.FillReference fillReference4 = new A.FillReference() { Index = (UInt32Value)0U };
            A.SchemeColor schemeColor113 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            fillReference4.Append(schemeColor113);

            A.EffectReference effectReference4 = new A.EffectReference() { Index = (UInt32Value)1U };
            A.SchemeColor schemeColor114 = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

            effectReference4.Append(schemeColor114);

            A.FontReference fontReference4 = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
            A.SchemeColor schemeColor115 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

            fontReference4.Append(schemeColor115);

            shapeStyle4.Append(lineReference4);
            shapeStyle4.Append(fillReference4);
            shapeStyle4.Append(effectReference4);
            shapeStyle4.Append(fontReference4);

            lineDefault2.Append(shapeProperties64);
            lineDefault2.Append(bodyProperties59);
            lineDefault2.Append(listStyle59);
            lineDefault2.Append(shapeStyle4);

            objectDefaults2.Append(lineDefault2);
            A.ExtraColorSchemeList extraColorSchemeList2 = new A.ExtraColorSchemeList();

            A.OfficeStyleSheetExtensionList officeStyleSheetExtensionList2 = new A.OfficeStyleSheetExtensionList();

            A.OfficeStyleSheetExtension officeStyleSheetExtension2 = new A.OfficeStyleSheetExtension() { Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

            Thm15.ThemeFamily themeFamily2 = new Thm15.ThemeFamily() { Name = "Office Theme", Id = "{2E142A2C-CD16-42D6-873A-C26D2A0506FA}", Vid = "{1BDDFF52-6CD6-40A5-AB3C-68EB2F1E4D0A}" };
            themeFamily2.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

            officeStyleSheetExtension2.Append(themeFamily2);

            officeStyleSheetExtensionList2.Append(officeStyleSheetExtension2);

            theme2.Append(themeElements2);
            theme2.Append(objectDefaults2);
            theme2.Append(extraColorSchemeList2);
            theme2.Append(officeStyleSheetExtensionList2);

            themePart2.Theme = theme2;
        }

        // Generates content of slideLayoutPart3.
        private void GenerateSlideLayoutPart3Content(SlideLayoutPart slideLayoutPart3)
        {
            SlideLayout slideLayout3 = new SlideLayout() { Type = SlideLayoutValues.SectionHeader, Preserve = true };
            slideLayout3.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout3.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData7 = new CommonSlideData() { Name = "Section Header" };

            ShapeTree shapeTree7 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties19 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties82 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties19 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties82 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties19.Append(nonVisualDrawingProperties82);
            nonVisualGroupShapeProperties19.Append(nonVisualGroupShapeDrawingProperties19);
            nonVisualGroupShapeProperties19.Append(applicationNonVisualDrawingProperties82);

            GroupShapeProperties groupShapeProperties19 = new GroupShapeProperties();

            A.TransformGroup transformGroup19 = new A.TransformGroup();
            A.Offset offset76 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents76 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset19 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents19 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup19.Append(offset76);
            transformGroup19.Append(extents76);
            transformGroup19.Append(childOffset19);
            transformGroup19.Append(childExtents19);

            groupShapeProperties19.Append(transformGroup19);

            Shape shape38 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties38 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties83 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList65 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension65 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement91 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{06FD457A-9158-68DA-5754-A276B3701B0A}\" />");

            // nonVisualDrawingPropertiesExtension65.Append(openXmlUnknownElement91);

            nonVisualDrawingPropertiesExtensionList65.Append(nonVisualDrawingPropertiesExtension65);

            nonVisualDrawingProperties83.Append(nonVisualDrawingPropertiesExtensionList65);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties38 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks21 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties38.Append(shapeLocks21);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties83 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape21 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties83.Append(placeholderShape21);

            nonVisualShapeProperties38.Append(nonVisualDrawingProperties83);
            nonVisualShapeProperties38.Append(nonVisualShapeDrawingProperties38);
            nonVisualShapeProperties38.Append(applicationNonVisualDrawingProperties83);

            ShapeProperties shapeProperties65 = new ShapeProperties();

            A.Transform2D transform2D57 = new A.Transform2D();
            A.Offset offset77 = new A.Offset() { X = 831850L, Y = 1709738L };
            A.Extents extents77 = new A.Extents() { Cx = 10515600L, Cy = 2852737L };

            transform2D57.Append(offset77);
            transform2D57.Append(extents77);

            shapeProperties65.Append(transform2D57);

            TextBody textBody58 = new TextBody();
            A.BodyProperties bodyProperties60 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle60 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties16 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties81 = new A.DefaultRunProperties() { FontSize = 6000 };

            level1ParagraphProperties16.Append(defaultRunProperties81);

            listStyle60.Append(level1ParagraphProperties16);

            A.Paragraph paragraph70 = new A.Paragraph();

            A.Run run45 = new A.Run();
            A.RunProperties runProperties52 = new A.RunProperties() { Language = "en-US" };
            A.Text text52 = new A.Text();
            text52.Text = "Click to edit Master title style";

            run45.Append(runProperties52);
            run45.Append(text52);

            paragraph70.Append(run45);

            textBody58.Append(bodyProperties60);
            textBody58.Append(listStyle60);
            textBody58.Append(paragraph70);

            shape38.Append(nonVisualShapeProperties38);
            shape38.Append(shapeProperties65);
            shape38.Append(textBody58);

            Shape shape39 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties39 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties84 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Text Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList66 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension66 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement92 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{3B80F2A6-F746-E78B-494C-4F6013ADC9E3}\" />");

            // nonVisualDrawingPropertiesExtension66.Append(openXmlUnknownElement92);

            nonVisualDrawingPropertiesExtensionList66.Append(nonVisualDrawingPropertiesExtension66);

            nonVisualDrawingProperties84.Append(nonVisualDrawingPropertiesExtensionList66);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties39 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks22 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties39.Append(shapeLocks22);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties84 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape22 = new PlaceholderShape() { Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties84.Append(placeholderShape22);

            nonVisualShapeProperties39.Append(nonVisualDrawingProperties84);
            nonVisualShapeProperties39.Append(nonVisualShapeDrawingProperties39);
            nonVisualShapeProperties39.Append(applicationNonVisualDrawingProperties84);

            ShapeProperties shapeProperties66 = new ShapeProperties();

            A.Transform2D transform2D58 = new A.Transform2D();
            A.Offset offset78 = new A.Offset() { X = 831850L, Y = 4589463L };
            A.Extents extents78 = new A.Extents() { Cx = 10515600L, Cy = 1500187L };

            transform2D58.Append(offset78);
            transform2D58.Append(extents78);

            shapeProperties66.Append(transform2D58);

            TextBody textBody59 = new TextBody();
            A.BodyProperties bodyProperties61 = new A.BodyProperties();

            A.ListStyle listStyle61 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties17 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet25 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties82 = new A.DefaultRunProperties() { FontSize = 2400 };

            A.SolidFill solidFill114 = new A.SolidFill();

            A.SchemeColor schemeColor116 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint18 = new A.Tint() { Val = 82000 };

            schemeColor116.Append(tint18);

            solidFill114.Append(schemeColor116);

            defaultRunProperties82.Append(solidFill114);

            level1ParagraphProperties17.Append(noBullet25);
            level1ParagraphProperties17.Append(defaultRunProperties82);

            A.Level2ParagraphProperties level2ParagraphProperties7 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet26 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties83 = new A.DefaultRunProperties() { FontSize = 2000 };

            A.SolidFill solidFill115 = new A.SolidFill();

            A.SchemeColor schemeColor117 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint19 = new A.Tint() { Val = 82000 };

            schemeColor117.Append(tint19);

            solidFill115.Append(schemeColor117);

            defaultRunProperties83.Append(solidFill115);

            level2ParagraphProperties7.Append(noBullet26);
            level2ParagraphProperties7.Append(defaultRunProperties83);

            A.Level3ParagraphProperties level3ParagraphProperties7 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet27 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties84 = new A.DefaultRunProperties() { FontSize = 1800 };

            A.SolidFill solidFill116 = new A.SolidFill();

            A.SchemeColor schemeColor118 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint20 = new A.Tint() { Val = 82000 };

            schemeColor118.Append(tint20);

            solidFill116.Append(schemeColor118);

            defaultRunProperties84.Append(solidFill116);

            level3ParagraphProperties7.Append(noBullet27);
            level3ParagraphProperties7.Append(defaultRunProperties84);

            A.Level4ParagraphProperties level4ParagraphProperties7 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet28 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties85 = new A.DefaultRunProperties() { FontSize = 1600 };

            A.SolidFill solidFill117 = new A.SolidFill();

            A.SchemeColor schemeColor119 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint21 = new A.Tint() { Val = 82000 };

            schemeColor119.Append(tint21);

            solidFill117.Append(schemeColor119);

            defaultRunProperties85.Append(solidFill117);

            level4ParagraphProperties7.Append(noBullet28);
            level4ParagraphProperties7.Append(defaultRunProperties85);

            A.Level5ParagraphProperties level5ParagraphProperties7 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet29 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties86 = new A.DefaultRunProperties() { FontSize = 1600 };

            A.SolidFill solidFill118 = new A.SolidFill();

            A.SchemeColor schemeColor120 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint22 = new A.Tint() { Val = 82000 };

            schemeColor120.Append(tint22);

            solidFill118.Append(schemeColor120);

            defaultRunProperties86.Append(solidFill118);

            level5ParagraphProperties7.Append(noBullet29);
            level5ParagraphProperties7.Append(defaultRunProperties86);

            A.Level6ParagraphProperties level6ParagraphProperties7 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet30 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties87 = new A.DefaultRunProperties() { FontSize = 1600 };

            A.SolidFill solidFill119 = new A.SolidFill();

            A.SchemeColor schemeColor121 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint23 = new A.Tint() { Val = 82000 };

            schemeColor121.Append(tint23);

            solidFill119.Append(schemeColor121);

            defaultRunProperties87.Append(solidFill119);

            level6ParagraphProperties7.Append(noBullet30);
            level6ParagraphProperties7.Append(defaultRunProperties87);

            A.Level7ParagraphProperties level7ParagraphProperties7 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet31 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties88 = new A.DefaultRunProperties() { FontSize = 1600 };

            A.SolidFill solidFill120 = new A.SolidFill();

            A.SchemeColor schemeColor122 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint24 = new A.Tint() { Val = 82000 };

            schemeColor122.Append(tint24);

            solidFill120.Append(schemeColor122);

            defaultRunProperties88.Append(solidFill120);

            level7ParagraphProperties7.Append(noBullet31);
            level7ParagraphProperties7.Append(defaultRunProperties88);

            A.Level8ParagraphProperties level8ParagraphProperties7 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet32 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties89 = new A.DefaultRunProperties() { FontSize = 1600 };

            A.SolidFill solidFill121 = new A.SolidFill();

            A.SchemeColor schemeColor123 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint25 = new A.Tint() { Val = 82000 };

            schemeColor123.Append(tint25);

            solidFill121.Append(schemeColor123);

            defaultRunProperties89.Append(solidFill121);

            level8ParagraphProperties7.Append(noBullet32);
            level8ParagraphProperties7.Append(defaultRunProperties89);

            A.Level9ParagraphProperties level9ParagraphProperties7 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet33 = new A.NoBullet();

            A.DefaultRunProperties defaultRunProperties90 = new A.DefaultRunProperties() { FontSize = 1600 };

            A.SolidFill solidFill122 = new A.SolidFill();

            A.SchemeColor schemeColor124 = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };
            A.Tint tint26 = new A.Tint() { Val = 82000 };

            schemeColor124.Append(tint26);

            solidFill122.Append(schemeColor124);

            defaultRunProperties90.Append(solidFill122);

            level9ParagraphProperties7.Append(noBullet33);
            level9ParagraphProperties7.Append(defaultRunProperties90);

            listStyle61.Append(level1ParagraphProperties17);
            listStyle61.Append(level2ParagraphProperties7);
            listStyle61.Append(level3ParagraphProperties7);
            listStyle61.Append(level4ParagraphProperties7);
            listStyle61.Append(level5ParagraphProperties7);
            listStyle61.Append(level6ParagraphProperties7);
            listStyle61.Append(level7ParagraphProperties7);
            listStyle61.Append(level8ParagraphProperties7);
            listStyle61.Append(level9ParagraphProperties7);

            A.Paragraph paragraph71 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties51 = new A.ParagraphProperties() { Level = 0 };

            A.Run run46 = new A.Run();
            A.RunProperties runProperties53 = new A.RunProperties() { Language = "en-US" };
            A.Text text53 = new A.Text();
            text53.Text = "Click to edit Master text styles";

            run46.Append(runProperties53);
            run46.Append(text53);

            paragraph71.Append(paragraphProperties51);
            paragraph71.Append(run46);

            textBody59.Append(bodyProperties61);
            textBody59.Append(listStyle61);
            textBody59.Append(paragraph71);

            shape39.Append(nonVisualShapeProperties39);
            shape39.Append(shapeProperties66);
            shape39.Append(textBody59);

            Shape shape40 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties40 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties85 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList67 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension67 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement93 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{57DD22BD-7A0B-BF5B-94F1-BB8376C35E95}\" />");

            // nonVisualDrawingPropertiesExtension67.Append(openXmlUnknownElement93);

            nonVisualDrawingPropertiesExtensionList67.Append(nonVisualDrawingPropertiesExtension67);

            nonVisualDrawingProperties85.Append(nonVisualDrawingPropertiesExtensionList67);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties40 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks23 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties40.Append(shapeLocks23);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties85 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape23 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties85.Append(placeholderShape23);

            nonVisualShapeProperties40.Append(nonVisualDrawingProperties85);
            nonVisualShapeProperties40.Append(nonVisualShapeDrawingProperties40);
            nonVisualShapeProperties40.Append(applicationNonVisualDrawingProperties85);
            ShapeProperties shapeProperties67 = new ShapeProperties();

            TextBody textBody60 = new TextBody();
            A.BodyProperties bodyProperties62 = new A.BodyProperties();
            A.ListStyle listStyle62 = new A.ListStyle();

            A.Paragraph paragraph72 = new A.Paragraph();

            A.Field field8 = new A.Field() { Id = "{B5CF329D-DBF6-4E15-9822-D0979B6928D0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties54 = new A.RunProperties() { Language = "en-US" };
            runProperties54.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text54 = new A.Text();
            text54.Text = "4/24/2024";

            field8.Append(runProperties54);
            field8.Append(text54);
            A.EndParagraphRunProperties endParagraphRunProperties47 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph72.Append(field8);
            paragraph72.Append(endParagraphRunProperties47);

            textBody60.Append(bodyProperties62);
            textBody60.Append(listStyle62);
            textBody60.Append(paragraph72);

            shape40.Append(nonVisualShapeProperties40);
            shape40.Append(shapeProperties67);
            shape40.Append(textBody60);

            Shape shape41 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties41 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties86 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList68 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension68 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement94 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{C2C2AFEF-9BDB-F65D-D323-296941250BD3}\" />");

            // nonVisualDrawingPropertiesExtension68.Append(openXmlUnknownElement94);

            nonVisualDrawingPropertiesExtensionList68.Append(nonVisualDrawingPropertiesExtension68);

            nonVisualDrawingProperties86.Append(nonVisualDrawingPropertiesExtensionList68);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties41 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks24 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties41.Append(shapeLocks24);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties86 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape24 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties86.Append(placeholderShape24);

            nonVisualShapeProperties41.Append(nonVisualDrawingProperties86);
            nonVisualShapeProperties41.Append(nonVisualShapeDrawingProperties41);
            nonVisualShapeProperties41.Append(applicationNonVisualDrawingProperties86);
            ShapeProperties shapeProperties68 = new ShapeProperties();

            TextBody textBody61 = new TextBody();
            A.BodyProperties bodyProperties63 = new A.BodyProperties();
            A.ListStyle listStyle63 = new A.ListStyle();

            A.Paragraph paragraph73 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties48 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph73.Append(endParagraphRunProperties48);

            textBody61.Append(bodyProperties63);
            textBody61.Append(listStyle63);
            textBody61.Append(paragraph73);

            shape41.Append(nonVisualShapeProperties41);
            shape41.Append(shapeProperties68);
            shape41.Append(textBody61);

            Shape shape42 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties42 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties87 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList69 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension69 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement95 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{BF65F736-363B-4A5A-B931-2F8B3E07A3CE}\" />");

            // nonVisualDrawingPropertiesExtension69.Append(openXmlUnknownElement95);

            nonVisualDrawingPropertiesExtensionList69.Append(nonVisualDrawingPropertiesExtension69);

            nonVisualDrawingProperties87.Append(nonVisualDrawingPropertiesExtensionList69);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties42 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks25 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties42.Append(shapeLocks25);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties87 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape25 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties87.Append(placeholderShape25);

            nonVisualShapeProperties42.Append(nonVisualDrawingProperties87);
            nonVisualShapeProperties42.Append(nonVisualShapeDrawingProperties42);
            nonVisualShapeProperties42.Append(applicationNonVisualDrawingProperties87);
            ShapeProperties shapeProperties69 = new ShapeProperties();

            TextBody textBody62 = new TextBody();
            A.BodyProperties bodyProperties64 = new A.BodyProperties();
            A.ListStyle listStyle64 = new A.ListStyle();

            A.Paragraph paragraph74 = new A.Paragraph();

            A.Field field9 = new A.Field() { Id = "{19E6023F-2333-4A80-AFF1-A7F70B52B230}", Type = "slidenum" };

            A.RunProperties runProperties55 = new A.RunProperties() { Language = "en-US" };
            runProperties55.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text55 = new A.Text();
            text55.Text = "‹#›";

            field9.Append(runProperties55);
            field9.Append(text55);
            A.EndParagraphRunProperties endParagraphRunProperties49 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph74.Append(field9);
            paragraph74.Append(endParagraphRunProperties49);

            textBody62.Append(bodyProperties64);
            textBody62.Append(listStyle64);
            textBody62.Append(paragraph74);

            shape42.Append(nonVisualShapeProperties42);
            shape42.Append(shapeProperties69);
            shape42.Append(textBody62);

            shapeTree7.Append(nonVisualGroupShapeProperties19);
            shapeTree7.Append(groupShapeProperties19);
            shapeTree7.Append(shape38);
            shapeTree7.Append(shape39);
            shapeTree7.Append(shape40);
            shapeTree7.Append(shape41);
            shapeTree7.Append(shape42);

            CommonSlideDataExtensionList commonSlideDataExtensionList7 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension7 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId7 = new P14.CreationId() { Val = (UInt32Value)1714439822U };
            creationId7.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension7.Append(creationId7);

            commonSlideDataExtensionList7.Append(commonSlideDataExtension7);

            commonSlideData7.Append(shapeTree7);
            commonSlideData7.Append(commonSlideDataExtensionList7);

            ColorMapOverride colorMapOverride5 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping5 = new A.MasterColorMapping();

            colorMapOverride5.Append(masterColorMapping5);

            slideLayout3.Append(commonSlideData7);
            slideLayout3.Append(colorMapOverride5);

            slideLayoutPart3.SlideLayout = slideLayout3;
        }

        // Generates content of slideLayoutPart4.
        private void GenerateSlideLayoutPart4Content(SlideLayoutPart slideLayoutPart4)
        {
            SlideLayout slideLayout4 = new SlideLayout() { Type = SlideLayoutValues.Blank, Preserve = true };
            slideLayout4.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout4.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout4.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData8 = new CommonSlideData() { Name = "Blank" };

            ShapeTree shapeTree8 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties20 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties88 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties20 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties88 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties20.Append(nonVisualDrawingProperties88);
            nonVisualGroupShapeProperties20.Append(nonVisualGroupShapeDrawingProperties20);
            nonVisualGroupShapeProperties20.Append(applicationNonVisualDrawingProperties88);

            GroupShapeProperties groupShapeProperties20 = new GroupShapeProperties();

            A.TransformGroup transformGroup20 = new A.TransformGroup();
            A.Offset offset79 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents79 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset20 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents20 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup20.Append(offset79);
            transformGroup20.Append(extents79);
            transformGroup20.Append(childOffset20);
            transformGroup20.Append(childExtents20);

            groupShapeProperties20.Append(transformGroup20);

            Shape shape43 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties43 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties89 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Date Placeholder 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList70 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension70 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement96 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{984F6B67-F93B-D8E0-C934-3555B195AA5A}\" />");

            // nonVisualDrawingPropertiesExtension70.Append(openXmlUnknownElement96);

            nonVisualDrawingPropertiesExtensionList70.Append(nonVisualDrawingPropertiesExtension70);

            nonVisualDrawingProperties89.Append(nonVisualDrawingPropertiesExtensionList70);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties43 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks26 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties43.Append(shapeLocks26);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties89 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape26 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties89.Append(placeholderShape26);

            nonVisualShapeProperties43.Append(nonVisualDrawingProperties89);
            nonVisualShapeProperties43.Append(nonVisualShapeDrawingProperties43);
            nonVisualShapeProperties43.Append(applicationNonVisualDrawingProperties89);
            ShapeProperties shapeProperties70 = new ShapeProperties();

            TextBody textBody63 = new TextBody();
            A.BodyProperties bodyProperties65 = new A.BodyProperties();
            A.ListStyle listStyle65 = new A.ListStyle();

            A.Paragraph paragraph75 = new A.Paragraph();

            A.Field field10 = new A.Field() { Id = "{B5CF329D-DBF6-4E15-9822-D0979B6928D0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties56 = new A.RunProperties() { Language = "en-US" };
            runProperties56.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text56 = new A.Text();
            text56.Text = "4/24/2024";

            field10.Append(runProperties56);
            field10.Append(text56);
            A.EndParagraphRunProperties endParagraphRunProperties50 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph75.Append(field10);
            paragraph75.Append(endParagraphRunProperties50);

            textBody63.Append(bodyProperties65);
            textBody63.Append(listStyle65);
            textBody63.Append(paragraph75);

            shape43.Append(nonVisualShapeProperties43);
            shape43.Append(shapeProperties70);
            shape43.Append(textBody63);

            Shape shape44 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties44 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties90 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Footer Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList71 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension71 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement97 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{ACFC260E-DF7D-09B0-8A84-2CFA93D57DC8}\" />");

            // nonVisualDrawingPropertiesExtension71.Append(openXmlUnknownElement97);

            nonVisualDrawingPropertiesExtensionList71.Append(nonVisualDrawingPropertiesExtension71);

            nonVisualDrawingProperties90.Append(nonVisualDrawingPropertiesExtensionList71);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties44 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks27 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties44.Append(shapeLocks27);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties90 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape27 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties90.Append(placeholderShape27);

            nonVisualShapeProperties44.Append(nonVisualDrawingProperties90);
            nonVisualShapeProperties44.Append(nonVisualShapeDrawingProperties44);
            nonVisualShapeProperties44.Append(applicationNonVisualDrawingProperties90);
            ShapeProperties shapeProperties71 = new ShapeProperties();

            TextBody textBody64 = new TextBody();
            A.BodyProperties bodyProperties66 = new A.BodyProperties();
            A.ListStyle listStyle66 = new A.ListStyle();

            A.Paragraph paragraph76 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties51 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph76.Append(endParagraphRunProperties51);

            textBody64.Append(bodyProperties66);
            textBody64.Append(listStyle66);
            textBody64.Append(paragraph76);

            shape44.Append(nonVisualShapeProperties44);
            shape44.Append(shapeProperties71);
            shape44.Append(textBody64);

            Shape shape45 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties45 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties91 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Slide Number Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList72 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension72 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement98 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{F24C5098-6AA5-2494-7A3E-AAEB7C814742}\" />");

            // nonVisualDrawingPropertiesExtension72.Append(openXmlUnknownElement98);

            nonVisualDrawingPropertiesExtensionList72.Append(nonVisualDrawingPropertiesExtension72);

            nonVisualDrawingProperties91.Append(nonVisualDrawingPropertiesExtensionList72);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties45 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks28 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties45.Append(shapeLocks28);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties91 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape28 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties91.Append(placeholderShape28);

            nonVisualShapeProperties45.Append(nonVisualDrawingProperties91);
            nonVisualShapeProperties45.Append(nonVisualShapeDrawingProperties45);
            nonVisualShapeProperties45.Append(applicationNonVisualDrawingProperties91);
            ShapeProperties shapeProperties72 = new ShapeProperties();

            TextBody textBody65 = new TextBody();
            A.BodyProperties bodyProperties67 = new A.BodyProperties();
            A.ListStyle listStyle67 = new A.ListStyle();

            A.Paragraph paragraph77 = new A.Paragraph();

            A.Field field11 = new A.Field() { Id = "{19E6023F-2333-4A80-AFF1-A7F70B52B230}", Type = "slidenum" };

            A.RunProperties runProperties57 = new A.RunProperties() { Language = "en-US" };
            runProperties57.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text57 = new A.Text();
            text57.Text = "‹#›";

            field11.Append(runProperties57);
            field11.Append(text57);
            A.EndParagraphRunProperties endParagraphRunProperties52 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph77.Append(field11);
            paragraph77.Append(endParagraphRunProperties52);

            textBody65.Append(bodyProperties67);
            textBody65.Append(listStyle67);
            textBody65.Append(paragraph77);

            shape45.Append(nonVisualShapeProperties45);
            shape45.Append(shapeProperties72);
            shape45.Append(textBody65);

            shapeTree8.Append(nonVisualGroupShapeProperties20);
            shapeTree8.Append(groupShapeProperties20);
            shapeTree8.Append(shape43);
            shapeTree8.Append(shape44);
            shapeTree8.Append(shape45);

            CommonSlideDataExtensionList commonSlideDataExtensionList8 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension8 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId8 = new P14.CreationId() { Val = (UInt32Value)1168424369U };
            creationId8.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension8.Append(creationId8);

            commonSlideDataExtensionList8.Append(commonSlideDataExtension8);

            commonSlideData8.Append(shapeTree8);
            commonSlideData8.Append(commonSlideDataExtensionList8);

            ColorMapOverride colorMapOverride6 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping6 = new A.MasterColorMapping();

            colorMapOverride6.Append(masterColorMapping6);

            slideLayout4.Append(commonSlideData8);
            slideLayout4.Append(colorMapOverride6);

            slideLayoutPart4.SlideLayout = slideLayout4;
        }

        // Generates content of slideLayoutPart5.
        private void GenerateSlideLayoutPart5Content(SlideLayoutPart slideLayoutPart5)
        {
            SlideLayout slideLayout5 = new SlideLayout() { Type = SlideLayoutValues.Object, Preserve = true };
            slideLayout5.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout5.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout5.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData9 = new CommonSlideData() { Name = "Title and Content" };

            ShapeTree shapeTree9 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties21 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties92 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties21 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties92 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties21.Append(nonVisualDrawingProperties92);
            nonVisualGroupShapeProperties21.Append(nonVisualGroupShapeDrawingProperties21);
            nonVisualGroupShapeProperties21.Append(applicationNonVisualDrawingProperties92);

            GroupShapeProperties groupShapeProperties21 = new GroupShapeProperties();

            A.TransformGroup transformGroup21 = new A.TransformGroup();
            A.Offset offset80 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents80 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset21 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents21 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup21.Append(offset80);
            transformGroup21.Append(extents80);
            transformGroup21.Append(childOffset21);
            transformGroup21.Append(childExtents21);

            groupShapeProperties21.Append(transformGroup21);

            Shape shape46 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties46 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties93 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList73 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension73 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement99 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{448509D4-3072-B41F-B24D-30E0575BAD2C}\" />");

            // nonVisualDrawingPropertiesExtension73.Append(openXmlUnknownElement99);

            nonVisualDrawingPropertiesExtensionList73.Append(nonVisualDrawingPropertiesExtension73);

            nonVisualDrawingProperties93.Append(nonVisualDrawingPropertiesExtensionList73);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties46 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks29 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties46.Append(shapeLocks29);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties93 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape29 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties93.Append(placeholderShape29);

            nonVisualShapeProperties46.Append(nonVisualDrawingProperties93);
            nonVisualShapeProperties46.Append(nonVisualShapeDrawingProperties46);
            nonVisualShapeProperties46.Append(applicationNonVisualDrawingProperties93);
            ShapeProperties shapeProperties73 = new ShapeProperties();

            TextBody textBody66 = new TextBody();
            A.BodyProperties bodyProperties68 = new A.BodyProperties();
            A.ListStyle listStyle68 = new A.ListStyle();

            A.Paragraph paragraph78 = new A.Paragraph();

            A.Run run47 = new A.Run();
            A.RunProperties runProperties58 = new A.RunProperties() { Language = "en-US" };
            A.Text text58 = new A.Text();
            text58.Text = "Click to edit Master title style";

            run47.Append(runProperties58);
            run47.Append(text58);

            paragraph78.Append(run47);

            textBody66.Append(bodyProperties68);
            textBody66.Append(listStyle68);
            textBody66.Append(paragraph78);

            shape46.Append(nonVisualShapeProperties46);
            shape46.Append(shapeProperties73);
            shape46.Append(textBody66);

            Shape shape47 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties47 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties94 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Content Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList74 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension74 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement100 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{FAE03E93-3EF8-69CC-8C19-DFE0FD77A09F}\" />");

            // nonVisualDrawingPropertiesExtension74.Append(openXmlUnknownElement100);

            nonVisualDrawingPropertiesExtensionList74.Append(nonVisualDrawingPropertiesExtension74);

            nonVisualDrawingProperties94.Append(nonVisualDrawingPropertiesExtensionList74);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties47 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks30 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties47.Append(shapeLocks30);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties94 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape30 = new PlaceholderShape() { Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties94.Append(placeholderShape30);

            nonVisualShapeProperties47.Append(nonVisualDrawingProperties94);
            nonVisualShapeProperties47.Append(nonVisualShapeDrawingProperties47);
            nonVisualShapeProperties47.Append(applicationNonVisualDrawingProperties94);
            ShapeProperties shapeProperties74 = new ShapeProperties();

            TextBody textBody67 = new TextBody();
            A.BodyProperties bodyProperties69 = new A.BodyProperties();
            A.ListStyle listStyle69 = new A.ListStyle();

            A.Paragraph paragraph79 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties52 = new A.ParagraphProperties() { Level = 0 };

            A.Run run48 = new A.Run();
            A.RunProperties runProperties59 = new A.RunProperties() { Language = "en-US" };
            A.Text text59 = new A.Text();
            text59.Text = "Click to edit Master text styles";

            run48.Append(runProperties59);
            run48.Append(text59);

            paragraph79.Append(paragraphProperties52);
            paragraph79.Append(run48);

            A.Paragraph paragraph80 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties53 = new A.ParagraphProperties() { Level = 1 };

            A.Run run49 = new A.Run();
            A.RunProperties runProperties60 = new A.RunProperties() { Language = "en-US" };
            A.Text text60 = new A.Text();
            text60.Text = "Second level";

            run49.Append(runProperties60);
            run49.Append(text60);

            paragraph80.Append(paragraphProperties53);
            paragraph80.Append(run49);

            A.Paragraph paragraph81 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties54 = new A.ParagraphProperties() { Level = 2 };

            A.Run run50 = new A.Run();
            A.RunProperties runProperties61 = new A.RunProperties() { Language = "en-US" };
            A.Text text61 = new A.Text();
            text61.Text = "Third level";

            run50.Append(runProperties61);
            run50.Append(text61);

            paragraph81.Append(paragraphProperties54);
            paragraph81.Append(run50);

            A.Paragraph paragraph82 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties55 = new A.ParagraphProperties() { Level = 3 };

            A.Run run51 = new A.Run();
            A.RunProperties runProperties62 = new A.RunProperties() { Language = "en-US" };
            A.Text text62 = new A.Text();
            text62.Text = "Fourth level";

            run51.Append(runProperties62);
            run51.Append(text62);

            paragraph82.Append(paragraphProperties55);
            paragraph82.Append(run51);

            A.Paragraph paragraph83 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties56 = new A.ParagraphProperties() { Level = 4 };

            A.Run run52 = new A.Run();
            A.RunProperties runProperties63 = new A.RunProperties() { Language = "en-US" };
            A.Text text63 = new A.Text();
            text63.Text = "Fifth level";

            run52.Append(runProperties63);
            run52.Append(text63);

            paragraph83.Append(paragraphProperties56);
            paragraph83.Append(run52);

            textBody67.Append(bodyProperties69);
            textBody67.Append(listStyle69);
            textBody67.Append(paragraph79);
            textBody67.Append(paragraph80);
            textBody67.Append(paragraph81);
            textBody67.Append(paragraph82);
            textBody67.Append(paragraph83);

            shape47.Append(nonVisualShapeProperties47);
            shape47.Append(shapeProperties74);
            shape47.Append(textBody67);

            Shape shape48 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties48 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties95 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList75 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension75 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement101 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{F88C09C0-22C2-4233-8415-93558FB2475E}\" />");

            // nonVisualDrawingPropertiesExtension75.Append(openXmlUnknownElement101);

            nonVisualDrawingPropertiesExtensionList75.Append(nonVisualDrawingPropertiesExtension75);

            nonVisualDrawingProperties95.Append(nonVisualDrawingPropertiesExtensionList75);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties48 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks31 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties48.Append(shapeLocks31);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties95 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape31 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties95.Append(placeholderShape31);

            nonVisualShapeProperties48.Append(nonVisualDrawingProperties95);
            nonVisualShapeProperties48.Append(nonVisualShapeDrawingProperties48);
            nonVisualShapeProperties48.Append(applicationNonVisualDrawingProperties95);
            ShapeProperties shapeProperties75 = new ShapeProperties();

            TextBody textBody68 = new TextBody();
            A.BodyProperties bodyProperties70 = new A.BodyProperties();
            A.ListStyle listStyle70 = new A.ListStyle();

            A.Paragraph paragraph84 = new A.Paragraph();

            A.Field field12 = new A.Field() { Id = "{B5CF329D-DBF6-4E15-9822-D0979B6928D0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties64 = new A.RunProperties() { Language = "en-US" };
            runProperties64.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text64 = new A.Text();
            text64.Text = "4/24/2024";

            field12.Append(runProperties64);
            field12.Append(text64);
            A.EndParagraphRunProperties endParagraphRunProperties53 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph84.Append(field12);
            paragraph84.Append(endParagraphRunProperties53);

            textBody68.Append(bodyProperties70);
            textBody68.Append(listStyle70);
            textBody68.Append(paragraph84);

            shape48.Append(nonVisualShapeProperties48);
            shape48.Append(shapeProperties75);
            shape48.Append(textBody68);

            Shape shape49 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties49 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties96 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList76 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension76 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement102 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{EBF4681E-AD3F-5B5C-A03D-D9FD62CDB51D}\" />");

            // nonVisualDrawingPropertiesExtension76.Append(openXmlUnknownElement102);

            nonVisualDrawingPropertiesExtensionList76.Append(nonVisualDrawingPropertiesExtension76);

            nonVisualDrawingProperties96.Append(nonVisualDrawingPropertiesExtensionList76);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties49 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks32 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties49.Append(shapeLocks32);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties96 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape32 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties96.Append(placeholderShape32);

            nonVisualShapeProperties49.Append(nonVisualDrawingProperties96);
            nonVisualShapeProperties49.Append(nonVisualShapeDrawingProperties49);
            nonVisualShapeProperties49.Append(applicationNonVisualDrawingProperties96);
            ShapeProperties shapeProperties76 = new ShapeProperties();

            TextBody textBody69 = new TextBody();
            A.BodyProperties bodyProperties71 = new A.BodyProperties();
            A.ListStyle listStyle71 = new A.ListStyle();

            A.Paragraph paragraph85 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties54 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph85.Append(endParagraphRunProperties54);

            textBody69.Append(bodyProperties71);
            textBody69.Append(listStyle71);
            textBody69.Append(paragraph85);

            shape49.Append(nonVisualShapeProperties49);
            shape49.Append(shapeProperties76);
            shape49.Append(textBody69);

            Shape shape50 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties50 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties97 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList77 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension77 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement103 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{F1853765-E93A-00E5-ED2D-FE21DC3D184A}\" />");

            // nonVisualDrawingPropertiesExtension77.Append(openXmlUnknownElement103);

            nonVisualDrawingPropertiesExtensionList77.Append(nonVisualDrawingPropertiesExtension77);

            nonVisualDrawingProperties97.Append(nonVisualDrawingPropertiesExtensionList77);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties50 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks33 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties50.Append(shapeLocks33);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties97 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape33 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties97.Append(placeholderShape33);

            nonVisualShapeProperties50.Append(nonVisualDrawingProperties97);
            nonVisualShapeProperties50.Append(nonVisualShapeDrawingProperties50);
            nonVisualShapeProperties50.Append(applicationNonVisualDrawingProperties97);
            ShapeProperties shapeProperties77 = new ShapeProperties();

            TextBody textBody70 = new TextBody();
            A.BodyProperties bodyProperties72 = new A.BodyProperties();
            A.ListStyle listStyle72 = new A.ListStyle();

            A.Paragraph paragraph86 = new A.Paragraph();

            A.Field field13 = new A.Field() { Id = "{19E6023F-2333-4A80-AFF1-A7F70B52B230}", Type = "slidenum" };

            A.RunProperties runProperties65 = new A.RunProperties() { Language = "en-US" };
            runProperties65.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text65 = new A.Text();
            text65.Text = "‹#›";

            field13.Append(runProperties65);
            field13.Append(text65);
            A.EndParagraphRunProperties endParagraphRunProperties55 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph86.Append(field13);
            paragraph86.Append(endParagraphRunProperties55);

            textBody70.Append(bodyProperties72);
            textBody70.Append(listStyle72);
            textBody70.Append(paragraph86);

            shape50.Append(nonVisualShapeProperties50);
            shape50.Append(shapeProperties77);
            shape50.Append(textBody70);

            shapeTree9.Append(nonVisualGroupShapeProperties21);
            shapeTree9.Append(groupShapeProperties21);
            shapeTree9.Append(shape46);
            shapeTree9.Append(shape47);
            shapeTree9.Append(shape48);
            shapeTree9.Append(shape49);
            shapeTree9.Append(shape50);

            CommonSlideDataExtensionList commonSlideDataExtensionList9 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension9 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId9 = new P14.CreationId() { Val = (UInt32Value)3944956704U };
            creationId9.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension9.Append(creationId9);

            commonSlideDataExtensionList9.Append(commonSlideDataExtension9);

            commonSlideData9.Append(shapeTree9);
            commonSlideData9.Append(commonSlideDataExtensionList9);

            ColorMapOverride colorMapOverride7 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping7 = new A.MasterColorMapping();

            colorMapOverride7.Append(masterColorMapping7);

            slideLayout5.Append(commonSlideData9);
            slideLayout5.Append(colorMapOverride7);

            slideLayoutPart5.SlideLayout = slideLayout5;
        }

        // Generates content of slideLayoutPart6.
        private void GenerateSlideLayoutPart6Content(SlideLayoutPart slideLayoutPart6)
        {
            SlideLayout slideLayout6 = new SlideLayout() { Type = SlideLayoutValues.Title, Preserve = true };
            slideLayout6.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout6.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout6.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData10 = new CommonSlideData() { Name = "Title Slide" };

            ShapeTree shapeTree10 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties22 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties98 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties22 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties98 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties22.Append(nonVisualDrawingProperties98);
            nonVisualGroupShapeProperties22.Append(nonVisualGroupShapeDrawingProperties22);
            nonVisualGroupShapeProperties22.Append(applicationNonVisualDrawingProperties98);

            GroupShapeProperties groupShapeProperties22 = new GroupShapeProperties();

            A.TransformGroup transformGroup22 = new A.TransformGroup();
            A.Offset offset81 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents81 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset22 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents22 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup22.Append(offset81);
            transformGroup22.Append(extents81);
            transformGroup22.Append(childOffset22);
            transformGroup22.Append(childExtents22);

            groupShapeProperties22.Append(transformGroup22);

            Shape shape51 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties51 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties99 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList78 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension78 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement104 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{F4942847-AFB8-F2BE-18AD-ABE4B141BE47}\" />");

            // nonVisualDrawingPropertiesExtension78.Append(openXmlUnknownElement104);

            nonVisualDrawingPropertiesExtensionList78.Append(nonVisualDrawingPropertiesExtension78);

            nonVisualDrawingProperties99.Append(nonVisualDrawingPropertiesExtensionList78);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties51 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks34 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties51.Append(shapeLocks34);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties99 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape34 = new PlaceholderShape() { Type = PlaceholderValues.CenteredTitle };

            applicationNonVisualDrawingProperties99.Append(placeholderShape34);

            nonVisualShapeProperties51.Append(nonVisualDrawingProperties99);
            nonVisualShapeProperties51.Append(nonVisualShapeDrawingProperties51);
            nonVisualShapeProperties51.Append(applicationNonVisualDrawingProperties99);

            ShapeProperties shapeProperties78 = new ShapeProperties();

            A.Transform2D transform2D59 = new A.Transform2D();
            A.Offset offset82 = new A.Offset() { X = 1524000L, Y = 1122363L };
            A.Extents extents82 = new A.Extents() { Cx = 9144000L, Cy = 2387600L };

            transform2D59.Append(offset82);
            transform2D59.Append(extents82);

            shapeProperties78.Append(transform2D59);

            TextBody textBody71 = new TextBody();
            A.BodyProperties bodyProperties73 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle73 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties18 = new A.Level1ParagraphProperties() { Alignment = A.TextAlignmentTypeValues.Center };
            A.DefaultRunProperties defaultRunProperties91 = new A.DefaultRunProperties() { FontSize = 6000 };

            level1ParagraphProperties18.Append(defaultRunProperties91);

            listStyle73.Append(level1ParagraphProperties18);

            A.Paragraph paragraph87 = new A.Paragraph();

            A.Run run53 = new A.Run();
            A.RunProperties runProperties66 = new A.RunProperties() { Language = "en-US" };
            A.Text text66 = new A.Text();
            text66.Text = "Click to edit Master title style";

            run53.Append(runProperties66);
            run53.Append(text66);

            paragraph87.Append(run53);

            textBody71.Append(bodyProperties73);
            textBody71.Append(listStyle73);
            textBody71.Append(paragraph87);

            shape51.Append(nonVisualShapeProperties51);
            shape51.Append(shapeProperties78);
            shape51.Append(textBody71);

            Shape shape52 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties52 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties100 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Subtitle 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList79 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension79 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement105 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{482F2086-A7E0-EA07-BF14-F3419AB3335E}\" />");

            // nonVisualDrawingPropertiesExtension79.Append(openXmlUnknownElement105);

            nonVisualDrawingPropertiesExtensionList79.Append(nonVisualDrawingPropertiesExtension79);

            nonVisualDrawingProperties100.Append(nonVisualDrawingPropertiesExtensionList79);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties52 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks35 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties52.Append(shapeLocks35);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties100 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape35 = new PlaceholderShape() { Type = PlaceholderValues.SubTitle, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties100.Append(placeholderShape35);

            nonVisualShapeProperties52.Append(nonVisualDrawingProperties100);
            nonVisualShapeProperties52.Append(nonVisualShapeDrawingProperties52);
            nonVisualShapeProperties52.Append(applicationNonVisualDrawingProperties100);

            ShapeProperties shapeProperties79 = new ShapeProperties();

            A.Transform2D transform2D60 = new A.Transform2D();
            A.Offset offset83 = new A.Offset() { X = 1524000L, Y = 3602038L };
            A.Extents extents83 = new A.Extents() { Cx = 9144000L, Cy = 1655762L };

            transform2D60.Append(offset83);
            transform2D60.Append(extents83);

            shapeProperties79.Append(transform2D60);

            TextBody textBody72 = new TextBody();
            A.BodyProperties bodyProperties74 = new A.BodyProperties();

            A.ListStyle listStyle74 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties19 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet34 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties92 = new A.DefaultRunProperties() { FontSize = 2400 };

            level1ParagraphProperties19.Append(noBullet34);
            level1ParagraphProperties19.Append(defaultRunProperties92);

            A.Level2ParagraphProperties level2ParagraphProperties8 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet35 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties93 = new A.DefaultRunProperties() { FontSize = 2000 };

            level2ParagraphProperties8.Append(noBullet35);
            level2ParagraphProperties8.Append(defaultRunProperties93);

            A.Level3ParagraphProperties level3ParagraphProperties8 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet36 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties94 = new A.DefaultRunProperties() { FontSize = 1800 };

            level3ParagraphProperties8.Append(noBullet36);
            level3ParagraphProperties8.Append(defaultRunProperties94);

            A.Level4ParagraphProperties level4ParagraphProperties8 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet37 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties95 = new A.DefaultRunProperties() { FontSize = 1600 };

            level4ParagraphProperties8.Append(noBullet37);
            level4ParagraphProperties8.Append(defaultRunProperties95);

            A.Level5ParagraphProperties level5ParagraphProperties8 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet38 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties96 = new A.DefaultRunProperties() { FontSize = 1600 };

            level5ParagraphProperties8.Append(noBullet38);
            level5ParagraphProperties8.Append(defaultRunProperties96);

            A.Level6ParagraphProperties level6ParagraphProperties8 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet39 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties97 = new A.DefaultRunProperties() { FontSize = 1600 };

            level6ParagraphProperties8.Append(noBullet39);
            level6ParagraphProperties8.Append(defaultRunProperties97);

            A.Level7ParagraphProperties level7ParagraphProperties8 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet40 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties98 = new A.DefaultRunProperties() { FontSize = 1600 };

            level7ParagraphProperties8.Append(noBullet40);
            level7ParagraphProperties8.Append(defaultRunProperties98);

            A.Level8ParagraphProperties level8ParagraphProperties8 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet41 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties99 = new A.DefaultRunProperties() { FontSize = 1600 };

            level8ParagraphProperties8.Append(noBullet41);
            level8ParagraphProperties8.Append(defaultRunProperties99);

            A.Level9ParagraphProperties level9ParagraphProperties8 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Indent = 0, Alignment = A.TextAlignmentTypeValues.Center };
            A.NoBullet noBullet42 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties100 = new A.DefaultRunProperties() { FontSize = 1600 };

            level9ParagraphProperties8.Append(noBullet42);
            level9ParagraphProperties8.Append(defaultRunProperties100);

            listStyle74.Append(level1ParagraphProperties19);
            listStyle74.Append(level2ParagraphProperties8);
            listStyle74.Append(level3ParagraphProperties8);
            listStyle74.Append(level4ParagraphProperties8);
            listStyle74.Append(level5ParagraphProperties8);
            listStyle74.Append(level6ParagraphProperties8);
            listStyle74.Append(level7ParagraphProperties8);
            listStyle74.Append(level8ParagraphProperties8);
            listStyle74.Append(level9ParagraphProperties8);

            A.Paragraph paragraph88 = new A.Paragraph();

            A.Run run54 = new A.Run();
            A.RunProperties runProperties67 = new A.RunProperties() { Language = "en-US" };
            A.Text text67 = new A.Text();
            text67.Text = "Click to edit Master subtitle style";

            run54.Append(runProperties67);
            run54.Append(text67);

            paragraph88.Append(run54);

            textBody72.Append(bodyProperties74);
            textBody72.Append(listStyle74);
            textBody72.Append(paragraph88);

            shape52.Append(nonVisualShapeProperties52);
            shape52.Append(shapeProperties79);
            shape52.Append(textBody72);

            Shape shape53 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties53 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties101 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList80 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension80 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement106 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{75948292-E795-6B47-9E6A-A5C784152537}\" />");

            // nonVisualDrawingPropertiesExtension80.Append(openXmlUnknownElement106);

            nonVisualDrawingPropertiesExtensionList80.Append(nonVisualDrawingPropertiesExtension80);

            nonVisualDrawingProperties101.Append(nonVisualDrawingPropertiesExtensionList80);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties53 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks36 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties53.Append(shapeLocks36);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties101 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape36 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties101.Append(placeholderShape36);

            nonVisualShapeProperties53.Append(nonVisualDrawingProperties101);
            nonVisualShapeProperties53.Append(nonVisualShapeDrawingProperties53);
            nonVisualShapeProperties53.Append(applicationNonVisualDrawingProperties101);
            ShapeProperties shapeProperties80 = new ShapeProperties();

            TextBody textBody73 = new TextBody();
            A.BodyProperties bodyProperties75 = new A.BodyProperties();
            A.ListStyle listStyle75 = new A.ListStyle();

            A.Paragraph paragraph89 = new A.Paragraph();

            A.Field field14 = new A.Field() { Id = "{B5CF329D-DBF6-4E15-9822-D0979B6928D0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties68 = new A.RunProperties() { Language = "en-US" };
            runProperties68.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text68 = new A.Text();
            text68.Text = "4/24/2024";

            field14.Append(runProperties68);
            field14.Append(text68);
            A.EndParagraphRunProperties endParagraphRunProperties56 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph89.Append(field14);
            paragraph89.Append(endParagraphRunProperties56);

            textBody73.Append(bodyProperties75);
            textBody73.Append(listStyle75);
            textBody73.Append(paragraph89);

            shape53.Append(nonVisualShapeProperties53);
            shape53.Append(shapeProperties80);
            shape53.Append(textBody73);

            Shape shape54 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties54 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties102 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList81 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension81 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement107 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{DBFA5441-93D2-AA16-20F4-4ECC92AAA447}\" />");

            // nonVisualDrawingPropertiesExtension81.Append(openXmlUnknownElement107);

            nonVisualDrawingPropertiesExtensionList81.Append(nonVisualDrawingPropertiesExtension81);

            nonVisualDrawingProperties102.Append(nonVisualDrawingPropertiesExtensionList81);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties54 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks37 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties54.Append(shapeLocks37);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties102 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape37 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties102.Append(placeholderShape37);

            nonVisualShapeProperties54.Append(nonVisualDrawingProperties102);
            nonVisualShapeProperties54.Append(nonVisualShapeDrawingProperties54);
            nonVisualShapeProperties54.Append(applicationNonVisualDrawingProperties102);
            ShapeProperties shapeProperties81 = new ShapeProperties();

            TextBody textBody74 = new TextBody();
            A.BodyProperties bodyProperties76 = new A.BodyProperties();
            A.ListStyle listStyle76 = new A.ListStyle();

            A.Paragraph paragraph90 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties57 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph90.Append(endParagraphRunProperties57);

            textBody74.Append(bodyProperties76);
            textBody74.Append(listStyle76);
            textBody74.Append(paragraph90);

            shape54.Append(nonVisualShapeProperties54);
            shape54.Append(shapeProperties81);
            shape54.Append(textBody74);

            Shape shape55 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties55 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties103 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList82 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension82 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement108 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{95D0DBDA-05EC-6960-CFE5-D8A657A445C1}\" />");

            // nonVisualDrawingPropertiesExtension82.Append(openXmlUnknownElement108);

            nonVisualDrawingPropertiesExtensionList82.Append(nonVisualDrawingPropertiesExtension82);

            nonVisualDrawingProperties103.Append(nonVisualDrawingPropertiesExtensionList82);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties55 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks38 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties55.Append(shapeLocks38);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties103 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape38 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties103.Append(placeholderShape38);

            nonVisualShapeProperties55.Append(nonVisualDrawingProperties103);
            nonVisualShapeProperties55.Append(nonVisualShapeDrawingProperties55);
            nonVisualShapeProperties55.Append(applicationNonVisualDrawingProperties103);
            ShapeProperties shapeProperties82 = new ShapeProperties();

            TextBody textBody75 = new TextBody();
            A.BodyProperties bodyProperties77 = new A.BodyProperties();
            A.ListStyle listStyle77 = new A.ListStyle();

            A.Paragraph paragraph91 = new A.Paragraph();

            A.Field field15 = new A.Field() { Id = "{19E6023F-2333-4A80-AFF1-A7F70B52B230}", Type = "slidenum" };

            A.RunProperties runProperties69 = new A.RunProperties() { Language = "en-US" };
            runProperties69.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text69 = new A.Text();
            text69.Text = "‹#›";

            field15.Append(runProperties69);
            field15.Append(text69);
            A.EndParagraphRunProperties endParagraphRunProperties58 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph91.Append(field15);
            paragraph91.Append(endParagraphRunProperties58);

            textBody75.Append(bodyProperties77);
            textBody75.Append(listStyle77);
            textBody75.Append(paragraph91);

            shape55.Append(nonVisualShapeProperties55);
            shape55.Append(shapeProperties82);
            shape55.Append(textBody75);

            shapeTree10.Append(nonVisualGroupShapeProperties22);
            shapeTree10.Append(groupShapeProperties22);
            shapeTree10.Append(shape51);
            shapeTree10.Append(shape52);
            shapeTree10.Append(shape53);
            shapeTree10.Append(shape54);
            shapeTree10.Append(shape55);

            CommonSlideDataExtensionList commonSlideDataExtensionList10 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension10 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId10 = new P14.CreationId() { Val = (UInt32Value)2939110116U };
            creationId10.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension10.Append(creationId10);

            commonSlideDataExtensionList10.Append(commonSlideDataExtension10);

            commonSlideData10.Append(shapeTree10);
            commonSlideData10.Append(commonSlideDataExtensionList10);

            ColorMapOverride colorMapOverride8 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping8 = new A.MasterColorMapping();

            colorMapOverride8.Append(masterColorMapping8);

            slideLayout6.Append(commonSlideData10);
            slideLayout6.Append(colorMapOverride8);

            slideLayoutPart6.SlideLayout = slideLayout6;
        }

        // Generates content of slideLayoutPart7.
        private void GenerateSlideLayoutPart7Content(SlideLayoutPart slideLayoutPart7)
        {
            SlideLayout slideLayout7 = new SlideLayout() { Type = SlideLayoutValues.TitleOnly, Preserve = true };
            slideLayout7.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout7.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout7.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData11 = new CommonSlideData() { Name = "Title Only" };

            ShapeTree shapeTree11 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties23 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties104 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties23 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties104 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties23.Append(nonVisualDrawingProperties104);
            nonVisualGroupShapeProperties23.Append(nonVisualGroupShapeDrawingProperties23);
            nonVisualGroupShapeProperties23.Append(applicationNonVisualDrawingProperties104);

            GroupShapeProperties groupShapeProperties23 = new GroupShapeProperties();

            A.TransformGroup transformGroup23 = new A.TransformGroup();
            A.Offset offset84 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents84 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset23 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents23 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup23.Append(offset84);
            transformGroup23.Append(extents84);
            transformGroup23.Append(childOffset23);
            transformGroup23.Append(childExtents23);

            groupShapeProperties23.Append(transformGroup23);

            Shape shape56 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties56 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties105 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList83 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension83 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement109 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{F138358B-E1C4-DEE9-83B1-4FACE6E0525A}\" />");

            // nonVisualDrawingPropertiesExtension83.Append(openXmlUnknownElement109);

            nonVisualDrawingPropertiesExtensionList83.Append(nonVisualDrawingPropertiesExtension83);

            nonVisualDrawingProperties105.Append(nonVisualDrawingPropertiesExtensionList83);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties56 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks39 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties56.Append(shapeLocks39);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties105 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape39 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties105.Append(placeholderShape39);

            nonVisualShapeProperties56.Append(nonVisualDrawingProperties105);
            nonVisualShapeProperties56.Append(nonVisualShapeDrawingProperties56);
            nonVisualShapeProperties56.Append(applicationNonVisualDrawingProperties105);
            ShapeProperties shapeProperties83 = new ShapeProperties();

            TextBody textBody76 = new TextBody();
            A.BodyProperties bodyProperties78 = new A.BodyProperties();
            A.ListStyle listStyle78 = new A.ListStyle();

            A.Paragraph paragraph92 = new A.Paragraph();

            A.Run run55 = new A.Run();
            A.RunProperties runProperties70 = new A.RunProperties() { Language = "en-US" };
            A.Text text70 = new A.Text();
            text70.Text = "Click to edit Master title style";

            run55.Append(runProperties70);
            run55.Append(text70);

            paragraph92.Append(run55);

            textBody76.Append(bodyProperties78);
            textBody76.Append(listStyle78);
            textBody76.Append(paragraph92);

            shape56.Append(nonVisualShapeProperties56);
            shape56.Append(shapeProperties83);
            shape56.Append(textBody76);

            Shape shape57 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties57 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties106 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Date Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList84 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension84 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement110 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{F1F9B688-765E-7E02-2D58-06F814EDA087}\" />");

            // nonVisualDrawingPropertiesExtension84.Append(openXmlUnknownElement110);

            nonVisualDrawingPropertiesExtensionList84.Append(nonVisualDrawingPropertiesExtension84);

            nonVisualDrawingProperties106.Append(nonVisualDrawingPropertiesExtensionList84);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties57 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks40 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties57.Append(shapeLocks40);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties106 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape40 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties106.Append(placeholderShape40);

            nonVisualShapeProperties57.Append(nonVisualDrawingProperties106);
            nonVisualShapeProperties57.Append(nonVisualShapeDrawingProperties57);
            nonVisualShapeProperties57.Append(applicationNonVisualDrawingProperties106);
            ShapeProperties shapeProperties84 = new ShapeProperties();

            TextBody textBody77 = new TextBody();
            A.BodyProperties bodyProperties79 = new A.BodyProperties();
            A.ListStyle listStyle79 = new A.ListStyle();

            A.Paragraph paragraph93 = new A.Paragraph();

            A.Field field16 = new A.Field() { Id = "{B5CF329D-DBF6-4E15-9822-D0979B6928D0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties71 = new A.RunProperties() { Language = "en-US" };
            runProperties71.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text71 = new A.Text();
            text71.Text = "4/24/2024";

            field16.Append(runProperties71);
            field16.Append(text71);
            A.EndParagraphRunProperties endParagraphRunProperties59 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph93.Append(field16);
            paragraph93.Append(endParagraphRunProperties59);

            textBody77.Append(bodyProperties79);
            textBody77.Append(listStyle79);
            textBody77.Append(paragraph93);

            shape57.Append(nonVisualShapeProperties57);
            shape57.Append(shapeProperties84);
            shape57.Append(textBody77);

            Shape shape58 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties58 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties107 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Footer Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList85 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension85 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement111 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{86BB40C4-BBFC-D645-048E-ECF2EC97756E}\" />");

            // nonVisualDrawingPropertiesExtension85.Append(openXmlUnknownElement111);

            nonVisualDrawingPropertiesExtensionList85.Append(nonVisualDrawingPropertiesExtension85);

            nonVisualDrawingProperties107.Append(nonVisualDrawingPropertiesExtensionList85);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties58 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks41 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties58.Append(shapeLocks41);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties107 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape41 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties107.Append(placeholderShape41);

            nonVisualShapeProperties58.Append(nonVisualDrawingProperties107);
            nonVisualShapeProperties58.Append(nonVisualShapeDrawingProperties58);
            nonVisualShapeProperties58.Append(applicationNonVisualDrawingProperties107);
            ShapeProperties shapeProperties85 = new ShapeProperties();

            TextBody textBody78 = new TextBody();
            A.BodyProperties bodyProperties80 = new A.BodyProperties();
            A.ListStyle listStyle80 = new A.ListStyle();

            A.Paragraph paragraph94 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties60 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph94.Append(endParagraphRunProperties60);

            textBody78.Append(bodyProperties80);
            textBody78.Append(listStyle80);
            textBody78.Append(paragraph94);

            shape58.Append(nonVisualShapeProperties58);
            shape58.Append(shapeProperties85);
            shape58.Append(textBody78);

            Shape shape59 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties59 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties108 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Slide Number Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList86 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension86 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement112 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{563EFA97-9213-F392-5272-4AE58D867D83}\" />");

            // nonVisualDrawingPropertiesExtension86.Append(openXmlUnknownElement112);

            nonVisualDrawingPropertiesExtensionList86.Append(nonVisualDrawingPropertiesExtension86);

            nonVisualDrawingProperties108.Append(nonVisualDrawingPropertiesExtensionList86);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties59 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks42 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties59.Append(shapeLocks42);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties108 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape42 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties108.Append(placeholderShape42);

            nonVisualShapeProperties59.Append(nonVisualDrawingProperties108);
            nonVisualShapeProperties59.Append(nonVisualShapeDrawingProperties59);
            nonVisualShapeProperties59.Append(applicationNonVisualDrawingProperties108);
            ShapeProperties shapeProperties86 = new ShapeProperties();

            TextBody textBody79 = new TextBody();
            A.BodyProperties bodyProperties81 = new A.BodyProperties();
            A.ListStyle listStyle81 = new A.ListStyle();

            A.Paragraph paragraph95 = new A.Paragraph();

            A.Field field17 = new A.Field() { Id = "{19E6023F-2333-4A80-AFF1-A7F70B52B230}", Type = "slidenum" };

            A.RunProperties runProperties72 = new A.RunProperties() { Language = "en-US" };
            runProperties72.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text72 = new A.Text();
            text72.Text = "‹#›";

            field17.Append(runProperties72);
            field17.Append(text72);
            A.EndParagraphRunProperties endParagraphRunProperties61 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph95.Append(field17);
            paragraph95.Append(endParagraphRunProperties61);

            textBody79.Append(bodyProperties81);
            textBody79.Append(listStyle81);
            textBody79.Append(paragraph95);

            shape59.Append(nonVisualShapeProperties59);
            shape59.Append(shapeProperties86);
            shape59.Append(textBody79);

            shapeTree11.Append(nonVisualGroupShapeProperties23);
            shapeTree11.Append(groupShapeProperties23);
            shapeTree11.Append(shape56);
            shapeTree11.Append(shape57);
            shapeTree11.Append(shape58);
            shapeTree11.Append(shape59);

            CommonSlideDataExtensionList commonSlideDataExtensionList11 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension11 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId11 = new P14.CreationId() { Val = (UInt32Value)1345740126U };
            creationId11.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension11.Append(creationId11);

            commonSlideDataExtensionList11.Append(commonSlideDataExtension11);

            commonSlideData11.Append(shapeTree11);
            commonSlideData11.Append(commonSlideDataExtensionList11);

            ColorMapOverride colorMapOverride9 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping9 = new A.MasterColorMapping();

            colorMapOverride9.Append(masterColorMapping9);

            slideLayout7.Append(commonSlideData11);
            slideLayout7.Append(colorMapOverride9);

            slideLayoutPart7.SlideLayout = slideLayout7;
        }

        // Generates content of slideLayoutPart8.
        private void GenerateSlideLayoutPart8Content(SlideLayoutPart slideLayoutPart8)
        {
            SlideLayout slideLayout8 = new SlideLayout() { Type = SlideLayoutValues.VerticalTitleAndText, Preserve = true };
            slideLayout8.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout8.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout8.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData12 = new CommonSlideData() { Name = "Vertical Title and Text" };

            ShapeTree shapeTree12 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties24 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties109 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties24 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties109 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties24.Append(nonVisualDrawingProperties109);
            nonVisualGroupShapeProperties24.Append(nonVisualGroupShapeDrawingProperties24);
            nonVisualGroupShapeProperties24.Append(applicationNonVisualDrawingProperties109);

            GroupShapeProperties groupShapeProperties24 = new GroupShapeProperties();

            A.TransformGroup transformGroup24 = new A.TransformGroup();
            A.Offset offset85 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents85 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset24 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents24 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup24.Append(offset85);
            transformGroup24.Append(extents85);
            transformGroup24.Append(childOffset24);
            transformGroup24.Append(childExtents24);

            groupShapeProperties24.Append(transformGroup24);

            Shape shape60 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties60 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties110 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Vertical Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList87 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension87 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement113 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{772BEEC1-2B59-5BC3-3E10-AEFD6782EBB6}\" />");

            // nonVisualDrawingPropertiesExtension87.Append(openXmlUnknownElement113);

            nonVisualDrawingPropertiesExtensionList87.Append(nonVisualDrawingPropertiesExtension87);

            nonVisualDrawingProperties110.Append(nonVisualDrawingPropertiesExtensionList87);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties60 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks43 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties60.Append(shapeLocks43);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties110 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape43 = new PlaceholderShape() { Type = PlaceholderValues.Title, Orientation = DirectionValues.Vertical };

            applicationNonVisualDrawingProperties110.Append(placeholderShape43);

            nonVisualShapeProperties60.Append(nonVisualDrawingProperties110);
            nonVisualShapeProperties60.Append(nonVisualShapeDrawingProperties60);
            nonVisualShapeProperties60.Append(applicationNonVisualDrawingProperties110);

            ShapeProperties shapeProperties87 = new ShapeProperties();

            A.Transform2D transform2D61 = new A.Transform2D();
            A.Offset offset86 = new A.Offset() { X = 8724900L, Y = 365125L };
            A.Extents extents86 = new A.Extents() { Cx = 2628900L, Cy = 5811838L };

            transform2D61.Append(offset86);
            transform2D61.Append(extents86);

            shapeProperties87.Append(transform2D61);

            TextBody textBody80 = new TextBody();
            A.BodyProperties bodyProperties82 = new A.BodyProperties() { Vertical = A.TextVerticalValues.EastAsianVetical };
            A.ListStyle listStyle82 = new A.ListStyle();

            A.Paragraph paragraph96 = new A.Paragraph();

            A.Run run56 = new A.Run();
            A.RunProperties runProperties73 = new A.RunProperties() { Language = "en-US" };
            A.Text text73 = new A.Text();
            text73.Text = "Click to edit Master title style";

            run56.Append(runProperties73);
            run56.Append(text73);

            paragraph96.Append(run56);

            textBody80.Append(bodyProperties82);
            textBody80.Append(listStyle82);
            textBody80.Append(paragraph96);

            shape60.Append(nonVisualShapeProperties60);
            shape60.Append(shapeProperties87);
            shape60.Append(textBody80);

            Shape shape61 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties61 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties111 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Vertical Text Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList88 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension88 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement114 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{67E39F03-6DCC-C9D3-0868-919ACA891DC5}\" />");

            // nonVisualDrawingPropertiesExtension88.Append(openXmlUnknownElement114);

            nonVisualDrawingPropertiesExtensionList88.Append(nonVisualDrawingPropertiesExtension88);

            nonVisualDrawingProperties111.Append(nonVisualDrawingPropertiesExtensionList88);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties61 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks44 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties61.Append(shapeLocks44);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties111 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape44 = new PlaceholderShape() { Type = PlaceholderValues.Body, Orientation = DirectionValues.Vertical, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties111.Append(placeholderShape44);

            nonVisualShapeProperties61.Append(nonVisualDrawingProperties111);
            nonVisualShapeProperties61.Append(nonVisualShapeDrawingProperties61);
            nonVisualShapeProperties61.Append(applicationNonVisualDrawingProperties111);

            ShapeProperties shapeProperties88 = new ShapeProperties();

            A.Transform2D transform2D62 = new A.Transform2D();
            A.Offset offset87 = new A.Offset() { X = 838200L, Y = 365125L };
            A.Extents extents87 = new A.Extents() { Cx = 7734300L, Cy = 5811838L };

            transform2D62.Append(offset87);
            transform2D62.Append(extents87);

            shapeProperties88.Append(transform2D62);

            TextBody textBody81 = new TextBody();
            A.BodyProperties bodyProperties83 = new A.BodyProperties() { Vertical = A.TextVerticalValues.EastAsianVetical };
            A.ListStyle listStyle83 = new A.ListStyle();

            A.Paragraph paragraph97 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties57 = new A.ParagraphProperties() { Level = 0 };

            A.Run run57 = new A.Run();
            A.RunProperties runProperties74 = new A.RunProperties() { Language = "en-US" };
            A.Text text74 = new A.Text();
            text74.Text = "Click to edit Master text styles";

            run57.Append(runProperties74);
            run57.Append(text74);

            paragraph97.Append(paragraphProperties57);
            paragraph97.Append(run57);

            A.Paragraph paragraph98 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties58 = new A.ParagraphProperties() { Level = 1 };

            A.Run run58 = new A.Run();
            A.RunProperties runProperties75 = new A.RunProperties() { Language = "en-US" };
            A.Text text75 = new A.Text();
            text75.Text = "Second level";

            run58.Append(runProperties75);
            run58.Append(text75);

            paragraph98.Append(paragraphProperties58);
            paragraph98.Append(run58);

            A.Paragraph paragraph99 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties59 = new A.ParagraphProperties() { Level = 2 };

            A.Run run59 = new A.Run();
            A.RunProperties runProperties76 = new A.RunProperties() { Language = "en-US" };
            A.Text text76 = new A.Text();
            text76.Text = "Third level";

            run59.Append(runProperties76);
            run59.Append(text76);

            paragraph99.Append(paragraphProperties59);
            paragraph99.Append(run59);

            A.Paragraph paragraph100 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties60 = new A.ParagraphProperties() { Level = 3 };

            A.Run run60 = new A.Run();
            A.RunProperties runProperties77 = new A.RunProperties() { Language = "en-US" };
            A.Text text77 = new A.Text();
            text77.Text = "Fourth level";

            run60.Append(runProperties77);
            run60.Append(text77);

            paragraph100.Append(paragraphProperties60);
            paragraph100.Append(run60);

            A.Paragraph paragraph101 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties61 = new A.ParagraphProperties() { Level = 4 };

            A.Run run61 = new A.Run();
            A.RunProperties runProperties78 = new A.RunProperties() { Language = "en-US" };
            A.Text text78 = new A.Text();
            text78.Text = "Fifth level";

            run61.Append(runProperties78);
            run61.Append(text78);

            paragraph101.Append(paragraphProperties61);
            paragraph101.Append(run61);

            textBody81.Append(bodyProperties83);
            textBody81.Append(listStyle83);
            textBody81.Append(paragraph97);
            textBody81.Append(paragraph98);
            textBody81.Append(paragraph99);
            textBody81.Append(paragraph100);
            textBody81.Append(paragraph101);

            shape61.Append(nonVisualShapeProperties61);
            shape61.Append(shapeProperties88);
            shape61.Append(textBody81);

            Shape shape62 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties62 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties112 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList89 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension89 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement115 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{76D9B9B8-796A-70AB-F5BE-7E4BB73B9932}\" />");

            // nonVisualDrawingPropertiesExtension89.Append(openXmlUnknownElement115);

            nonVisualDrawingPropertiesExtensionList89.Append(nonVisualDrawingPropertiesExtension89);

            nonVisualDrawingProperties112.Append(nonVisualDrawingPropertiesExtensionList89);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties62 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks45 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties62.Append(shapeLocks45);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties112 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape45 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties112.Append(placeholderShape45);

            nonVisualShapeProperties62.Append(nonVisualDrawingProperties112);
            nonVisualShapeProperties62.Append(nonVisualShapeDrawingProperties62);
            nonVisualShapeProperties62.Append(applicationNonVisualDrawingProperties112);
            ShapeProperties shapeProperties89 = new ShapeProperties();

            TextBody textBody82 = new TextBody();
            A.BodyProperties bodyProperties84 = new A.BodyProperties();
            A.ListStyle listStyle84 = new A.ListStyle();

            A.Paragraph paragraph102 = new A.Paragraph();

            A.Field field18 = new A.Field() { Id = "{B5CF329D-DBF6-4E15-9822-D0979B6928D0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties79 = new A.RunProperties() { Language = "en-US" };
            runProperties79.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text79 = new A.Text();
            text79.Text = "4/24/2024";

            field18.Append(runProperties79);
            field18.Append(text79);
            A.EndParagraphRunProperties endParagraphRunProperties62 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph102.Append(field18);
            paragraph102.Append(endParagraphRunProperties62);

            textBody82.Append(bodyProperties84);
            textBody82.Append(listStyle84);
            textBody82.Append(paragraph102);

            shape62.Append(nonVisualShapeProperties62);
            shape62.Append(shapeProperties89);
            shape62.Append(textBody82);

            Shape shape63 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties63 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties113 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList90 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension90 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement116 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{194D330F-AABE-0DCE-B065-68427DA8C849}\" />");

            // nonVisualDrawingPropertiesExtension90.Append(openXmlUnknownElement116);

            nonVisualDrawingPropertiesExtensionList90.Append(nonVisualDrawingPropertiesExtension90);

            nonVisualDrawingProperties113.Append(nonVisualDrawingPropertiesExtensionList90);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties63 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks46 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties63.Append(shapeLocks46);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties113 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape46 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties113.Append(placeholderShape46);

            nonVisualShapeProperties63.Append(nonVisualDrawingProperties113);
            nonVisualShapeProperties63.Append(nonVisualShapeDrawingProperties63);
            nonVisualShapeProperties63.Append(applicationNonVisualDrawingProperties113);
            ShapeProperties shapeProperties90 = new ShapeProperties();

            TextBody textBody83 = new TextBody();
            A.BodyProperties bodyProperties85 = new A.BodyProperties();
            A.ListStyle listStyle85 = new A.ListStyle();

            A.Paragraph paragraph103 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties63 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph103.Append(endParagraphRunProperties63);

            textBody83.Append(bodyProperties85);
            textBody83.Append(listStyle85);
            textBody83.Append(paragraph103);

            shape63.Append(nonVisualShapeProperties63);
            shape63.Append(shapeProperties90);
            shape63.Append(textBody83);

            Shape shape64 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties64 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties114 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList91 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension91 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement117 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{503B0D1F-FF55-46B1-326D-6942DB351A28}\" />");

            // nonVisualDrawingPropertiesExtension91.Append(openXmlUnknownElement117);

            nonVisualDrawingPropertiesExtensionList91.Append(nonVisualDrawingPropertiesExtension91);

            nonVisualDrawingProperties114.Append(nonVisualDrawingPropertiesExtensionList91);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties64 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks47 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties64.Append(shapeLocks47);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties114 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape47 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties114.Append(placeholderShape47);

            nonVisualShapeProperties64.Append(nonVisualDrawingProperties114);
            nonVisualShapeProperties64.Append(nonVisualShapeDrawingProperties64);
            nonVisualShapeProperties64.Append(applicationNonVisualDrawingProperties114);
            ShapeProperties shapeProperties91 = new ShapeProperties();

            TextBody textBody84 = new TextBody();
            A.BodyProperties bodyProperties86 = new A.BodyProperties();
            A.ListStyle listStyle86 = new A.ListStyle();

            A.Paragraph paragraph104 = new A.Paragraph();

            A.Field field19 = new A.Field() { Id = "{19E6023F-2333-4A80-AFF1-A7F70B52B230}", Type = "slidenum" };

            A.RunProperties runProperties80 = new A.RunProperties() { Language = "en-US" };
            runProperties80.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text80 = new A.Text();
            text80.Text = "‹#›";

            field19.Append(runProperties80);
            field19.Append(text80);
            A.EndParagraphRunProperties endParagraphRunProperties64 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph104.Append(field19);
            paragraph104.Append(endParagraphRunProperties64);

            textBody84.Append(bodyProperties86);
            textBody84.Append(listStyle86);
            textBody84.Append(paragraph104);

            shape64.Append(nonVisualShapeProperties64);
            shape64.Append(shapeProperties91);
            shape64.Append(textBody84);

            shapeTree12.Append(nonVisualGroupShapeProperties24);
            shapeTree12.Append(groupShapeProperties24);
            shapeTree12.Append(shape60);
            shapeTree12.Append(shape61);
            shapeTree12.Append(shape62);
            shapeTree12.Append(shape63);
            shapeTree12.Append(shape64);

            CommonSlideDataExtensionList commonSlideDataExtensionList12 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension12 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId12 = new P14.CreationId() { Val = (UInt32Value)4147973603U };
            creationId12.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension12.Append(creationId12);

            commonSlideDataExtensionList12.Append(commonSlideDataExtension12);

            commonSlideData12.Append(shapeTree12);
            commonSlideData12.Append(commonSlideDataExtensionList12);

            ColorMapOverride colorMapOverride10 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping10 = new A.MasterColorMapping();

            colorMapOverride10.Append(masterColorMapping10);

            slideLayout8.Append(commonSlideData12);
            slideLayout8.Append(colorMapOverride10);

            slideLayoutPart8.SlideLayout = slideLayout8;
        }

        // Generates content of slideLayoutPart9.
        private void GenerateSlideLayoutPart9Content(SlideLayoutPart slideLayoutPart9)
        {
            SlideLayout slideLayout9 = new SlideLayout() { Type = SlideLayoutValues.TwoTextAndTwoObjects, Preserve = true };
            slideLayout9.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout9.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout9.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData13 = new CommonSlideData() { Name = "Comparison" };

            ShapeTree shapeTree13 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties25 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties115 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties25 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties115 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties25.Append(nonVisualDrawingProperties115);
            nonVisualGroupShapeProperties25.Append(nonVisualGroupShapeDrawingProperties25);
            nonVisualGroupShapeProperties25.Append(applicationNonVisualDrawingProperties115);

            GroupShapeProperties groupShapeProperties25 = new GroupShapeProperties();

            A.TransformGroup transformGroup25 = new A.TransformGroup();
            A.Offset offset88 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents88 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset25 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents25 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup25.Append(offset88);
            transformGroup25.Append(extents88);
            transformGroup25.Append(childOffset25);
            transformGroup25.Append(childExtents25);

            groupShapeProperties25.Append(transformGroup25);

            Shape shape65 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties65 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties116 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList92 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension92 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement118 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{830EA02E-77C2-94A4-1BF9-5B2129FCEDF2}\" />");

            // nonVisualDrawingPropertiesExtension92.Append(openXmlUnknownElement118);

            nonVisualDrawingPropertiesExtensionList92.Append(nonVisualDrawingPropertiesExtension92);

            nonVisualDrawingProperties116.Append(nonVisualDrawingPropertiesExtensionList92);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties65 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks48 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties65.Append(shapeLocks48);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties116 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape48 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties116.Append(placeholderShape48);

            nonVisualShapeProperties65.Append(nonVisualDrawingProperties116);
            nonVisualShapeProperties65.Append(nonVisualShapeDrawingProperties65);
            nonVisualShapeProperties65.Append(applicationNonVisualDrawingProperties116);

            ShapeProperties shapeProperties92 = new ShapeProperties();

            A.Transform2D transform2D63 = new A.Transform2D();
            A.Offset offset89 = new A.Offset() { X = 839788L, Y = 365125L };
            A.Extents extents89 = new A.Extents() { Cx = 10515600L, Cy = 1325563L };

            transform2D63.Append(offset89);
            transform2D63.Append(extents89);

            shapeProperties92.Append(transform2D63);

            TextBody textBody85 = new TextBody();
            A.BodyProperties bodyProperties87 = new A.BodyProperties();
            A.ListStyle listStyle87 = new A.ListStyle();

            A.Paragraph paragraph105 = new A.Paragraph();

            A.Run run62 = new A.Run();
            A.RunProperties runProperties81 = new A.RunProperties() { Language = "en-US" };
            A.Text text81 = new A.Text();
            text81.Text = "Click to edit Master title style";

            run62.Append(runProperties81);
            run62.Append(text81);

            paragraph105.Append(run62);

            textBody85.Append(bodyProperties87);
            textBody85.Append(listStyle87);
            textBody85.Append(paragraph105);

            shape65.Append(nonVisualShapeProperties65);
            shape65.Append(shapeProperties92);
            shape65.Append(textBody85);

            Shape shape66 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties66 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties117 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Text Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList93 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension93 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement119 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{7161378E-D824-47FC-7293-E4A43BA2B163}\" />");

            // nonVisualDrawingPropertiesExtension93.Append(openXmlUnknownElement119);

            nonVisualDrawingPropertiesExtensionList93.Append(nonVisualDrawingPropertiesExtension93);

            nonVisualDrawingProperties117.Append(nonVisualDrawingPropertiesExtensionList93);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties66 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks49 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties66.Append(shapeLocks49);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties117 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape49 = new PlaceholderShape() { Type = PlaceholderValues.Body, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties117.Append(placeholderShape49);

            nonVisualShapeProperties66.Append(nonVisualDrawingProperties117);
            nonVisualShapeProperties66.Append(nonVisualShapeDrawingProperties66);
            nonVisualShapeProperties66.Append(applicationNonVisualDrawingProperties117);

            ShapeProperties shapeProperties93 = new ShapeProperties();

            A.Transform2D transform2D64 = new A.Transform2D();
            A.Offset offset90 = new A.Offset() { X = 839788L, Y = 1681163L };
            A.Extents extents90 = new A.Extents() { Cx = 5157787L, Cy = 823912L };

            transform2D64.Append(offset90);
            transform2D64.Append(extents90);

            shapeProperties93.Append(transform2D64);

            TextBody textBody86 = new TextBody();
            A.BodyProperties bodyProperties88 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle88 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties20 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet43 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties101 = new A.DefaultRunProperties() { FontSize = 2400, Bold = true };

            level1ParagraphProperties20.Append(noBullet43);
            level1ParagraphProperties20.Append(defaultRunProperties101);

            A.Level2ParagraphProperties level2ParagraphProperties9 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet44 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties102 = new A.DefaultRunProperties() { FontSize = 2000, Bold = true };

            level2ParagraphProperties9.Append(noBullet44);
            level2ParagraphProperties9.Append(defaultRunProperties102);

            A.Level3ParagraphProperties level3ParagraphProperties9 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet45 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties103 = new A.DefaultRunProperties() { FontSize = 1800, Bold = true };

            level3ParagraphProperties9.Append(noBullet45);
            level3ParagraphProperties9.Append(defaultRunProperties103);

            A.Level4ParagraphProperties level4ParagraphProperties9 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet46 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties104 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level4ParagraphProperties9.Append(noBullet46);
            level4ParagraphProperties9.Append(defaultRunProperties104);

            A.Level5ParagraphProperties level5ParagraphProperties9 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet47 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties105 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level5ParagraphProperties9.Append(noBullet47);
            level5ParagraphProperties9.Append(defaultRunProperties105);

            A.Level6ParagraphProperties level6ParagraphProperties9 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet48 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties106 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level6ParagraphProperties9.Append(noBullet48);
            level6ParagraphProperties9.Append(defaultRunProperties106);

            A.Level7ParagraphProperties level7ParagraphProperties9 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet49 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties107 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level7ParagraphProperties9.Append(noBullet49);
            level7ParagraphProperties9.Append(defaultRunProperties107);

            A.Level8ParagraphProperties level8ParagraphProperties9 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet50 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties108 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level8ParagraphProperties9.Append(noBullet50);
            level8ParagraphProperties9.Append(defaultRunProperties108);

            A.Level9ParagraphProperties level9ParagraphProperties9 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet51 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties109 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level9ParagraphProperties9.Append(noBullet51);
            level9ParagraphProperties9.Append(defaultRunProperties109);

            listStyle88.Append(level1ParagraphProperties20);
            listStyle88.Append(level2ParagraphProperties9);
            listStyle88.Append(level3ParagraphProperties9);
            listStyle88.Append(level4ParagraphProperties9);
            listStyle88.Append(level5ParagraphProperties9);
            listStyle88.Append(level6ParagraphProperties9);
            listStyle88.Append(level7ParagraphProperties9);
            listStyle88.Append(level8ParagraphProperties9);
            listStyle88.Append(level9ParagraphProperties9);

            A.Paragraph paragraph106 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties62 = new A.ParagraphProperties() { Level = 0 };

            A.Run run63 = new A.Run();
            A.RunProperties runProperties82 = new A.RunProperties() { Language = "en-US" };
            A.Text text82 = new A.Text();
            text82.Text = "Click to edit Master text styles";

            run63.Append(runProperties82);
            run63.Append(text82);

            paragraph106.Append(paragraphProperties62);
            paragraph106.Append(run63);

            textBody86.Append(bodyProperties88);
            textBody86.Append(listStyle88);
            textBody86.Append(paragraph106);

            shape66.Append(nonVisualShapeProperties66);
            shape66.Append(shapeProperties93);
            shape66.Append(textBody86);

            Shape shape67 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties67 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties118 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Content Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList94 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension94 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement120 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{9769F676-09C8-0BE7-8FB9-F8390DB7AD0F}\" />");

            // nonVisualDrawingPropertiesExtension94.Append(openXmlUnknownElement120);

            nonVisualDrawingPropertiesExtensionList94.Append(nonVisualDrawingPropertiesExtension94);

            nonVisualDrawingProperties118.Append(nonVisualDrawingPropertiesExtensionList94);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties67 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks50 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties67.Append(shapeLocks50);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties118 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape50 = new PlaceholderShape() { Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties118.Append(placeholderShape50);

            nonVisualShapeProperties67.Append(nonVisualDrawingProperties118);
            nonVisualShapeProperties67.Append(nonVisualShapeDrawingProperties67);
            nonVisualShapeProperties67.Append(applicationNonVisualDrawingProperties118);

            ShapeProperties shapeProperties94 = new ShapeProperties();

            A.Transform2D transform2D65 = new A.Transform2D();
            A.Offset offset91 = new A.Offset() { X = 839788L, Y = 2505075L };
            A.Extents extents91 = new A.Extents() { Cx = 5157787L, Cy = 3684588L };

            transform2D65.Append(offset91);
            transform2D65.Append(extents91);

            shapeProperties94.Append(transform2D65);

            TextBody textBody87 = new TextBody();
            A.BodyProperties bodyProperties89 = new A.BodyProperties();
            A.ListStyle listStyle89 = new A.ListStyle();

            A.Paragraph paragraph107 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties63 = new A.ParagraphProperties() { Level = 0 };

            A.Run run64 = new A.Run();
            A.RunProperties runProperties83 = new A.RunProperties() { Language = "en-US" };
            A.Text text83 = new A.Text();
            text83.Text = "Click to edit Master text styles";

            run64.Append(runProperties83);
            run64.Append(text83);

            paragraph107.Append(paragraphProperties63);
            paragraph107.Append(run64);

            A.Paragraph paragraph108 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties64 = new A.ParagraphProperties() { Level = 1 };

            A.Run run65 = new A.Run();
            A.RunProperties runProperties84 = new A.RunProperties() { Language = "en-US" };
            A.Text text84 = new A.Text();
            text84.Text = "Second level";

            run65.Append(runProperties84);
            run65.Append(text84);

            paragraph108.Append(paragraphProperties64);
            paragraph108.Append(run65);

            A.Paragraph paragraph109 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties65 = new A.ParagraphProperties() { Level = 2 };

            A.Run run66 = new A.Run();
            A.RunProperties runProperties85 = new A.RunProperties() { Language = "en-US" };
            A.Text text85 = new A.Text();
            text85.Text = "Third level";

            run66.Append(runProperties85);
            run66.Append(text85);

            paragraph109.Append(paragraphProperties65);
            paragraph109.Append(run66);

            A.Paragraph paragraph110 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties66 = new A.ParagraphProperties() { Level = 3 };

            A.Run run67 = new A.Run();
            A.RunProperties runProperties86 = new A.RunProperties() { Language = "en-US" };
            A.Text text86 = new A.Text();
            text86.Text = "Fourth level";

            run67.Append(runProperties86);
            run67.Append(text86);

            paragraph110.Append(paragraphProperties66);
            paragraph110.Append(run67);

            A.Paragraph paragraph111 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties67 = new A.ParagraphProperties() { Level = 4 };

            A.Run run68 = new A.Run();
            A.RunProperties runProperties87 = new A.RunProperties() { Language = "en-US" };
            A.Text text87 = new A.Text();
            text87.Text = "Fifth level";

            run68.Append(runProperties87);
            run68.Append(text87);

            paragraph111.Append(paragraphProperties67);
            paragraph111.Append(run68);

            textBody87.Append(bodyProperties89);
            textBody87.Append(listStyle89);
            textBody87.Append(paragraph107);
            textBody87.Append(paragraph108);
            textBody87.Append(paragraph109);
            textBody87.Append(paragraph110);
            textBody87.Append(paragraph111);

            shape67.Append(nonVisualShapeProperties67);
            shape67.Append(shapeProperties94);
            shape67.Append(textBody87);

            Shape shape68 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties68 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties119 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Text Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList95 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension95 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement121 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{C2589E4A-A93B-12E9-0A02-5DBE64D7D8AF}\" />");

            // nonVisualDrawingPropertiesExtension95.Append(openXmlUnknownElement121);

            nonVisualDrawingPropertiesExtensionList95.Append(nonVisualDrawingPropertiesExtension95);

            nonVisualDrawingProperties119.Append(nonVisualDrawingPropertiesExtensionList95);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties68 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks51 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties68.Append(shapeLocks51);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties119 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape51 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)3U };

            applicationNonVisualDrawingProperties119.Append(placeholderShape51);

            nonVisualShapeProperties68.Append(nonVisualDrawingProperties119);
            nonVisualShapeProperties68.Append(nonVisualShapeDrawingProperties68);
            nonVisualShapeProperties68.Append(applicationNonVisualDrawingProperties119);

            ShapeProperties shapeProperties95 = new ShapeProperties();

            A.Transform2D transform2D66 = new A.Transform2D();
            A.Offset offset92 = new A.Offset() { X = 6172200L, Y = 1681163L };
            A.Extents extents92 = new A.Extents() { Cx = 5183188L, Cy = 823912L };

            transform2D66.Append(offset92);
            transform2D66.Append(extents92);

            shapeProperties95.Append(transform2D66);

            TextBody textBody88 = new TextBody();
            A.BodyProperties bodyProperties90 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle90 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties21 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet52 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties110 = new A.DefaultRunProperties() { FontSize = 2400, Bold = true };

            level1ParagraphProperties21.Append(noBullet52);
            level1ParagraphProperties21.Append(defaultRunProperties110);

            A.Level2ParagraphProperties level2ParagraphProperties10 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet53 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties111 = new A.DefaultRunProperties() { FontSize = 2000, Bold = true };

            level2ParagraphProperties10.Append(noBullet53);
            level2ParagraphProperties10.Append(defaultRunProperties111);

            A.Level3ParagraphProperties level3ParagraphProperties10 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet54 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties112 = new A.DefaultRunProperties() { FontSize = 1800, Bold = true };

            level3ParagraphProperties10.Append(noBullet54);
            level3ParagraphProperties10.Append(defaultRunProperties112);

            A.Level4ParagraphProperties level4ParagraphProperties10 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet55 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties113 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level4ParagraphProperties10.Append(noBullet55);
            level4ParagraphProperties10.Append(defaultRunProperties113);

            A.Level5ParagraphProperties level5ParagraphProperties10 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet56 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties114 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level5ParagraphProperties10.Append(noBullet56);
            level5ParagraphProperties10.Append(defaultRunProperties114);

            A.Level6ParagraphProperties level6ParagraphProperties10 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet57 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties115 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level6ParagraphProperties10.Append(noBullet57);
            level6ParagraphProperties10.Append(defaultRunProperties115);

            A.Level7ParagraphProperties level7ParagraphProperties10 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet58 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties116 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level7ParagraphProperties10.Append(noBullet58);
            level7ParagraphProperties10.Append(defaultRunProperties116);

            A.Level8ParagraphProperties level8ParagraphProperties10 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet59 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties117 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level8ParagraphProperties10.Append(noBullet59);
            level8ParagraphProperties10.Append(defaultRunProperties117);

            A.Level9ParagraphProperties level9ParagraphProperties10 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet60 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties118 = new A.DefaultRunProperties() { FontSize = 1600, Bold = true };

            level9ParagraphProperties10.Append(noBullet60);
            level9ParagraphProperties10.Append(defaultRunProperties118);

            listStyle90.Append(level1ParagraphProperties21);
            listStyle90.Append(level2ParagraphProperties10);
            listStyle90.Append(level3ParagraphProperties10);
            listStyle90.Append(level4ParagraphProperties10);
            listStyle90.Append(level5ParagraphProperties10);
            listStyle90.Append(level6ParagraphProperties10);
            listStyle90.Append(level7ParagraphProperties10);
            listStyle90.Append(level8ParagraphProperties10);
            listStyle90.Append(level9ParagraphProperties10);

            A.Paragraph paragraph112 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties68 = new A.ParagraphProperties() { Level = 0 };

            A.Run run69 = new A.Run();
            A.RunProperties runProperties88 = new A.RunProperties() { Language = "en-US" };
            A.Text text88 = new A.Text();
            text88.Text = "Click to edit Master text styles";

            run69.Append(runProperties88);
            run69.Append(text88);

            paragraph112.Append(paragraphProperties68);
            paragraph112.Append(run69);

            textBody88.Append(bodyProperties90);
            textBody88.Append(listStyle90);
            textBody88.Append(paragraph112);

            shape68.Append(nonVisualShapeProperties68);
            shape68.Append(shapeProperties95);
            shape68.Append(textBody88);

            Shape shape69 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties69 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties120 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Content Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList96 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension96 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement122 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{7E9A1A8E-CA90-002F-B078-C63FB5820989}\" />");

            // nonVisualDrawingPropertiesExtension96.Append(openXmlUnknownElement122);

            nonVisualDrawingPropertiesExtensionList96.Append(nonVisualDrawingPropertiesExtension96);

            nonVisualDrawingProperties120.Append(nonVisualDrawingPropertiesExtensionList96);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties69 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks52 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties69.Append(shapeLocks52);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties120 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape52 = new PlaceholderShape() { Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)4U };

            applicationNonVisualDrawingProperties120.Append(placeholderShape52);

            nonVisualShapeProperties69.Append(nonVisualDrawingProperties120);
            nonVisualShapeProperties69.Append(nonVisualShapeDrawingProperties69);
            nonVisualShapeProperties69.Append(applicationNonVisualDrawingProperties120);

            ShapeProperties shapeProperties96 = new ShapeProperties();

            A.Transform2D transform2D67 = new A.Transform2D();
            A.Offset offset93 = new A.Offset() { X = 6172200L, Y = 2505075L };
            A.Extents extents93 = new A.Extents() { Cx = 5183188L, Cy = 3684588L };

            transform2D67.Append(offset93);
            transform2D67.Append(extents93);

            shapeProperties96.Append(transform2D67);

            TextBody textBody89 = new TextBody();
            A.BodyProperties bodyProperties91 = new A.BodyProperties();
            A.ListStyle listStyle91 = new A.ListStyle();

            A.Paragraph paragraph113 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties69 = new A.ParagraphProperties() { Level = 0 };

            A.Run run70 = new A.Run();
            A.RunProperties runProperties89 = new A.RunProperties() { Language = "en-US" };
            A.Text text89 = new A.Text();
            text89.Text = "Click to edit Master text styles";

            run70.Append(runProperties89);
            run70.Append(text89);

            paragraph113.Append(paragraphProperties69);
            paragraph113.Append(run70);

            A.Paragraph paragraph114 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties70 = new A.ParagraphProperties() { Level = 1 };

            A.Run run71 = new A.Run();
            A.RunProperties runProperties90 = new A.RunProperties() { Language = "en-US" };
            A.Text text90 = new A.Text();
            text90.Text = "Second level";

            run71.Append(runProperties90);
            run71.Append(text90);

            paragraph114.Append(paragraphProperties70);
            paragraph114.Append(run71);

            A.Paragraph paragraph115 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties71 = new A.ParagraphProperties() { Level = 2 };

            A.Run run72 = new A.Run();
            A.RunProperties runProperties91 = new A.RunProperties() { Language = "en-US" };
            A.Text text91 = new A.Text();
            text91.Text = "Third level";

            run72.Append(runProperties91);
            run72.Append(text91);

            paragraph115.Append(paragraphProperties71);
            paragraph115.Append(run72);

            A.Paragraph paragraph116 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties72 = new A.ParagraphProperties() { Level = 3 };

            A.Run run73 = new A.Run();
            A.RunProperties runProperties92 = new A.RunProperties() { Language = "en-US" };
            A.Text text92 = new A.Text();
            text92.Text = "Fourth level";

            run73.Append(runProperties92);
            run73.Append(text92);

            paragraph116.Append(paragraphProperties72);
            paragraph116.Append(run73);

            A.Paragraph paragraph117 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties73 = new A.ParagraphProperties() { Level = 4 };

            A.Run run74 = new A.Run();
            A.RunProperties runProperties93 = new A.RunProperties() { Language = "en-US" };
            A.Text text93 = new A.Text();
            text93.Text = "Fifth level";

            run74.Append(runProperties93);
            run74.Append(text93);

            paragraph117.Append(paragraphProperties73);
            paragraph117.Append(run74);

            textBody89.Append(bodyProperties91);
            textBody89.Append(listStyle91);
            textBody89.Append(paragraph113);
            textBody89.Append(paragraph114);
            textBody89.Append(paragraph115);
            textBody89.Append(paragraph116);
            textBody89.Append(paragraph117);

            shape69.Append(nonVisualShapeProperties69);
            shape69.Append(shapeProperties96);
            shape69.Append(textBody89);

            Shape shape70 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties70 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties121 = new NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Date Placeholder 6" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList97 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension97 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement123 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{7DA36AA5-BAF1-7357-DA0A-44D10F6E3A98}\" />");

            // nonVisualDrawingPropertiesExtension97.Append(openXmlUnknownElement123);

            nonVisualDrawingPropertiesExtensionList97.Append(nonVisualDrawingPropertiesExtension97);

            nonVisualDrawingProperties121.Append(nonVisualDrawingPropertiesExtensionList97);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties70 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks53 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties70.Append(shapeLocks53);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties121 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape53 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties121.Append(placeholderShape53);

            nonVisualShapeProperties70.Append(nonVisualDrawingProperties121);
            nonVisualShapeProperties70.Append(nonVisualShapeDrawingProperties70);
            nonVisualShapeProperties70.Append(applicationNonVisualDrawingProperties121);
            ShapeProperties shapeProperties97 = new ShapeProperties();

            TextBody textBody90 = new TextBody();
            A.BodyProperties bodyProperties92 = new A.BodyProperties();
            A.ListStyle listStyle92 = new A.ListStyle();

            A.Paragraph paragraph118 = new A.Paragraph();

            A.Field field20 = new A.Field() { Id = "{B5CF329D-DBF6-4E15-9822-D0979B6928D0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties94 = new A.RunProperties() { Language = "en-US" };
            runProperties94.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text94 = new A.Text();
            text94.Text = "4/24/2024";

            field20.Append(runProperties94);
            field20.Append(text94);
            A.EndParagraphRunProperties endParagraphRunProperties65 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph118.Append(field20);
            paragraph118.Append(endParagraphRunProperties65);

            textBody90.Append(bodyProperties92);
            textBody90.Append(listStyle92);
            textBody90.Append(paragraph118);

            shape70.Append(nonVisualShapeProperties70);
            shape70.Append(shapeProperties97);
            shape70.Append(textBody90);

            Shape shape71 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties71 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties122 = new NonVisualDrawingProperties() { Id = (UInt32Value)8U, Name = "Footer Placeholder 7" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList98 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension98 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement124 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{3EED7000-AFE3-02F3-FF13-0687A75AA37D}\" />");

            // nonVisualDrawingPropertiesExtension98.Append(openXmlUnknownElement124);

            nonVisualDrawingPropertiesExtensionList98.Append(nonVisualDrawingPropertiesExtension98);

            nonVisualDrawingProperties122.Append(nonVisualDrawingPropertiesExtensionList98);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties71 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks54 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties71.Append(shapeLocks54);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties122 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape54 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties122.Append(placeholderShape54);

            nonVisualShapeProperties71.Append(nonVisualDrawingProperties122);
            nonVisualShapeProperties71.Append(nonVisualShapeDrawingProperties71);
            nonVisualShapeProperties71.Append(applicationNonVisualDrawingProperties122);
            ShapeProperties shapeProperties98 = new ShapeProperties();

            TextBody textBody91 = new TextBody();
            A.BodyProperties bodyProperties93 = new A.BodyProperties();
            A.ListStyle listStyle93 = new A.ListStyle();

            A.Paragraph paragraph119 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties66 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph119.Append(endParagraphRunProperties66);

            textBody91.Append(bodyProperties93);
            textBody91.Append(listStyle93);
            textBody91.Append(paragraph119);

            shape71.Append(nonVisualShapeProperties71);
            shape71.Append(shapeProperties98);
            shape71.Append(textBody91);

            Shape shape72 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties72 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties123 = new NonVisualDrawingProperties() { Id = (UInt32Value)9U, Name = "Slide Number Placeholder 8" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList99 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension99 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement125 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{D316224B-3773-4AEE-09FE-524262EFFDB7}\" />");

            // nonVisualDrawingPropertiesExtension99.Append(openXmlUnknownElement125);

            nonVisualDrawingPropertiesExtensionList99.Append(nonVisualDrawingPropertiesExtension99);

            nonVisualDrawingProperties123.Append(nonVisualDrawingPropertiesExtensionList99);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties72 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks55 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties72.Append(shapeLocks55);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties123 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape55 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties123.Append(placeholderShape55);

            nonVisualShapeProperties72.Append(nonVisualDrawingProperties123);
            nonVisualShapeProperties72.Append(nonVisualShapeDrawingProperties72);
            nonVisualShapeProperties72.Append(applicationNonVisualDrawingProperties123);
            ShapeProperties shapeProperties99 = new ShapeProperties();

            TextBody textBody92 = new TextBody();
            A.BodyProperties bodyProperties94 = new A.BodyProperties();
            A.ListStyle listStyle94 = new A.ListStyle();

            A.Paragraph paragraph120 = new A.Paragraph();

            A.Field field21 = new A.Field() { Id = "{19E6023F-2333-4A80-AFF1-A7F70B52B230}", Type = "slidenum" };

            A.RunProperties runProperties95 = new A.RunProperties() { Language = "en-US" };
            runProperties95.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text95 = new A.Text();
            text95.Text = "‹#›";

            field21.Append(runProperties95);
            field21.Append(text95);
            A.EndParagraphRunProperties endParagraphRunProperties67 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph120.Append(field21);
            paragraph120.Append(endParagraphRunProperties67);

            textBody92.Append(bodyProperties94);
            textBody92.Append(listStyle94);
            textBody92.Append(paragraph120);

            shape72.Append(nonVisualShapeProperties72);
            shape72.Append(shapeProperties99);
            shape72.Append(textBody92);

            shapeTree13.Append(nonVisualGroupShapeProperties25);
            shapeTree13.Append(groupShapeProperties25);
            shapeTree13.Append(shape65);
            shapeTree13.Append(shape66);
            shapeTree13.Append(shape67);
            shapeTree13.Append(shape68);
            shapeTree13.Append(shape69);
            shapeTree13.Append(shape70);
            shapeTree13.Append(shape71);
            shapeTree13.Append(shape72);

            CommonSlideDataExtensionList commonSlideDataExtensionList13 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension13 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId13 = new P14.CreationId() { Val = (UInt32Value)1377362199U };
            creationId13.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension13.Append(creationId13);

            commonSlideDataExtensionList13.Append(commonSlideDataExtension13);

            commonSlideData13.Append(shapeTree13);
            commonSlideData13.Append(commonSlideDataExtensionList13);

            ColorMapOverride colorMapOverride11 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping11 = new A.MasterColorMapping();

            colorMapOverride11.Append(masterColorMapping11);

            slideLayout9.Append(commonSlideData13);
            slideLayout9.Append(colorMapOverride11);

            slideLayoutPart9.SlideLayout = slideLayout9;
        }

        // Generates content of slideLayoutPart10.
        private void GenerateSlideLayoutPart10Content(SlideLayoutPart slideLayoutPart10)
        {
            SlideLayout slideLayout10 = new SlideLayout() { Type = SlideLayoutValues.VerticalText, Preserve = true };
            slideLayout10.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout10.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout10.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData14 = new CommonSlideData() { Name = "Title and Vertical Text" };

            ShapeTree shapeTree14 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties26 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties124 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties26 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties124 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties26.Append(nonVisualDrawingProperties124);
            nonVisualGroupShapeProperties26.Append(nonVisualGroupShapeDrawingProperties26);
            nonVisualGroupShapeProperties26.Append(applicationNonVisualDrawingProperties124);

            GroupShapeProperties groupShapeProperties26 = new GroupShapeProperties();

            A.TransformGroup transformGroup26 = new A.TransformGroup();
            A.Offset offset94 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents94 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset26 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents26 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup26.Append(offset94);
            transformGroup26.Append(extents94);
            transformGroup26.Append(childOffset26);
            transformGroup26.Append(childExtents26);

            groupShapeProperties26.Append(transformGroup26);

            Shape shape73 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties73 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties125 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList100 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension100 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement126 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{A233F505-B8E9-1C2D-4718-A67167FDC442}\" />");

            // nonVisualDrawingPropertiesExtension100.Append(openXmlUnknownElement126);

            nonVisualDrawingPropertiesExtensionList100.Append(nonVisualDrawingPropertiesExtension100);

            nonVisualDrawingProperties125.Append(nonVisualDrawingPropertiesExtensionList100);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties73 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks56 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties73.Append(shapeLocks56);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties125 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape56 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties125.Append(placeholderShape56);

            nonVisualShapeProperties73.Append(nonVisualDrawingProperties125);
            nonVisualShapeProperties73.Append(nonVisualShapeDrawingProperties73);
            nonVisualShapeProperties73.Append(applicationNonVisualDrawingProperties125);
            ShapeProperties shapeProperties100 = new ShapeProperties();

            TextBody textBody93 = new TextBody();
            A.BodyProperties bodyProperties95 = new A.BodyProperties();
            A.ListStyle listStyle95 = new A.ListStyle();

            A.Paragraph paragraph121 = new A.Paragraph();

            A.Run run75 = new A.Run();
            A.RunProperties runProperties96 = new A.RunProperties() { Language = "en-US" };
            A.Text text96 = new A.Text();
            text96.Text = "Click to edit Master title style";

            run75.Append(runProperties96);
            run75.Append(text96);

            paragraph121.Append(run75);

            textBody93.Append(bodyProperties95);
            textBody93.Append(listStyle95);
            textBody93.Append(paragraph121);

            shape73.Append(nonVisualShapeProperties73);
            shape73.Append(shapeProperties100);
            shape73.Append(textBody93);

            Shape shape74 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties74 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties126 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Vertical Text Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList101 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension101 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement127 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{B11A60DF-6161-392C-E676-0544AC24975C}\" />");

            // nonVisualDrawingPropertiesExtension101.Append(openXmlUnknownElement127);

            nonVisualDrawingPropertiesExtensionList101.Append(nonVisualDrawingPropertiesExtension101);

            nonVisualDrawingProperties126.Append(nonVisualDrawingPropertiesExtensionList101);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties74 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks57 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties74.Append(shapeLocks57);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties126 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape57 = new PlaceholderShape() { Type = PlaceholderValues.Body, Orientation = DirectionValues.Vertical, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties126.Append(placeholderShape57);

            nonVisualShapeProperties74.Append(nonVisualDrawingProperties126);
            nonVisualShapeProperties74.Append(nonVisualShapeDrawingProperties74);
            nonVisualShapeProperties74.Append(applicationNonVisualDrawingProperties126);
            ShapeProperties shapeProperties101 = new ShapeProperties();

            TextBody textBody94 = new TextBody();
            A.BodyProperties bodyProperties96 = new A.BodyProperties() { Vertical = A.TextVerticalValues.EastAsianVetical };
            A.ListStyle listStyle96 = new A.ListStyle();

            A.Paragraph paragraph122 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties74 = new A.ParagraphProperties() { Level = 0 };

            A.Run run76 = new A.Run();
            A.RunProperties runProperties97 = new A.RunProperties() { Language = "en-US" };
            A.Text text97 = new A.Text();
            text97.Text = "Click to edit Master text styles";

            run76.Append(runProperties97);
            run76.Append(text97);

            paragraph122.Append(paragraphProperties74);
            paragraph122.Append(run76);

            A.Paragraph paragraph123 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties75 = new A.ParagraphProperties() { Level = 1 };

            A.Run run77 = new A.Run();
            A.RunProperties runProperties98 = new A.RunProperties() { Language = "en-US" };
            A.Text text98 = new A.Text();
            text98.Text = "Second level";

            run77.Append(runProperties98);
            run77.Append(text98);

            paragraph123.Append(paragraphProperties75);
            paragraph123.Append(run77);

            A.Paragraph paragraph124 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties76 = new A.ParagraphProperties() { Level = 2 };

            A.Run run78 = new A.Run();
            A.RunProperties runProperties99 = new A.RunProperties() { Language = "en-US" };
            A.Text text99 = new A.Text();
            text99.Text = "Third level";

            run78.Append(runProperties99);
            run78.Append(text99);

            paragraph124.Append(paragraphProperties76);
            paragraph124.Append(run78);

            A.Paragraph paragraph125 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties77 = new A.ParagraphProperties() { Level = 3 };

            A.Run run79 = new A.Run();
            A.RunProperties runProperties100 = new A.RunProperties() { Language = "en-US" };
            A.Text text100 = new A.Text();
            text100.Text = "Fourth level";

            run79.Append(runProperties100);
            run79.Append(text100);

            paragraph125.Append(paragraphProperties77);
            paragraph125.Append(run79);

            A.Paragraph paragraph126 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties78 = new A.ParagraphProperties() { Level = 4 };

            A.Run run80 = new A.Run();
            A.RunProperties runProperties101 = new A.RunProperties() { Language = "en-US" };
            A.Text text101 = new A.Text();
            text101.Text = "Fifth level";

            run80.Append(runProperties101);
            run80.Append(text101);

            paragraph126.Append(paragraphProperties78);
            paragraph126.Append(run80);

            textBody94.Append(bodyProperties96);
            textBody94.Append(listStyle96);
            textBody94.Append(paragraph122);
            textBody94.Append(paragraph123);
            textBody94.Append(paragraph124);
            textBody94.Append(paragraph125);
            textBody94.Append(paragraph126);

            shape74.Append(nonVisualShapeProperties74);
            shape74.Append(shapeProperties101);
            shape74.Append(textBody94);

            Shape shape75 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties75 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties127 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Date Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList102 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension102 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement128 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{20D0331A-E739-9246-E031-333F8B59055E}\" />");

            // nonVisualDrawingPropertiesExtension102.Append(openXmlUnknownElement128);

            nonVisualDrawingPropertiesExtensionList102.Append(nonVisualDrawingPropertiesExtension102);

            nonVisualDrawingProperties127.Append(nonVisualDrawingPropertiesExtensionList102);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties75 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks58 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties75.Append(shapeLocks58);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties127 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape58 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties127.Append(placeholderShape58);

            nonVisualShapeProperties75.Append(nonVisualDrawingProperties127);
            nonVisualShapeProperties75.Append(nonVisualShapeDrawingProperties75);
            nonVisualShapeProperties75.Append(applicationNonVisualDrawingProperties127);
            ShapeProperties shapeProperties102 = new ShapeProperties();

            TextBody textBody95 = new TextBody();
            A.BodyProperties bodyProperties97 = new A.BodyProperties();
            A.ListStyle listStyle97 = new A.ListStyle();

            A.Paragraph paragraph127 = new A.Paragraph();

            A.Field field22 = new A.Field() { Id = "{B5CF329D-DBF6-4E15-9822-D0979B6928D0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties102 = new A.RunProperties() { Language = "en-US" };
            runProperties102.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text102 = new A.Text();
            text102.Text = "4/24/2024";

            field22.Append(runProperties102);
            field22.Append(text102);
            A.EndParagraphRunProperties endParagraphRunProperties68 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph127.Append(field22);
            paragraph127.Append(endParagraphRunProperties68);

            textBody95.Append(bodyProperties97);
            textBody95.Append(listStyle97);
            textBody95.Append(paragraph127);

            shape75.Append(nonVisualShapeProperties75);
            shape75.Append(shapeProperties102);
            shape75.Append(textBody95);

            Shape shape76 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties76 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties128 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Footer Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList103 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension103 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement129 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{DAC61614-CC95-18F6-9983-82595DF4C6FA}\" />");

            // nonVisualDrawingPropertiesExtension103.Append(openXmlUnknownElement129);

            nonVisualDrawingPropertiesExtensionList103.Append(nonVisualDrawingPropertiesExtension103);

            nonVisualDrawingProperties128.Append(nonVisualDrawingPropertiesExtensionList103);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties76 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks59 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties76.Append(shapeLocks59);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties128 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape59 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties128.Append(placeholderShape59);

            nonVisualShapeProperties76.Append(nonVisualDrawingProperties128);
            nonVisualShapeProperties76.Append(nonVisualShapeDrawingProperties76);
            nonVisualShapeProperties76.Append(applicationNonVisualDrawingProperties128);
            ShapeProperties shapeProperties103 = new ShapeProperties();

            TextBody textBody96 = new TextBody();
            A.BodyProperties bodyProperties98 = new A.BodyProperties();
            A.ListStyle listStyle98 = new A.ListStyle();

            A.Paragraph paragraph128 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties69 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph128.Append(endParagraphRunProperties69);

            textBody96.Append(bodyProperties98);
            textBody96.Append(listStyle98);
            textBody96.Append(paragraph128);

            shape76.Append(nonVisualShapeProperties76);
            shape76.Append(shapeProperties103);
            shape76.Append(textBody96);

            Shape shape77 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties77 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties129 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Slide Number Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList104 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension104 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement130 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{EEAEA546-2D88-376A-57A2-EBDE9905D79E}\" />");

            // nonVisualDrawingPropertiesExtension104.Append(openXmlUnknownElement130);

            nonVisualDrawingPropertiesExtensionList104.Append(nonVisualDrawingPropertiesExtension104);

            nonVisualDrawingProperties129.Append(nonVisualDrawingPropertiesExtensionList104);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties77 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks60 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties77.Append(shapeLocks60);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties129 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape60 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties129.Append(placeholderShape60);

            nonVisualShapeProperties77.Append(nonVisualDrawingProperties129);
            nonVisualShapeProperties77.Append(nonVisualShapeDrawingProperties77);
            nonVisualShapeProperties77.Append(applicationNonVisualDrawingProperties129);
            ShapeProperties shapeProperties104 = new ShapeProperties();

            TextBody textBody97 = new TextBody();
            A.BodyProperties bodyProperties99 = new A.BodyProperties();
            A.ListStyle listStyle99 = new A.ListStyle();

            A.Paragraph paragraph129 = new A.Paragraph();

            A.Field field23 = new A.Field() { Id = "{19E6023F-2333-4A80-AFF1-A7F70B52B230}", Type = "slidenum" };

            A.RunProperties runProperties103 = new A.RunProperties() { Language = "en-US" };
            runProperties103.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text103 = new A.Text();
            text103.Text = "‹#›";

            field23.Append(runProperties103);
            field23.Append(text103);
            A.EndParagraphRunProperties endParagraphRunProperties70 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph129.Append(field23);
            paragraph129.Append(endParagraphRunProperties70);

            textBody97.Append(bodyProperties99);
            textBody97.Append(listStyle99);
            textBody97.Append(paragraph129);

            shape77.Append(nonVisualShapeProperties77);
            shape77.Append(shapeProperties104);
            shape77.Append(textBody97);

            shapeTree14.Append(nonVisualGroupShapeProperties26);
            shapeTree14.Append(groupShapeProperties26);
            shapeTree14.Append(shape73);
            shapeTree14.Append(shape74);
            shapeTree14.Append(shape75);
            shapeTree14.Append(shape76);
            shapeTree14.Append(shape77);

            CommonSlideDataExtensionList commonSlideDataExtensionList14 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension14 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId14 = new P14.CreationId() { Val = (UInt32Value)3954769401U };
            creationId14.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension14.Append(creationId14);

            commonSlideDataExtensionList14.Append(commonSlideDataExtension14);

            commonSlideData14.Append(shapeTree14);
            commonSlideData14.Append(commonSlideDataExtensionList14);

            ColorMapOverride colorMapOverride12 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping12 = new A.MasterColorMapping();

            colorMapOverride12.Append(masterColorMapping12);

            slideLayout10.Append(commonSlideData14);
            slideLayout10.Append(colorMapOverride12);

            slideLayoutPart10.SlideLayout = slideLayout10;
        }

        // Generates content of slideLayoutPart11.
        private void GenerateSlideLayoutPart11Content(SlideLayoutPart slideLayoutPart11)
        {
            SlideLayout slideLayout11 = new SlideLayout() { Type = SlideLayoutValues.TwoObjects, Preserve = true };
            slideLayout11.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout11.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout11.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData15 = new CommonSlideData() { Name = "Two Content" };

            ShapeTree shapeTree15 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties27 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties130 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties27 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties130 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties27.Append(nonVisualDrawingProperties130);
            nonVisualGroupShapeProperties27.Append(nonVisualGroupShapeDrawingProperties27);
            nonVisualGroupShapeProperties27.Append(applicationNonVisualDrawingProperties130);

            GroupShapeProperties groupShapeProperties27 = new GroupShapeProperties();

            A.TransformGroup transformGroup27 = new A.TransformGroup();
            A.Offset offset95 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents95 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset27 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents27 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup27.Append(offset95);
            transformGroup27.Append(extents95);
            transformGroup27.Append(childOffset27);
            transformGroup27.Append(childExtents27);

            groupShapeProperties27.Append(transformGroup27);

            Shape shape78 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties78 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties131 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList105 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension105 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement131 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{24F6859F-777A-D154-9DA1-4AC9CB4429CB}\" />");

            // nonVisualDrawingPropertiesExtension105.Append(openXmlUnknownElement131);

            nonVisualDrawingPropertiesExtensionList105.Append(nonVisualDrawingPropertiesExtension105);

            nonVisualDrawingProperties131.Append(nonVisualDrawingPropertiesExtensionList105);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties78 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks61 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties78.Append(shapeLocks61);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties131 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape61 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties131.Append(placeholderShape61);

            nonVisualShapeProperties78.Append(nonVisualDrawingProperties131);
            nonVisualShapeProperties78.Append(nonVisualShapeDrawingProperties78);
            nonVisualShapeProperties78.Append(applicationNonVisualDrawingProperties131);
            ShapeProperties shapeProperties105 = new ShapeProperties();

            TextBody textBody98 = new TextBody();
            A.BodyProperties bodyProperties100 = new A.BodyProperties();
            A.ListStyle listStyle100 = new A.ListStyle();

            A.Paragraph paragraph130 = new A.Paragraph();

            A.Run run81 = new A.Run();
            A.RunProperties runProperties104 = new A.RunProperties() { Language = "en-US" };
            A.Text text104 = new A.Text();
            text104.Text = "Click to edit Master title style";

            run81.Append(runProperties104);
            run81.Append(text104);

            paragraph130.Append(run81);

            textBody98.Append(bodyProperties100);
            textBody98.Append(listStyle100);
            textBody98.Append(paragraph130);

            shape78.Append(nonVisualShapeProperties78);
            shape78.Append(shapeProperties105);
            shape78.Append(textBody98);

            Shape shape79 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties79 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties132 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Content Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList106 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension106 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement132 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{66A4C86A-71ED-68B0-490A-C720C7EA691C}\" />");

            // nonVisualDrawingPropertiesExtension106.Append(openXmlUnknownElement132);

            nonVisualDrawingPropertiesExtensionList106.Append(nonVisualDrawingPropertiesExtension106);

            nonVisualDrawingProperties132.Append(nonVisualDrawingPropertiesExtensionList106);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties79 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks62 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties79.Append(shapeLocks62);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties132 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape62 = new PlaceholderShape() { Size = PlaceholderSizeValues.Half, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties132.Append(placeholderShape62);

            nonVisualShapeProperties79.Append(nonVisualDrawingProperties132);
            nonVisualShapeProperties79.Append(nonVisualShapeDrawingProperties79);
            nonVisualShapeProperties79.Append(applicationNonVisualDrawingProperties132);

            ShapeProperties shapeProperties106 = new ShapeProperties();

            A.Transform2D transform2D68 = new A.Transform2D();
            A.Offset offset96 = new A.Offset() { X = 838200L, Y = 1825625L };
            A.Extents extents96 = new A.Extents() { Cx = 5181600L, Cy = 4351338L };

            transform2D68.Append(offset96);
            transform2D68.Append(extents96);

            shapeProperties106.Append(transform2D68);

            TextBody textBody99 = new TextBody();
            A.BodyProperties bodyProperties101 = new A.BodyProperties();
            A.ListStyle listStyle101 = new A.ListStyle();

            A.Paragraph paragraph131 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties79 = new A.ParagraphProperties() { Level = 0 };

            A.Run run82 = new A.Run();
            A.RunProperties runProperties105 = new A.RunProperties() { Language = "en-US" };
            A.Text text105 = new A.Text();
            text105.Text = "Click to edit Master text styles";

            run82.Append(runProperties105);
            run82.Append(text105);

            paragraph131.Append(paragraphProperties79);
            paragraph131.Append(run82);

            A.Paragraph paragraph132 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties80 = new A.ParagraphProperties() { Level = 1 };

            A.Run run83 = new A.Run();
            A.RunProperties runProperties106 = new A.RunProperties() { Language = "en-US" };
            A.Text text106 = new A.Text();
            text106.Text = "Second level";

            run83.Append(runProperties106);
            run83.Append(text106);

            paragraph132.Append(paragraphProperties80);
            paragraph132.Append(run83);

            A.Paragraph paragraph133 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties81 = new A.ParagraphProperties() { Level = 2 };

            A.Run run84 = new A.Run();
            A.RunProperties runProperties107 = new A.RunProperties() { Language = "en-US" };
            A.Text text107 = new A.Text();
            text107.Text = "Third level";

            run84.Append(runProperties107);
            run84.Append(text107);

            paragraph133.Append(paragraphProperties81);
            paragraph133.Append(run84);

            A.Paragraph paragraph134 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties82 = new A.ParagraphProperties() { Level = 3 };

            A.Run run85 = new A.Run();
            A.RunProperties runProperties108 = new A.RunProperties() { Language = "en-US" };
            A.Text text108 = new A.Text();
            text108.Text = "Fourth level";

            run85.Append(runProperties108);
            run85.Append(text108);

            paragraph134.Append(paragraphProperties82);
            paragraph134.Append(run85);

            A.Paragraph paragraph135 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties83 = new A.ParagraphProperties() { Level = 4 };

            A.Run run86 = new A.Run();
            A.RunProperties runProperties109 = new A.RunProperties() { Language = "en-US" };
            A.Text text109 = new A.Text();
            text109.Text = "Fifth level";

            run86.Append(runProperties109);
            run86.Append(text109);

            paragraph135.Append(paragraphProperties83);
            paragraph135.Append(run86);

            textBody99.Append(bodyProperties101);
            textBody99.Append(listStyle101);
            textBody99.Append(paragraph131);
            textBody99.Append(paragraph132);
            textBody99.Append(paragraph133);
            textBody99.Append(paragraph134);
            textBody99.Append(paragraph135);

            shape79.Append(nonVisualShapeProperties79);
            shape79.Append(shapeProperties106);
            shape79.Append(textBody99);

            Shape shape80 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties80 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties133 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Content Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList107 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension107 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement133 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{022AB907-F439-9F67-54F2-D9153D15DA5D}\" />");

            // nonVisualDrawingPropertiesExtension107.Append(openXmlUnknownElement133);

            nonVisualDrawingPropertiesExtensionList107.Append(nonVisualDrawingPropertiesExtension107);

            nonVisualDrawingProperties133.Append(nonVisualDrawingPropertiesExtensionList107);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties80 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks63 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties80.Append(shapeLocks63);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties133 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape63 = new PlaceholderShape() { Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties133.Append(placeholderShape63);

            nonVisualShapeProperties80.Append(nonVisualDrawingProperties133);
            nonVisualShapeProperties80.Append(nonVisualShapeDrawingProperties80);
            nonVisualShapeProperties80.Append(applicationNonVisualDrawingProperties133);

            ShapeProperties shapeProperties107 = new ShapeProperties();

            A.Transform2D transform2D69 = new A.Transform2D();
            A.Offset offset97 = new A.Offset() { X = 6172200L, Y = 1825625L };
            A.Extents extents97 = new A.Extents() { Cx = 5181600L, Cy = 4351338L };

            transform2D69.Append(offset97);
            transform2D69.Append(extents97);

            shapeProperties107.Append(transform2D69);

            TextBody textBody100 = new TextBody();
            A.BodyProperties bodyProperties102 = new A.BodyProperties();
            A.ListStyle listStyle102 = new A.ListStyle();

            A.Paragraph paragraph136 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties84 = new A.ParagraphProperties() { Level = 0 };

            A.Run run87 = new A.Run();
            A.RunProperties runProperties110 = new A.RunProperties() { Language = "en-US" };
            A.Text text110 = new A.Text();
            text110.Text = "Click to edit Master text styles";

            run87.Append(runProperties110);
            run87.Append(text110);

            paragraph136.Append(paragraphProperties84);
            paragraph136.Append(run87);

            A.Paragraph paragraph137 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties85 = new A.ParagraphProperties() { Level = 1 };

            A.Run run88 = new A.Run();
            A.RunProperties runProperties111 = new A.RunProperties() { Language = "en-US" };
            A.Text text111 = new A.Text();
            text111.Text = "Second level";

            run88.Append(runProperties111);
            run88.Append(text111);

            paragraph137.Append(paragraphProperties85);
            paragraph137.Append(run88);

            A.Paragraph paragraph138 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties86 = new A.ParagraphProperties() { Level = 2 };

            A.Run run89 = new A.Run();
            A.RunProperties runProperties112 = new A.RunProperties() { Language = "en-US" };
            A.Text text112 = new A.Text();
            text112.Text = "Third level";

            run89.Append(runProperties112);
            run89.Append(text112);

            paragraph138.Append(paragraphProperties86);
            paragraph138.Append(run89);

            A.Paragraph paragraph139 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties87 = new A.ParagraphProperties() { Level = 3 };

            A.Run run90 = new A.Run();
            A.RunProperties runProperties113 = new A.RunProperties() { Language = "en-US" };
            A.Text text113 = new A.Text();
            text113.Text = "Fourth level";

            run90.Append(runProperties113);
            run90.Append(text113);

            paragraph139.Append(paragraphProperties87);
            paragraph139.Append(run90);

            A.Paragraph paragraph140 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties88 = new A.ParagraphProperties() { Level = 4 };

            A.Run run91 = new A.Run();
            A.RunProperties runProperties114 = new A.RunProperties() { Language = "en-US" };
            A.Text text114 = new A.Text();
            text114.Text = "Fifth level";

            run91.Append(runProperties114);
            run91.Append(text114);

            paragraph140.Append(paragraphProperties88);
            paragraph140.Append(run91);

            textBody100.Append(bodyProperties102);
            textBody100.Append(listStyle102);
            textBody100.Append(paragraph136);
            textBody100.Append(paragraph137);
            textBody100.Append(paragraph138);
            textBody100.Append(paragraph139);
            textBody100.Append(paragraph140);

            shape80.Append(nonVisualShapeProperties80);
            shape80.Append(shapeProperties107);
            shape80.Append(textBody100);

            Shape shape81 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties81 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties134 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Date Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList108 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension108 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement134 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{CA2AFAFA-16C4-1FA0-730C-871F065D1426}\" />");

            // nonVisualDrawingPropertiesExtension108.Append(openXmlUnknownElement134);

            nonVisualDrawingPropertiesExtensionList108.Append(nonVisualDrawingPropertiesExtension108);

            nonVisualDrawingProperties134.Append(nonVisualDrawingPropertiesExtensionList108);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties81 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks64 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties81.Append(shapeLocks64);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties134 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape64 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties134.Append(placeholderShape64);

            nonVisualShapeProperties81.Append(nonVisualDrawingProperties134);
            nonVisualShapeProperties81.Append(nonVisualShapeDrawingProperties81);
            nonVisualShapeProperties81.Append(applicationNonVisualDrawingProperties134);
            ShapeProperties shapeProperties108 = new ShapeProperties();

            TextBody textBody101 = new TextBody();
            A.BodyProperties bodyProperties103 = new A.BodyProperties();
            A.ListStyle listStyle103 = new A.ListStyle();

            A.Paragraph paragraph141 = new A.Paragraph();

            A.Field field24 = new A.Field() { Id = "{B5CF329D-DBF6-4E15-9822-D0979B6928D0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties115 = new A.RunProperties() { Language = "en-US" };
            runProperties115.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text115 = new A.Text();
            text115.Text = "4/24/2024";

            field24.Append(runProperties115);
            field24.Append(text115);
            A.EndParagraphRunProperties endParagraphRunProperties71 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph141.Append(field24);
            paragraph141.Append(endParagraphRunProperties71);

            textBody101.Append(bodyProperties103);
            textBody101.Append(listStyle103);
            textBody101.Append(paragraph141);

            shape81.Append(nonVisualShapeProperties81);
            shape81.Append(shapeProperties108);
            shape81.Append(textBody101);

            Shape shape82 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties82 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties135 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Footer Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList109 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension109 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement135 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{E7D92203-DBA4-0864-E05F-4CF8E8654EDD}\" />");

            // nonVisualDrawingPropertiesExtension109.Append(openXmlUnknownElement135);

            nonVisualDrawingPropertiesExtensionList109.Append(nonVisualDrawingPropertiesExtension109);

            nonVisualDrawingProperties135.Append(nonVisualDrawingPropertiesExtensionList109);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties82 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks65 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties82.Append(shapeLocks65);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties135 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape65 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties135.Append(placeholderShape65);

            nonVisualShapeProperties82.Append(nonVisualDrawingProperties135);
            nonVisualShapeProperties82.Append(nonVisualShapeDrawingProperties82);
            nonVisualShapeProperties82.Append(applicationNonVisualDrawingProperties135);
            ShapeProperties shapeProperties109 = new ShapeProperties();

            TextBody textBody102 = new TextBody();
            A.BodyProperties bodyProperties104 = new A.BodyProperties();
            A.ListStyle listStyle104 = new A.ListStyle();

            A.Paragraph paragraph142 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties72 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph142.Append(endParagraphRunProperties72);

            textBody102.Append(bodyProperties104);
            textBody102.Append(listStyle104);
            textBody102.Append(paragraph142);

            shape82.Append(nonVisualShapeProperties82);
            shape82.Append(shapeProperties109);
            shape82.Append(textBody102);

            Shape shape83 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties83 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties136 = new NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Slide Number Placeholder 6" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList110 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension110 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement136 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{1EF211E8-A7A3-767A-5975-4A982C1F343A}\" />");

            // nonVisualDrawingPropertiesExtension110.Append(openXmlUnknownElement136);

            nonVisualDrawingPropertiesExtensionList110.Append(nonVisualDrawingPropertiesExtension110);

            nonVisualDrawingProperties136.Append(nonVisualDrawingPropertiesExtensionList110);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties83 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks66 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties83.Append(shapeLocks66);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties136 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape66 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties136.Append(placeholderShape66);

            nonVisualShapeProperties83.Append(nonVisualDrawingProperties136);
            nonVisualShapeProperties83.Append(nonVisualShapeDrawingProperties83);
            nonVisualShapeProperties83.Append(applicationNonVisualDrawingProperties136);
            ShapeProperties shapeProperties110 = new ShapeProperties();

            TextBody textBody103 = new TextBody();
            A.BodyProperties bodyProperties105 = new A.BodyProperties();
            A.ListStyle listStyle105 = new A.ListStyle();

            A.Paragraph paragraph143 = new A.Paragraph();

            A.Field field25 = new A.Field() { Id = "{19E6023F-2333-4A80-AFF1-A7F70B52B230}", Type = "slidenum" };

            A.RunProperties runProperties116 = new A.RunProperties() { Language = "en-US" };
            runProperties116.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text116 = new A.Text();
            text116.Text = "‹#›";

            field25.Append(runProperties116);
            field25.Append(text116);
            A.EndParagraphRunProperties endParagraphRunProperties73 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph143.Append(field25);
            paragraph143.Append(endParagraphRunProperties73);

            textBody103.Append(bodyProperties105);
            textBody103.Append(listStyle105);
            textBody103.Append(paragraph143);

            shape83.Append(nonVisualShapeProperties83);
            shape83.Append(shapeProperties110);
            shape83.Append(textBody103);

            shapeTree15.Append(nonVisualGroupShapeProperties27);
            shapeTree15.Append(groupShapeProperties27);
            shapeTree15.Append(shape78);
            shapeTree15.Append(shape79);
            shapeTree15.Append(shape80);
            shapeTree15.Append(shape81);
            shapeTree15.Append(shape82);
            shapeTree15.Append(shape83);

            CommonSlideDataExtensionList commonSlideDataExtensionList15 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension15 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId15 = new P14.CreationId() { Val = (UInt32Value)2017446292U };
            creationId15.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension15.Append(creationId15);

            commonSlideDataExtensionList15.Append(commonSlideDataExtension15);

            commonSlideData15.Append(shapeTree15);
            commonSlideData15.Append(commonSlideDataExtensionList15);

            ColorMapOverride colorMapOverride13 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping13 = new A.MasterColorMapping();

            colorMapOverride13.Append(masterColorMapping13);

            slideLayout11.Append(commonSlideData15);
            slideLayout11.Append(colorMapOverride13);

            slideLayoutPart11.SlideLayout = slideLayout11;
        }

        // Generates content of slideLayoutPart12.
        private void GenerateSlideLayoutPart12Content(SlideLayoutPart slideLayoutPart12)
        {
            SlideLayout slideLayout12 = new SlideLayout() { Type = SlideLayoutValues.PictureText, Preserve = true };
            slideLayout12.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slideLayout12.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slideLayout12.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            CommonSlideData commonSlideData16 = new CommonSlideData() { Name = "Picture with Caption" };

            ShapeTree shapeTree16 = new ShapeTree();

            NonVisualGroupShapeProperties nonVisualGroupShapeProperties28 = new NonVisualGroupShapeProperties();
            NonVisualDrawingProperties nonVisualDrawingProperties137 = new NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" };
            NonVisualGroupShapeDrawingProperties nonVisualGroupShapeDrawingProperties28 = new NonVisualGroupShapeDrawingProperties();
            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties137 = new ApplicationNonVisualDrawingProperties();

            nonVisualGroupShapeProperties28.Append(nonVisualDrawingProperties137);
            nonVisualGroupShapeProperties28.Append(nonVisualGroupShapeDrawingProperties28);
            nonVisualGroupShapeProperties28.Append(applicationNonVisualDrawingProperties137);

            GroupShapeProperties groupShapeProperties28 = new GroupShapeProperties();

            A.TransformGroup transformGroup28 = new A.TransformGroup();
            A.Offset offset98 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents98 = new A.Extents() { Cx = 0L, Cy = 0L };
            A.ChildOffset childOffset28 = new A.ChildOffset() { X = 0L, Y = 0L };
            A.ChildExtents childExtents28 = new A.ChildExtents() { Cx = 0L, Cy = 0L };

            transformGroup28.Append(offset98);
            transformGroup28.Append(extents98);
            transformGroup28.Append(childOffset28);
            transformGroup28.Append(childExtents28);

            groupShapeProperties28.Append(transformGroup28);

            Shape shape84 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties84 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties138 = new NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList111 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension111 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement137 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{B2C3D960-A94E-DCD6-1C10-13257865D72B}\" />");

            // nonVisualDrawingPropertiesExtension111.Append(openXmlUnknownElement137);

            nonVisualDrawingPropertiesExtensionList111.Append(nonVisualDrawingPropertiesExtension111);

            nonVisualDrawingProperties138.Append(nonVisualDrawingPropertiesExtensionList111);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties84 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks67 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties84.Append(shapeLocks67);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties138 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape67 = new PlaceholderShape() { Type = PlaceholderValues.Title };

            applicationNonVisualDrawingProperties138.Append(placeholderShape67);

            nonVisualShapeProperties84.Append(nonVisualDrawingProperties138);
            nonVisualShapeProperties84.Append(nonVisualShapeDrawingProperties84);
            nonVisualShapeProperties84.Append(applicationNonVisualDrawingProperties138);

            ShapeProperties shapeProperties111 = new ShapeProperties();

            A.Transform2D transform2D70 = new A.Transform2D();
            A.Offset offset99 = new A.Offset() { X = 839788L, Y = 457200L };
            A.Extents extents99 = new A.Extents() { Cx = 3932237L, Cy = 1600200L };

            transform2D70.Append(offset99);
            transform2D70.Append(extents99);

            shapeProperties111.Append(transform2D70);

            TextBody textBody104 = new TextBody();
            A.BodyProperties bodyProperties106 = new A.BodyProperties() { Anchor = A.TextAnchoringTypeValues.Bottom };

            A.ListStyle listStyle106 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties22 = new A.Level1ParagraphProperties();
            A.DefaultRunProperties defaultRunProperties119 = new A.DefaultRunProperties() { FontSize = 3200 };

            level1ParagraphProperties22.Append(defaultRunProperties119);

            listStyle106.Append(level1ParagraphProperties22);

            A.Paragraph paragraph144 = new A.Paragraph();

            A.Run run92 = new A.Run();
            A.RunProperties runProperties117 = new A.RunProperties() { Language = "en-US" };
            A.Text text117 = new A.Text();
            text117.Text = "Click to edit Master title style";

            run92.Append(runProperties117);
            run92.Append(text117);

            paragraph144.Append(run92);

            textBody104.Append(bodyProperties106);
            textBody104.Append(listStyle106);
            textBody104.Append(paragraph144);

            shape84.Append(nonVisualShapeProperties84);
            shape84.Append(shapeProperties111);
            shape84.Append(textBody104);

            Shape shape85 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties85 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties139 = new NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Picture Placeholder 2" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList112 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension112 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement138 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{1F9E2A89-6FD5-0138-9EFF-69314194D830}\" />");

            // nonVisualDrawingPropertiesExtension112.Append(openXmlUnknownElement138);

            nonVisualDrawingPropertiesExtensionList112.Append(nonVisualDrawingPropertiesExtension112);

            nonVisualDrawingProperties139.Append(nonVisualDrawingPropertiesExtensionList112);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties85 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks68 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties85.Append(shapeLocks68);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties139 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape68 = new PlaceholderShape() { Type = PlaceholderValues.Picture, Index = (UInt32Value)1U };

            applicationNonVisualDrawingProperties139.Append(placeholderShape68);

            nonVisualShapeProperties85.Append(nonVisualDrawingProperties139);
            nonVisualShapeProperties85.Append(nonVisualShapeDrawingProperties85);
            nonVisualShapeProperties85.Append(applicationNonVisualDrawingProperties139);

            ShapeProperties shapeProperties112 = new ShapeProperties();

            A.Transform2D transform2D71 = new A.Transform2D();
            A.Offset offset100 = new A.Offset() { X = 5183188L, Y = 987425L };
            A.Extents extents100 = new A.Extents() { Cx = 6172200L, Cy = 4873625L };

            transform2D71.Append(offset100);
            transform2D71.Append(extents100);

            shapeProperties112.Append(transform2D71);

            TextBody textBody105 = new TextBody();
            A.BodyProperties bodyProperties107 = new A.BodyProperties();

            A.ListStyle listStyle107 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties23 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet61 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties120 = new A.DefaultRunProperties() { FontSize = 3200 };

            level1ParagraphProperties23.Append(noBullet61);
            level1ParagraphProperties23.Append(defaultRunProperties120);

            A.Level2ParagraphProperties level2ParagraphProperties11 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet62 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties121 = new A.DefaultRunProperties() { FontSize = 2800 };

            level2ParagraphProperties11.Append(noBullet62);
            level2ParagraphProperties11.Append(defaultRunProperties121);

            A.Level3ParagraphProperties level3ParagraphProperties11 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet63 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties122 = new A.DefaultRunProperties() { FontSize = 2400 };

            level3ParagraphProperties11.Append(noBullet63);
            level3ParagraphProperties11.Append(defaultRunProperties122);

            A.Level4ParagraphProperties level4ParagraphProperties11 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet64 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties123 = new A.DefaultRunProperties() { FontSize = 2000 };

            level4ParagraphProperties11.Append(noBullet64);
            level4ParagraphProperties11.Append(defaultRunProperties123);

            A.Level5ParagraphProperties level5ParagraphProperties11 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet65 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties124 = new A.DefaultRunProperties() { FontSize = 2000 };

            level5ParagraphProperties11.Append(noBullet65);
            level5ParagraphProperties11.Append(defaultRunProperties124);

            A.Level6ParagraphProperties level6ParagraphProperties11 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet66 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties125 = new A.DefaultRunProperties() { FontSize = 2000 };

            level6ParagraphProperties11.Append(noBullet66);
            level6ParagraphProperties11.Append(defaultRunProperties125);

            A.Level7ParagraphProperties level7ParagraphProperties11 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet67 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties126 = new A.DefaultRunProperties() { FontSize = 2000 };

            level7ParagraphProperties11.Append(noBullet67);
            level7ParagraphProperties11.Append(defaultRunProperties126);

            A.Level8ParagraphProperties level8ParagraphProperties11 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet68 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties127 = new A.DefaultRunProperties() { FontSize = 2000 };

            level8ParagraphProperties11.Append(noBullet68);
            level8ParagraphProperties11.Append(defaultRunProperties127);

            A.Level9ParagraphProperties level9ParagraphProperties11 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet69 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties128 = new A.DefaultRunProperties() { FontSize = 2000 };

            level9ParagraphProperties11.Append(noBullet69);
            level9ParagraphProperties11.Append(defaultRunProperties128);

            listStyle107.Append(level1ParagraphProperties23);
            listStyle107.Append(level2ParagraphProperties11);
            listStyle107.Append(level3ParagraphProperties11);
            listStyle107.Append(level4ParagraphProperties11);
            listStyle107.Append(level5ParagraphProperties11);
            listStyle107.Append(level6ParagraphProperties11);
            listStyle107.Append(level7ParagraphProperties11);
            listStyle107.Append(level8ParagraphProperties11);
            listStyle107.Append(level9ParagraphProperties11);

            A.Paragraph paragraph145 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties74 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph145.Append(endParagraphRunProperties74);

            textBody105.Append(bodyProperties107);
            textBody105.Append(listStyle107);
            textBody105.Append(paragraph145);

            shape85.Append(nonVisualShapeProperties85);
            shape85.Append(shapeProperties112);
            shape85.Append(textBody105);

            Shape shape86 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties86 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties140 = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Text Placeholder 3" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList113 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension113 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement139 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{50193971-EE06-16F6-1EEF-E1EEDE8BC035}\" />");

            // nonVisualDrawingPropertiesExtension113.Append(openXmlUnknownElement139);

            nonVisualDrawingPropertiesExtensionList113.Append(nonVisualDrawingPropertiesExtension113);

            nonVisualDrawingProperties140.Append(nonVisualDrawingPropertiesExtensionList113);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties86 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks69 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties86.Append(shapeLocks69);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties140 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape69 = new PlaceholderShape() { Type = PlaceholderValues.Body, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)2U };

            applicationNonVisualDrawingProperties140.Append(placeholderShape69);

            nonVisualShapeProperties86.Append(nonVisualDrawingProperties140);
            nonVisualShapeProperties86.Append(nonVisualShapeDrawingProperties86);
            nonVisualShapeProperties86.Append(applicationNonVisualDrawingProperties140);

            ShapeProperties shapeProperties113 = new ShapeProperties();

            A.Transform2D transform2D72 = new A.Transform2D();
            A.Offset offset101 = new A.Offset() { X = 839788L, Y = 2057400L };
            A.Extents extents101 = new A.Extents() { Cx = 3932237L, Cy = 3811588L };

            transform2D72.Append(offset101);
            transform2D72.Append(extents101);

            shapeProperties113.Append(transform2D72);

            TextBody textBody106 = new TextBody();
            A.BodyProperties bodyProperties108 = new A.BodyProperties();

            A.ListStyle listStyle108 = new A.ListStyle();

            A.Level1ParagraphProperties level1ParagraphProperties24 = new A.Level1ParagraphProperties() { LeftMargin = 0, Indent = 0 };
            A.NoBullet noBullet70 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties129 = new A.DefaultRunProperties() { FontSize = 1600 };

            level1ParagraphProperties24.Append(noBullet70);
            level1ParagraphProperties24.Append(defaultRunProperties129);

            A.Level2ParagraphProperties level2ParagraphProperties12 = new A.Level2ParagraphProperties() { LeftMargin = 457200, Indent = 0 };
            A.NoBullet noBullet71 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties130 = new A.DefaultRunProperties() { FontSize = 1400 };

            level2ParagraphProperties12.Append(noBullet71);
            level2ParagraphProperties12.Append(defaultRunProperties130);

            A.Level3ParagraphProperties level3ParagraphProperties12 = new A.Level3ParagraphProperties() { LeftMargin = 914400, Indent = 0 };
            A.NoBullet noBullet72 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties131 = new A.DefaultRunProperties() { FontSize = 1200 };

            level3ParagraphProperties12.Append(noBullet72);
            level3ParagraphProperties12.Append(defaultRunProperties131);

            A.Level4ParagraphProperties level4ParagraphProperties12 = new A.Level4ParagraphProperties() { LeftMargin = 1371600, Indent = 0 };
            A.NoBullet noBullet73 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties132 = new A.DefaultRunProperties() { FontSize = 1000 };

            level4ParagraphProperties12.Append(noBullet73);
            level4ParagraphProperties12.Append(defaultRunProperties132);

            A.Level5ParagraphProperties level5ParagraphProperties12 = new A.Level5ParagraphProperties() { LeftMargin = 1828800, Indent = 0 };
            A.NoBullet noBullet74 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties133 = new A.DefaultRunProperties() { FontSize = 1000 };

            level5ParagraphProperties12.Append(noBullet74);
            level5ParagraphProperties12.Append(defaultRunProperties133);

            A.Level6ParagraphProperties level6ParagraphProperties12 = new A.Level6ParagraphProperties() { LeftMargin = 2286000, Indent = 0 };
            A.NoBullet noBullet75 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties134 = new A.DefaultRunProperties() { FontSize = 1000 };

            level6ParagraphProperties12.Append(noBullet75);
            level6ParagraphProperties12.Append(defaultRunProperties134);

            A.Level7ParagraphProperties level7ParagraphProperties12 = new A.Level7ParagraphProperties() { LeftMargin = 2743200, Indent = 0 };
            A.NoBullet noBullet76 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties135 = new A.DefaultRunProperties() { FontSize = 1000 };

            level7ParagraphProperties12.Append(noBullet76);
            level7ParagraphProperties12.Append(defaultRunProperties135);

            A.Level8ParagraphProperties level8ParagraphProperties12 = new A.Level8ParagraphProperties() { LeftMargin = 3200400, Indent = 0 };
            A.NoBullet noBullet77 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties136 = new A.DefaultRunProperties() { FontSize = 1000 };

            level8ParagraphProperties12.Append(noBullet77);
            level8ParagraphProperties12.Append(defaultRunProperties136);

            A.Level9ParagraphProperties level9ParagraphProperties12 = new A.Level9ParagraphProperties() { LeftMargin = 3657600, Indent = 0 };
            A.NoBullet noBullet78 = new A.NoBullet();
            A.DefaultRunProperties defaultRunProperties137 = new A.DefaultRunProperties() { FontSize = 1000 };

            level9ParagraphProperties12.Append(noBullet78);
            level9ParagraphProperties12.Append(defaultRunProperties137);

            listStyle108.Append(level1ParagraphProperties24);
            listStyle108.Append(level2ParagraphProperties12);
            listStyle108.Append(level3ParagraphProperties12);
            listStyle108.Append(level4ParagraphProperties12);
            listStyle108.Append(level5ParagraphProperties12);
            listStyle108.Append(level6ParagraphProperties12);
            listStyle108.Append(level7ParagraphProperties12);
            listStyle108.Append(level8ParagraphProperties12);
            listStyle108.Append(level9ParagraphProperties12);

            A.Paragraph paragraph146 = new A.Paragraph();
            A.ParagraphProperties paragraphProperties89 = new A.ParagraphProperties() { Level = 0 };

            A.Run run93 = new A.Run();
            A.RunProperties runProperties118 = new A.RunProperties() { Language = "en-US" };
            A.Text text118 = new A.Text();
            text118.Text = "Click to edit Master text styles";

            run93.Append(runProperties118);
            run93.Append(text118);

            paragraph146.Append(paragraphProperties89);
            paragraph146.Append(run93);

            textBody106.Append(bodyProperties108);
            textBody106.Append(listStyle108);
            textBody106.Append(paragraph146);

            shape86.Append(nonVisualShapeProperties86);
            shape86.Append(shapeProperties113);
            shape86.Append(textBody106);

            Shape shape87 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties87 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties141 = new NonVisualDrawingProperties() { Id = (UInt32Value)5U, Name = "Date Placeholder 4" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList114 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension114 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement140 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{2A36FA2E-0A1F-F303-ABD3-C31591E0AD7C}\" />");

            // nonVisualDrawingPropertiesExtension114.Append(openXmlUnknownElement140);

            nonVisualDrawingPropertiesExtensionList114.Append(nonVisualDrawingPropertiesExtension114);

            nonVisualDrawingProperties141.Append(nonVisualDrawingPropertiesExtensionList114);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties87 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks70 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties87.Append(shapeLocks70);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties141 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape70 = new PlaceholderShape() { Type = PlaceholderValues.DateAndTime, Size = PlaceholderSizeValues.Half, Index = (UInt32Value)10U };

            applicationNonVisualDrawingProperties141.Append(placeholderShape70);

            nonVisualShapeProperties87.Append(nonVisualDrawingProperties141);
            nonVisualShapeProperties87.Append(nonVisualShapeDrawingProperties87);
            nonVisualShapeProperties87.Append(applicationNonVisualDrawingProperties141);
            ShapeProperties shapeProperties114 = new ShapeProperties();

            TextBody textBody107 = new TextBody();
            A.BodyProperties bodyProperties109 = new A.BodyProperties();
            A.ListStyle listStyle109 = new A.ListStyle();

            A.Paragraph paragraph147 = new A.Paragraph();

            A.Field field26 = new A.Field() { Id = "{B5CF329D-DBF6-4E15-9822-D0979B6928D0}", Type = "datetimeFigureOut" };

            A.RunProperties runProperties119 = new A.RunProperties() { Language = "en-US" };
            runProperties119.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text119 = new A.Text();
            text119.Text = "4/24/2024";

            field26.Append(runProperties119);
            field26.Append(text119);
            A.EndParagraphRunProperties endParagraphRunProperties75 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph147.Append(field26);
            paragraph147.Append(endParagraphRunProperties75);

            textBody107.Append(bodyProperties109);
            textBody107.Append(listStyle109);
            textBody107.Append(paragraph147);

            shape87.Append(nonVisualShapeProperties87);
            shape87.Append(shapeProperties114);
            shape87.Append(textBody107);

            Shape shape88 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties88 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties142 = new NonVisualDrawingProperties() { Id = (UInt32Value)6U, Name = "Footer Placeholder 5" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList115 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension115 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement141 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{707B6BCF-15C2-8415-EBA9-077C45B061FA}\" />");

            // nonVisualDrawingPropertiesExtension115.Append(openXmlUnknownElement141);

            nonVisualDrawingPropertiesExtensionList115.Append(nonVisualDrawingPropertiesExtension115);

            nonVisualDrawingProperties142.Append(nonVisualDrawingPropertiesExtensionList115);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties88 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks71 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties88.Append(shapeLocks71);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties142 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape71 = new PlaceholderShape() { Type = PlaceholderValues.Footer, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)11U };

            applicationNonVisualDrawingProperties142.Append(placeholderShape71);

            nonVisualShapeProperties88.Append(nonVisualDrawingProperties142);
            nonVisualShapeProperties88.Append(nonVisualShapeDrawingProperties88);
            nonVisualShapeProperties88.Append(applicationNonVisualDrawingProperties142);
            ShapeProperties shapeProperties115 = new ShapeProperties();

            TextBody textBody108 = new TextBody();
            A.BodyProperties bodyProperties110 = new A.BodyProperties();
            A.ListStyle listStyle110 = new A.ListStyle();

            A.Paragraph paragraph148 = new A.Paragraph();
            A.EndParagraphRunProperties endParagraphRunProperties76 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph148.Append(endParagraphRunProperties76);

            textBody108.Append(bodyProperties110);
            textBody108.Append(listStyle110);
            textBody108.Append(paragraph148);

            shape88.Append(nonVisualShapeProperties88);
            shape88.Append(shapeProperties115);
            shape88.Append(textBody108);

            Shape shape89 = new Shape();

            NonVisualShapeProperties nonVisualShapeProperties89 = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties143 = new NonVisualDrawingProperties() { Id = (UInt32Value)7U, Name = "Slide Number Placeholder 6" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList116 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension116 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            // // // // OpenXmlUnknownElement openXmlUnknownElement142 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{F4A9A490-7D25-56EF-9A69-7A32EE610111}\" />");

            // nonVisualDrawingPropertiesExtension116.Append(openXmlUnknownElement142);

            nonVisualDrawingPropertiesExtensionList116.Append(nonVisualDrawingPropertiesExtension116);

            nonVisualDrawingProperties143.Append(nonVisualDrawingPropertiesExtensionList116);

            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties89 = new NonVisualShapeDrawingProperties();
            A.ShapeLocks shapeLocks72 = new A.ShapeLocks() { NoGrouping = true };

            nonVisualShapeDrawingProperties89.Append(shapeLocks72);

            ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties143 = new ApplicationNonVisualDrawingProperties();
            PlaceholderShape placeholderShape72 = new PlaceholderShape() { Type = PlaceholderValues.SlideNumber, Size = PlaceholderSizeValues.Quarter, Index = (UInt32Value)12U };

            applicationNonVisualDrawingProperties143.Append(placeholderShape72);

            nonVisualShapeProperties89.Append(nonVisualDrawingProperties143);
            nonVisualShapeProperties89.Append(nonVisualShapeDrawingProperties89);
            nonVisualShapeProperties89.Append(applicationNonVisualDrawingProperties143);
            ShapeProperties shapeProperties116 = new ShapeProperties();

            TextBody textBody109 = new TextBody();
            A.BodyProperties bodyProperties111 = new A.BodyProperties();
            A.ListStyle listStyle111 = new A.ListStyle();

            A.Paragraph paragraph149 = new A.Paragraph();

            A.Field field27 = new A.Field() { Id = "{19E6023F-2333-4A80-AFF1-A7F70B52B230}", Type = "slidenum" };

            A.RunProperties runProperties120 = new A.RunProperties() { Language = "en-US" };
            runProperties120.SetAttribute(new OpenXmlAttribute("", "smtClean", "", "0"));
            A.Text text120 = new A.Text();
            text120.Text = "‹#›";

            field27.Append(runProperties120);
            field27.Append(text120);
            A.EndParagraphRunProperties endParagraphRunProperties77 = new A.EndParagraphRunProperties() { Language = "en-US" };

            paragraph149.Append(field27);
            paragraph149.Append(endParagraphRunProperties77);

            textBody109.Append(bodyProperties111);
            textBody109.Append(listStyle111);
            textBody109.Append(paragraph149);

            shape89.Append(nonVisualShapeProperties89);
            shape89.Append(shapeProperties116);
            shape89.Append(textBody109);

            shapeTree16.Append(nonVisualGroupShapeProperties28);
            shapeTree16.Append(groupShapeProperties28);
            shapeTree16.Append(shape84);
            shapeTree16.Append(shape85);
            shapeTree16.Append(shape86);
            shapeTree16.Append(shape87);
            shapeTree16.Append(shape88);
            shapeTree16.Append(shape89);

            CommonSlideDataExtensionList commonSlideDataExtensionList16 = new CommonSlideDataExtensionList();

            CommonSlideDataExtension commonSlideDataExtension16 = new CommonSlideDataExtension() { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" };

            P14.CreationId creationId16 = new P14.CreationId() { Val = (UInt32Value)2232514174U };
            creationId16.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

            commonSlideDataExtension16.Append(creationId16);

            commonSlideDataExtensionList16.Append(commonSlideDataExtension16);

            commonSlideData16.Append(shapeTree16);
            commonSlideData16.Append(commonSlideDataExtensionList16);

            ColorMapOverride colorMapOverride14 = new ColorMapOverride();
            A.MasterColorMapping masterColorMapping14 = new A.MasterColorMapping();

            colorMapOverride14.Append(masterColorMapping14);

            slideLayout12.Append(commonSlideData16);
            slideLayout12.Append(colorMapOverride14);

            slideLayoutPart12.SlideLayout = slideLayout12;
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

        // Generates content of imagePart7.
        private void GenerateImagePart7Content(ImagePart imagePart7)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart7Data);
            imagePart7.FeedData(data);
            data.Close();
        }

        // Generates content of imagePart8.
        private void GenerateImagePart8Content(ImagePart imagePart8)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart8Data);
            imagePart8.FeedData(data);
            data.Close();
        }

        // Generates content of imagePart9.
        private void GenerateImagePart9Content(ImagePart imagePart9)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart9Data);
            imagePart9.FeedData(data);
            data.Close();
        }

        // Generates content of imagePart10.
        private void GenerateImagePart10Content(ImagePart imagePart10)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart10Data);
            imagePart10.FeedData(data);
            data.Close();
        }

        // Generates content of viewPropertiesPart1.
        private void GenerateViewPropertiesPart1Content(ViewPropertiesPart viewPropertiesPart1)
        {
            ViewProperties viewProperties1 = new ViewProperties();
            viewProperties1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            viewProperties1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            viewProperties1.AddNamespaceDeclaration("p", "http://schemas.openxmlformats.org/presentationml/2006/main");

            NormalViewProperties normalViewProperties1 = new NormalViewProperties();
            RestoredLeft restoredLeft1 = new RestoredLeft() { Size = 17015, AutoAdjust = false };
            RestoredTop restoredTop1 = new RestoredTop() { Size = 94660 };

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
            Origin origin1 = new Origin() { X = 276L, Y = 114L };

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
            GridSpacing gridSpacing1 = new GridSpacing() { Cx = 76200L, Cy = 76200L };

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
            totalTime1.Text = "2";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "39";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office PowerPoint";
            Ap.PresentationFormat presentationFormat1 = new Ap.PresentationFormat();
            presentationFormat1.Text = "Widescreen";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "24";
            Ap.Slides slides1 = new Ap.Slides();
            slides1.Text = "1";
            Ap.Notes notes1 = new Ap.Notes();
            notes1.Text = "1";
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
            vTInt323.Text = "1";

            variant6.Append(vTInt323);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);
            vTVector1.Append(variant3);
            vTVector1.Append(variant4);
            vTVector1.Append(variant5);
            vTVector1.Append(variant6);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)8U };
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

            vTVector2.Append(vTLPSTR4);
            vTVector2.Append(vTLPSTR5);
            vTVector2.Append(vTLPSTR6);
            vTVector2.Append(vTLPSTR7);
            vTVector2.Append(vTLPSTR8);
            vTVector2.Append(vTLPSTR9);
            vTVector2.Append(vTLPSTR10);
            vTVector2.Append(vTLPSTR11);

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
            document.PackageProperties.Revision = "1";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2024-04-24T09:28:32Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2024-04-24T09:31:13Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Pankaj Kumar";
        }

        #region Binary Data
        private string thumbnailPart1Data = "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCACQAQADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9K/iJ8RtE+F/hyfWteuGgtI1YgIu5pGCltq+5xxnFaPg/xHF4w8JaJr0EMltBqljBfRwzY3xrLGrhWx3AbBqfXPD+l+JrE2WsabZ6rZlg5t76BJoyw6HawIyKreLPEEPgvwnqmstZzXcOm2klx9ktAvmSBFJ2ICQMnGBkgVTat5nNGNb20pSkuTSytrfrd318tjZpCa8t0/8AaQ8GX8OpXH2i7itLOaSMXH2SSRZVSBJTIuwEgHcyqGwWZDgHjMF/+0h4as2u9trqTx28NzJ5rWrIrPDDbymPn7pIuUXLY+YEehqTouer0hNee6z8btC07RF1S0gvNUt9lxK4hRYiiQRCWcnzWQFkDbSgO7cGUgbG2x/8L+8GPfGyj1C4lvRJFCbeOymaQSSOI0QgLwd5KH+6yspwRimQ2eiFqYzV5lN+0V4Khm2vqMgja0W8jZbWZmdMurkKE/gZdrdw24EDaa63wn400nxxp819o9y11bRTG3aQxsnzgKSBuAzjcPxyO1MybN0tTd1NLVGWpmTkSFqbuqNmpm73pmbkS7/ek3ioS9J5lMjmJt9HmVX8z3pPM96QuYs7zR5lVvM96XzO+aB85Y8ylElVfM5pyyUi1ItrJT1eqgkp6yUGqkWw1PVqqq9PWT1pGikWt1LuqBXp4akaJkwanA1DupwakaIlopoalzQMWiiigAqC5jjuI2hmjWaOQFWjdQQw7gg1KTUbf6yL6n+RoJM9vDOkO6O2j2bPGGCMbaPKhlCNjjjKgKfUACmyeG9JdpWOj2ZaQEOTbJlwcZB456D8hVS90/xNJdSG11azitzKWVZLUswTHC5z64557/SujoCxz8PhDSINMg0/+yYZ7SGdrqOO5QTYmZ2kaXL5Jcu7MWJzlic81K3h3TmuEnOk2vnozOkvkJuVmfzGIOOCX+Y+p561t0UxciZgSeGNMlZGfRrR2Ri6s1vGSrEsSRxwSXc/8Db1NWLXTY9Ph8m1s0toslvLhRUXJOScCteii5Hsk+pmNDMf+WTfp/jTDbz/APPFvzH+Na1FPmZPsI9zHNvcf88W/Mf40021z/zxb8x/jW1RRzE/V49zD+x3P/PFvzH+NN+x3P8Azxb8x/jW9RRzMX1aPdmAbO5/54N+Y/xpPsd1/wA8G/Mf410FFHMxfVYd2c/9juv+eDfp/jR9juv+eDfmP8a6CijmYfVYd2c6LK7/AOeDfmP8actndf8APBvzH+NdBRRzMr6vHuYQtbn/AJ4N+Y/xp4trj/ni35j/ABraoo5ivYR7mQtvP/zxYfl/jUixTd4m/Ssz+zfFYAC63YEbhndYksRk55DgZxjHHbv1rpqVyvZLuZ4jlH/LNv0p6rJ/zzb9Ku0UXK5EVAsn9xv0pwD/ANxv0qzXK3Gm+LvO22+s2P2dmOXltf3irgYxg7d2c8kYwAMHOQi7HRjd/cb9KdvKjJVgBUtMm/1L/wC6aBiq2adUEbfKPpUitQAE4qnfMSq/WrDNVO8bhfr/AEpmV9TMGtWDSToL+3LwNtmXzlzG3HDc8HkdfUVLa6hb3wc21zFcCMhXMUgbaSquAcdMqyt9GB71y+r/AAn8Ja9dTXF/o0VzLMWZy0jgZY7mwA2Bk5Jx1LMf4jnotL0ey0WKSKxto7WORg7LGMAkKqA/98qo/Ckal3cfWjcfWkooAXcfWjcfWkooAXcfWjcfWkooAXcfWjcfWkooAXcfWjcfWkooAXcfWjcfWkooAXcfWjcfWkooAXcfWjcfWkooAXcfWjcfWkpCAwIIyDQBi+FPHXh/xxDeS6HrNrqEdmzJcNG/EZVmQ5J90b8s9CK6QWUpGQQR9a89+G3wa8OfC06p/ZC3ki37bmS6uNwiHmNIFjwARhn4JJPA5rqYPDOh20ySx6Tbq6HKnkhec8D0zn86mPNb3i5ct/c2NhbOV1DKyspGQQ3BpfsM3sfxogu1toY4YoVSONQqqDwAOgp/9pN2QVRBUbchIbII61lx+J9Mkku0W+h3WrBJ8uAI2PABz7gj6gjtWrI5kYs3JNY0fhPSopbyWOzSOS8cSTsmQZGGcE/mfzNBnPn+yWG8QaerBWvYFJOADKvqV9fUEfUGr+7K5zkVkR+EdHiVUTT7dEUgqqxgBSGLAgY4O4k8dzmtU/Khx2FAo8/2jSjb5R9KkVqrRt8o+lTK1MpMRm61EuGuYQRnk/yNOkPNRRt/pkH1P/oJpmN/eRdmkt7fZ5rRx72CLvIG5j0A9T7UjTWySrE0kSys20ISMk4Jxj1wCfoDVbVtBsNdtxBfWyzxBxJtyV+YZAOQR6n86paT4I0PQ5Eew0+O2ZZfOBRm+/tK5PPPBIx05qTpNvy0/ur+VHlp/dX8qdRQA3y0/ur+VHlp/dX8qdRQA3y0/ur+VHlp/dX8qdRQA3y0/ur+VHlp/dX8qdRQA3y0/ur+VHlp/dX8qdRQA3y0/ur+VHlp/dX8qdRQA3y0/ur+VHlp/dX8qdRQA3y0/ur+VHlp/dX8qdRQA3y0/ur+VNk8qKNncKqKMsxwAAO9SU2SNZo2R1DowKsrDIIPagDhvA/xl8HfEWzvLrQdR+1QWb7J2mt3g2HeyciRVOCUb9PUV00fiDTJioSeFy2MBXU5z07965vwT8FvCfw+j1CPRdOaCO+bdMksrSrgMzBQGJwoLtx7+wrej8F6HC4ePS7SNwSQywIDk8ntUx5re8XLlv7uxch1K0uI1kiAkjboyYIP45p/2yD/AJ5n/vkUkGkW1rCkMKeVEgCpHGAqqB0AAHAp/wDZ8Xq351RBKvlNHvAXbjPSsFfHOgvJfILuHNi4S5LFVETEkAEk8cg/lXQLGqx7APlxjFYqeCdGjmuZVsow9y4klO0HewzgnI6gkn6knvQZy5/sjW8ZaIrKrXlurMcAGVMk7iuOv95WH1BHateXy5LOR1AIKEg49qyF8B6FHt26Zart+7iBBjndxx6kn6mtidBHZyqowqxkD8qBR5/tFCJ/lH0qwrVRhb5R9KsxtVEpjnbk1HCc3kH1P/oJp0neorf/AI/4Pqf/AEE0zL7SNSa4it9nmypHvYIu9gNzHoB6k+lI11AkyxNNGsrNtVCw3E4JwB64BP0Bqpq+g2GvW4gvrZZ4g4k25K/MM4OQR6n86p6T4I0PQ3R7HTo7Zll88MpbO/aUzyf7rEfjUHYblFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAU2SRYo2d2CooyWJwAPWnU2SNZo2R1DowKsrDIIPUUAcT4H+NHhH4jWd5daBqhu4LN9k7SwSQbDuZeRIqnGUbn6eorqV1uzbG2ZTnGMMOc9O/euU8F/BXwp8Po9Qj0SxktkvmDSh53kwAzMqruJ2gM7EY56egrah8B6DbSCSHTLaJgcgpEq989h61Mb294uXLf3NjUi1W3mjWSNvMjYZVlwQR7Gnf2hF6N+VR2+j29rCkUQaONBhVGOKk/s+P+835j/CqIJ1kVo94Py4zmueT4haLJJfIlzuaxdY7jCn92xJAz+II47giuhWNUj2AfLWWvhPSFknkWwgV7hxJKyxgGRh0LHHJHqaDOXP8AZKJ+IWiK0am8TMhwnP3jvKYHqdylceoxW7cSCSxldeVMZI/KqC+E9JXgWMIHXHlj39vc/nV+6UR2MyqMARsAPwoFHn+0Y0LfKKtxtVKH7q/SrUZ6VZlEmk6morf/AI/oPqf/AEE1PJ1NRQj/AE6D6n/0E0E/aRoXF3BahDPNHCHYIvmMF3MegGep4PFMXUrSSRUW6hZ2bYqiQElsFsAZ64Vjj0B9Ki1XQ7DXLcQX1qlxEHEm1v7w6Hj6mqen+C9C0uaOa20q2jnjl85JdmXV9hTcGPIO0kfQ1B2G1RRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFNdxGjM3CqMmnUUAeWfDP9ojw/wDFKPUn0+y1CxFi5Vlvo1BkHmvEGQIzZBKA/wDAh3yB11v8QdEu2VYdQtpWbGAsoJ5OP5/lUuifD/w34be9bStEstPa9Yvcm3hCeaSWJ3Y6jLMcdPmPrWt/ZdruDeQu4HOe+amN0veLly391aFWx8QW+pWkV1akTW8o3JIp4YetT/2kP+ef61N9hgHRP1NH2KH+5+pqiB6zK0Pmfw4zXIx/FHS5Z9UiSOYvp0qwzAqRlmJA28c8qenpXYqoVdoGB6VW/su1DMwhUFjkkcZoM5qb+FnKn4raMrwruc+c21CqMQTvZCMhcDlT14xg9DmurnkE2nyuvRo2P6Uf2dbf88h+Zp10oWzmAGB5bAflQKKmviZhw/dX6Vaj7VXhX5R9KtRrVmUSZzVa4HzRn3/pVuRetVJ8hk9M/wBKRfLqikNZ09pJ0F9bF4G2zL5y5jbjhueDyOvqKktdQtb5pBbXMNwY9u8RSBtu5Qy5x0ypBHqCDXM6v8J/CWvXU1xf6NFcyzFmctI4GWO5sANgZOScdSzH+I53dJ8O6bock72FpHatOEEnl5wwRdq8ewqTc0aKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoopCAwIIyDQBieFPHHh7xxDeS6HrFrqEdmzJcNG/EZVmQ5JHqjflnoRXSCxlIyNpH1rz74bfBrw58LTqn9kLeSLftuZLq43CIeY0gWPABGGfgkk8Dmupg8M6HbTJLHpNurocqeSF5zwPTOfzqY81veLly39zY147N5UV0ZHRhkMrZBpfsEvoD+NJa3CWdukEMKxxIMKoJ4FS/2k3ZBVEFRlKsQRgisSbxhpUGrXOmtc4vLeLz5Y9jfKmAc5xgnBBwDmtySQyMWbkmqUmk2UlxJcPaQtPInlvIYxuZf7pPce1BEub7JlyeOtDiVmfUIlVXaNic4VgqsQeOCAynn1raEy3Fr5iHKsuRkY7VD/AGVZ8f6NFx0+QVP5axwlFUKoXAAHAoJjz/aHQr8o+lWo1qOFflH0qwq1ZKRIy1Gqj7VDx3P8jVhlqMD/AEiL6n+RqTWxZ2r/AHR+VG1f7o/Ks3XNBXXLeOI3l3ZFJVmElrIAxIzgfMCMc56dQD2rP0/wWbG6hnfXdXvWin88LdTRsD8hTacIMrhj755zmkUdFtX+6Pyo2r/dH5U6igBu1f7o/Kjav90flTqKAG7V/uj8qNq/3R+VOooAbtX+6Pyo2r/dH5U6igBu1f7o/Kjav90flTqKAG7V/uj8qNq/3R+VOooAbtX+6Pyo2r/dH5U6igBu1f7o/Kjav90flTqKAG7V/uj8qNq/3R+VOpDyCKAK0N/ZXLSLDcW8rRkq4R1JUgkEHHQ5BH1BqXzIf70f5ivGvhR+zj/wq/UfEMya/JfW2qM+yJYBEybpWkLO2Tvb5sZwB14+bjvofA0kE0TrrWpERnIRrgMp9AQV5A//AF5qYttaqxckk7Rdzp/Mh/vR/mKPMh/vR/mKzLHQ2sLSOBZmlCcb5W3Mec8nHNWP7Ok/vLVEF7ahGcLj1qouqWLySIs8bPGcOq8lTjODViKDy4PLJzkc1yUfwv0yGfU5UklDajKss4Y7hlSSMA8AZYn8aDOTmvhR0/261/vr+RqWZVa3kIAIKHBH0rkf+FVaT5kT+ZcExfdzPIerlznLfN8zE85/LiuuMYhtGReioR+lAoub+JFWNflH0qVVpI1+UfSpVWmUkSMtRsp3Kw4KnNTVW1Cy+3WrQ+dNb7ip8yB9jjDA8H8MH1BIpFjzJJ/s/l/9emmaT/Z/I/41jL4TYRoDrOqsyjG77QOeOeMf5xT08MrFOrjUNQKjnyjcHacnJ7Z/WmSabXEo/ufkf8aYbyYf3P8Avk/41l/8IwVQr/aupEcHPngEEf8AAe/emx+GTGD/AMTTUZMvvO6YdfThenTj2+tBOpptfT+kf/fJ/wAaY2pXA7R/98n/ABrI/wCETxIT/ampsuQQjXHAx+GfT8vrlbjw0Z3jb+0dQj8vOFjnwDkknPHPXHPQAYpk6mmdUufSP/vk/wCNNOrXX92L/vk/41zV94MvbicPDr9/bxhUXZuLZ245JyOTjnjnJpJPCN+4/wCQ1cKd24FQ3X6b8dqNCby7nRtrN0P4Yv8Avk/40065d/3Yf++T/jXKTeCdSmVQfEV8vyKjFQRu2qBk/NwTjJIxyT610sdqyRIrEyMoALEYyfWnoReXckOu3n92H/vk/wCNJ/b95/dh/wC+D/jUZtz6Gmm3PpT0J5p9yX/hILz+7D/3yf8AGk/4SC8/uQ/98H/4qovs5/u037Of7po0J5p9yf8A4SC8/uQ/98H/AOKo/wCEgvP7kP8A3wf/AIqoPs5/un8qX7Of7tGgc0+5P/wkF5/ch/74b/GlGvXn92H/AL5P+NQfZz6U4W59P0o0HzT7k41y7/uw/wDfJ/xp41m7P8MP/fJ/xqFbc+hp6wH0/SjQtOXcmXVro/wxf98n/Gnrqdye0f8A3yf8aiWE+lSLCfSkWnIkXULg9o/++T/jT1vpz/zz/wC+T/jTFhPpUgjpF6jhdTf7H/fJ/wAaetxL/s/kf8aasftTwvtSK1F86T/Z/I/40rM7qVJXBGOn/wBelVacFpFDVXAAqRVpQtOoGf/Z";

        private string imagePart1Data = "iVBORw0KGgoAAAANSUhEUgAAAFAAAABQCAYAAACOEfKtAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAOw4AADsOAcy2oYMAAEAnSURBVHhe5b0HdFvnlbWtmdixLKuL6l2i2MTeO9gJ9ssC9t5Bgr1XsPfeO8VOiqTYm0SREimqW7RkSZYt23SPE9tRJpkkkzj2/s+FHNnJzKwvM8l8yXz/u9a7LkAAJvBwn7P3uYDgTX/PhSfNYt9+1sn99tOu6G8+76/9w08Hl//w06GNb39xYePrzwc7v3ynZzt7v3/7t3+z/Pbbbzd+97vfsXuZrjfSjv7tb3/LBbBT9B/7/8NCJ2fzL0Ytub9bcCr9w73Y9W826vDNp9349if9+Oan52kP45ufjQL/MsUeR34AkCFQIHj/btNtIJDrtCvpsjXdbbPol/2/tG7kKhhcTpNtXcuRffawXAVf9hrgm2uB+ObdSnz7WQ+++XwQ3/50FN9+cQHffjkG/GoW33wx/u8Afv311/j973//Yv85yO9g/uo3v/lNJx1NRL/8f+sSCjkvtQVK+naFSDw9FyyJrmApjEfJ4Ga2In7SpoevL/HwzcMMfPNxO7792TDBG8e3X03QngL+dYEUOPEC4C9/+UwE8De/+bUI0u9o/54F92cQfwjyu/30X3/7W1966EuiJ/W/ZZ3eu4Ure2TbutrJ7dCX3AlrRTG4aOxDJnMUHSFnsJKjjF+ct8Q3N0KAjxpFyvvmy0mCR/vLCSrhCSrtXgJYLQLYt7jCjF6/hYHLVzC9dhPL99bx+P338a+//rVIlX8OkFXns2e/wOtvvInh6Usoazy3Xt7UyRU9uX/kdXTPq6on9m5dPr53G47t3oIjOzfj6K7NOExHcbFXoXZiKwwkt8FKcTdKPMSx0cPF727FAO9Xi0Big45PCoHH6fjtWtjIl7MeIoBt84tM/dwikjv6Edfai4T2PjRPXcSnP/sC3/zhD39S0n+g648fPYIwKw8BggQwrr6wcfaFvVcorN0Clx39w1VFT/Yfae3Y8cqJPdteGdi/8zWIbX8VB2jv3boZO15+Cbs2v4w9W36M3a/8CGKvviSCyUKVPPAaPHX3oSZQHIvF2nin2xSf9BrhYZUKbmZJYkpwZKTHY7cI4ODyCtN3dQ0l5y8gr38EmT3n0TK7iM+++BLffvPNi9747bfEn+AFeHrA0NAUKuo6kDmrCBUNPahqcqCspgc1TUNo6pkMaHDMT4ie/N977dvxmsn+nVuf7d+1Dbu3b8GOLZux7RUC9+qPab+CXezlzT/GTjruJIi7tryEPa+yUF/Gwe2vQGLva1A59ipsFbYgzmgbsjlbkaHxGhKUt4xEqD8H2L24zHQtr6J+cgaNUzMoGZ1A28VlfPblVy8AEjv87Gc/RWIEH3JnTkNKXBLHDh+jfRSnTopDQlyKjmcgLSUHaWl5qKtqPlNV1fz7Go3Y1i2Be3du/Xrfzm0Q2/Eadu/YSgrcir3bXiNQBJDUt+NlgsaqcOurEHttM3ayP/vxj7CbIO7b9goO79iMU3tew9Gdr0Jy7yuwltyCeI1tSFBjAW4SAayfmWea5i6jfmwSrdOzqBibRselKy8AsmXLQjzfcw6aslIQP3oYRw8cgsSxI1A6c4KuH8HJI0dx7NARnDl1BuIEVEVWDjpKyl9rKqmEil7M/+X1EoGrPLB7Ow7s3oGje3fjxKF9OH5oP04c3ItjYrtxgGDupVLezYIkNe585WW6TIp87RXsJHg76Wc7NlNZb/0xlfwrOEIg9732Mk6LbYa19FYEixT4HGDb/ApTf3EFBUPjKB2dRM75STTNXyWAPxcB/IZq90MyFQ8bLmRPnYQ4gZM8cRRmytJw01eCztkzOH3kCAE8BPHjJyBN91GSloKphhrMNdVhoKrayOFw/u84tYzM3q1SJw7OnTl2gP7CtI8fhviJQzhz/BBOHdmP4wf34IjYThwiuIf3bMOBXawqqSdSebN7PymV3WKkyB1U1qwaxaicWSUeos2WtThBVDz0yghX/DnAwZWbTOvyDST3jNIeQ+rABOrnV/HZV88Bfv2Hb9BcUwklidNUvqegKE6bjgILLTSEOSKAqwMFuk3i+HFInToFBUlJqBBAjooyHAy0YaunDT0lhUsKCgr/s1ONmszxA3InD7159uQhyNAmkJA8fhBnjh7AqYNiOMqC27VDBO/o3p04tn+nCCZrLgdos0cW3pHdzwHv27FFVNZipMqDdNsxsa102xZROR/a/srI7t3PAfYtrzGdV9aQTwrM7h8jBU6gffEaAXwGIohPPvkElgZ6kDl1QgROTUocBvISqA2xwVp1POqjvGChoQxVgiZL8OQlzkBZShKacjLQV5CFva4qjNWUcfb0qbe0ZGWlRC/2b704pLyzx/e+KXlEDKcO7cbJA7sJ2l6cObKPesw+nDlMl6mEJQ7thSR7me4nTvc7fWAPTuzbRXCp3Hc8L+m9r5Eit20RlTPbG/duZdVJ0Aji8b3bQVEIR3ZteQGwjUykk1y4ZnoB1VMLqJyYxbmlNXz8sy9Fgbq9qYGewyEq00NQlTwNLenTcNKSxUxuIB6dy8NKQxYCLQ2gLX0GZ08ex5H9+3D68CGoiJ+EBsE2VpIWQTx74hg95wNPOSoqYqIX/bdaHOp5Zw7tnpM6KkagdhOk3Th7TAzSR/fiNJXsyf27cHo/C4ug0uVT+3ZC/MAuSNCWotul6P4SB+k2gnN01xbqeVuwh5S3ffNL5Ng/wg7qibu3EszXfoyDpEo2Rx7Z8z3A4pEZpmZuCVWT87QXkDs0gTq6/s4nP8Hn5LxuDraiXnxEbBc0pU5CR/o4Iiw1sN6cgvdGavHWUB0KA3jwN9WGjZos5I4dhNTh/VA5dQzepjpw5GggnDEjmCdF/VxLTmpZKBT+7Xqi4skDlXInD0BJ/BDkTuyHmsRBKJ8+AMlDe3DqwE6CtwuSBEqGgMqRGuXoKHtkDxSO7obycTEon9gDpePsbXvoye+G9MGdOEWK3LvtO4jkzDtJmTspO7JGc4AgsgoU/w6gsG+SEQ7PIeXceSR1DIJf14Oi0QW8TQCfvvcuDLQ0qM+Sk+/dBXOlM7BRkUC9gIc754rx9kwP3psbwGKNEAuVabhULcRIbjTak4OQRVCTfXkId2EgDPWDl6UJZI8fgSL1y/zc3FbRi/9rl52eYqCtnjLs9FVhoSkPEzVpmGvIwFRVGrqy4tCWOQGO7GkYyZ2GmcIpmCmehqnCSRjLHQdH+hD0pQ5Cj4CzR47U8+v69HPNMwdIpTuohCnaUClvF20K39859s4tL5MCn+dAQX03E90+jLC6LoRWn4N/ZScKhufx5JPP8eTpuzDV1yYFb6EXvx+pLhxk+9mgNiEIXXnJmK7Jxq2uMjwarsUHs+fwwVQ7KbIaq005mChNxGxFMlbo8lpnBZb621CWHA2pkydRXFqFr379+78u4giDHExSfO2+FobwkBLgAIGLBUIcTRDsaEZHMwQwRgi2N0aYgwnCHYwQxugjxFobPqbKcNGXg6O2NGw1xGGjegqWSidgrXwClrTNFY7Bgo66Mkeoj+4WZchdVNasAlmA23/8Eja//M8j9BREAIPzy5i87gF0zS+icWwOwUU1SGvpxcMPPsGHn30OxtxIBFD77HF0JrphsTYJTydb8Onl8/hoph2P+8vwoKcYG5Ot2Jhqxe2OfFypSsSNxgxcacjEfGUKVhtz8Wi6D29dW4Shnh7KSsrx1a9++/WXv/pvntUZEIac6MsMfNYjDEZ3diiaUv1RFu2JQoEbCmhnh7oiM5gHYZAjsoOdaDsiK9Ae6fTXT/QwR7STIcJsdRFgoQpPY3m46EnDVVcKzjqS1ODFYad2Clzlk1A5cwgnyGj279kOMZpmWIjPe+NLI7u/A5hQWcn0X1zAIyrXa3fvIjgpA4KsQjzc+AjvfvARzPS0sO3VzdClKrhUEY3HXXl4q7sA46UpaEsXYCQrGLcb0/HeZBM+uTyAu135uFwqQFdKAEJtDBHlYIqGKB9cbSzAh3dWkSQIR0tNIz5/9mt89S+/efbs2W//62PffJlgYLo4HJPFAowVhWOoIAI92WHozAhBc1owGlOCUJcYgMoYb5RHedDRE5UEuCTMBTmBDDJ8rZHoaooIB32E22gjiKtGTVwZfiaK8OTIgqctQQ39NPTPHoP08X04TE69b+fW5wApF/4QIF+Yw5S1tWN+YQrjwz0ICItATEYenpACr12/AVWKI05qZ8HVkMflqljcqoxEO59BMVVNgKkmin2ssFoRi/dHy0iVfXhzsBxrVTEYy6Qyj/XEcHogFsqScX+0E5+t38DiQA9a8gvx7sef4+e//A2+/Jd/HRBB+UvXxYpw1dniMEwXhWGGAE7khWEkLwJD+REYpN2XG4HenEicEwrQkhqCxqRA2gFooF1NmauY74KCEFKmvx1SPLiI5RlBYKeLUCsNEcgAAulpIAsnbUlq+qdE5nT6sBjlxK2iqYVV4Guv/NOLEubnFDMxBWVIEmYjISEBAYGhyCkqw3vUA/v7B0TZr8qOgxArQyrNWLzRFI8bBOgOW6IViVitTsFbvYX4cKQUH04148lYM+6fy8fj3iK8M1yNt89XY2O2G5+uzuInb9zExp01ZAT54c69+/jyl7/FL375a3zxi1//5WdxFkrClqdJdbPFEbRjCCD9tQricKEkEcPFCRgoiEF/fix6yM3OZUWgIyMc7Rl8tKaHoSk5GNWkylKBK4pCedTQfwhRB3xrTYRYqMPbSB7OVNLWamcoehyF5NH9otDNTijbX/5nvPbSP71QYH5bNyNs7kJAcjZcAwXw9AlCflE5Pv78C1SUFFOcOoRxctRhKtfbzZn44Hw5PppsxsfU/z6e6cAn8934eLYVH47X4eOxOmyQM791oRnvUsT5eK4Hny4O49OlMXy2OofP7qxg48ZVZPq6Y2CoHU8/e4LPvvgI73/63rIIzv9p3WpN416qiMFsaSypjzZJe7I8BZMV6ZiszsJoVSaGSlMxVJKCgaIk9BXEozuHBUllI4xAC0FsSAlBVbwfSsPdkBfogAwfayS7cxHjaAgB9UUBlTRbzh4GcmDIaPTlT1E8OojjVMaijLiFnPnlH5RwupBJr6xBdk09ojKzwfPyQ1JGFj78yedIjo3CYZpsxvmeuFOZiXvnivDJeDM+X+jDZwu9+GyOlDXVho9JeR+MVOFz6oOfkbm8PdWFh12F9LN6uu0cPp7uxofzw/jw6hw+ICMpEQSiuDgJVx/O48qDaVy6N47R6/+Hk7LLy8KX7nTkri9Wp2KuKhXTlJumKTdN1eZgqi4f47RHa/IwUpGDkbIsDJdlYIBA9uTF4Vx2NNqyqKTTwlCfFIzKOD96Eu7IC3JAJhlLiqcFEpxNEEVOHUHlHGCuCh/qh076sjCg6KN85ogolLMjHTve7dpCo9x3AJ39fZiCmnLMLM2hobsD9p5eCEtIwYN3nsLHlYd9lAELnKxwKz8O65T9Pppsx+cXBwngEMHpxPsD5dgYrMRtctxHFGnunyvHamsprlYk4W59Mt1ehfeHGvB0rAPvEMQPrl9GdlwUCrLTcfvpW1i6fx2rD29g7dGN9eXl5f88YL810eh7u7MIS015WGjMx2xDIabrCzFRR45WW4ALNfkYqc7DcHk2zpcK6ShEf0k6uglgBwuQFNiUGipSX1mUJ4qpjPOCHUTOnOZtKTIVVoURdnoIpknB31wF7sZKMKXgq3z6MMQJ4FGxHQSEPe31vQLtPD2ZtGL6vXNjqO89B+dgPsKzitE/fxWmBgY0Z2+Hr6EOntTlYp2e/3tjpDK2NGd78O5AJR53F+HpQDUuF8VjqTgeg/G+6I71xkJhFK7XxONBUzoetRfSY0tx73wL3rs8ibhgf+SkpWD9nQ9x7+kneOPdT3Hz8XtYf3vDVwTrP1r3L9Q/vd5VgisdpVjsqMZ8Rz1mW6sw2ViC8foiAkgQawtxvjKXlEfqK81EX3E6OnNi0ZopQFMan5w5GFUxviiL9EIRC5DMJItcOY11ZbYXuhgjwl4fodZaCCKIftQPbbRloCp+ECdpBDy8+/kZnF3bNn8P0D+YcY9KQEhGNoKzCiAoqMDCvcfoGZuHipwsQd9Njz+BycwYPOoswcPeUmyM1uP9wSo87MjB2yMN+IBK+Z3xFrqtHLebMnG7NRvrXcW4RZevkGEuFYRjoTgON3qqMFiSDVWakTOiI3Dn/lu4//QjvPFkA3fefBtvPH73qQjWn6+3ploN3qCEvtZViitdVVjsqieAdZhurnwOr5YtX1aB+RiiEu4lcH2kvp6iVJH6mlKCURcfgDoq36pYH5RStMkPc0YuKTA3yP5FGce7myHSXg98Ky0RRDZ4uxooQoeMRHz/Tirhzdi/41UazQjgd6OctU8wYx8YDruAcDhHJaFxdAofffVLtPcMQ/bMaRzb+/ysj7ehNm7XZ+Nuaxbe7MjG/aZU3K5LpMmjEu9caMTbw41UwtW431qOO1RdNxoKME89frogjNQYRoYZjfsXWuFsaU4z/36khAXgwRsPaFz8EE83PsS779N+bwNPn24YiKD9cD0cr299neR+jQAudZZjvq0cU6zyRKVLfa+ceh6BO1+Zg35SXndBMrrzk2gnknlEopHctz7en+JMEKrjfFEe7UU90BX5IU6kQgdk+dsi3YvKmAXoaEABVvvF9mdVqCEF1VP7Ic6eSNj1GvaxCvwjQDcfhhcUBtcQATwEMegbG8b7H3+ACgq7UseP4eieHQR+Kzn4FlK3Ke510FTRmYMHrWm4Xh1LZRuLFRrrbmSmYTkxCdOCCIxTluwLDkVdkAcmckNxnWBfa8zEal8dOGoqkD91DPFBPrh753U8fvtdPHznPTx88i7urr+J2+vrfzonb3R2br4/VPHsOkl6qTUfC9QDp8k4xiqFGK0gcFSuQ0VpGC7NwiCZx0CZkJSXhu7cRPQWEMS8eHQIo9BMEFvIgRtJjbXUZ2pivVAaQRBDCWCAjShcJ30XacLJSEJIgWwvDCY1uhsowE5dAnqSRyBzaDd71uYFQDsfH8aVHwb3cAIYHoHOwXN48/F95OXmU/w5iCNU9ux5RjIeqEmfxHRJPN7pKxbth905uFWbhtvFWXiQX4D13EJcS0zGGrn5lfIi3GorwZ3OfKzUp+EuGVBlsgCK0hJQkTiFKG8XXLl0Gddu3MXqzbtYuX4Hl6+uYeHq6jMyk+8/AfFwtJx7h37ZldY8XGzIxmxtFiaqhLhAJnGBgI2QYbB7rCIXF6iE2TJm+19vYQr6C1MxSDD7chNwLkNAmTCMxic+GhIDRQDLIl1RwAL0t0YmxZlMb4JIZiJg9BBC8NhQHWihCS9jmp/1zoKhXKh75jDO7N/5AqCnszPD9/NDJLsDAtDYWI+FuWlEB/tC+gjlRwJ4kMbAQ3QU27YFcc5meHCuAO/SxPHeaCVlwRZsTDTj7YE6POiowG2KY9fK03CtOQ+3u0ux1l6A5bo0zNcKYW+sC4njR0XnCQUeDrg8O4vVlRtYvXYLa2u3cePGHdy8eRvXr1//PtKsny8svX4uF8utObhEQ/UsudlkXR5Fl0LKfgWYoD1eSTGGPdYWE8ACDJXniIyEdWN2DxWloC+L8iABbCeAjclBz3thhBtNJs4UZ+xFk0m2P4NUDwvqg/oIptL1IRcO4KrD00gJPC1p2KuLw0zhBBSP739xNibKzZnJDvBBgT+1hjA+zje3YKK/B372XEgd2UslzMLbgWOUI8W2vSoC2kBj5WOaNp4OV+GD6RZ8QBDfHac+eL4ej3rK8AZV292eUtykfa2zGAvV6Qi2N8PJwwchfeIodM9KINLDEUvzcwSL4F2/jVu37uL119eprO/hxvVbpSJ47Lo3ULi+1skCzMWl5nzMU3OdbSqmCFNCpUzH2iJMVuVjgiCOVZELV+SJNgtxhN3FmRguSsdAHpVzJpUYTSYtFGdqqCc+74VuKKTxLjfIEWk+NtQHn4fqMCphH2NF+Jmpws1QAQ5aUrAjBVoqn4b6mUMj4t8BHBubZBbnFzAxeB7TwyNYvnQJsxPj8HGwEp2oZZ374K7tZCa7Re+77Nj8Y5q1pbFUGoH7VJ5PKEBvEMCPZrto92KDjIIF+WCkEXcGazBfJ4SXhR52k3pPsmfVabLRk5NEjI8zri0t4SaBW1m5jqu0WfXdun2XLt9YF8F7Mlkmdm+oEGudebjanofF5lzM19NuJIdqKsFcXTFmKL5M11AeJJAiBVaSqVQVYJigjlUXUqnnUp8kReanUAlHoS0jkqaRcJRE+yEz1FV0Kize2x6RrjYIZkzgb0Pqs6USJgPxNVOBp6G86PSXvaYUbFXEYaUsDl2pIy8Art28ydx7801cvXETa3fu4ibNqLPz83C3MSPn3oUje7bjMLnwif1ilAmpjGkcdNaRwnyBD65VRuFOcwaeUMJ4nx3fzjfgUXcZ7lL5rlK7Wm7KoRbjDPo12P3aqzh95CBOEUR9eSmEOnIxe74fV5av4CL1wkuXr1Ipr+HO7dexcu0mqA/upP5Xwl3vL8L1c3m40kYKbGJ3AS62lOAS2f1CUznBK8I0gZxqKCNnLscYQbxQXUQgCzBGt7F9sT8/HW1pMaiKC0FemDdivezhZWkExkADVnpqsOWow4FihoOhJnjGGnAxUYeHqQa8WfVxZOFIWdBOXRI2quwZ5TMwkjs+oi7+HODIyCAzOzeFC+OjmJiawNTcLHq7OuFgpC166+AoATy+j303kEZBArhvx1aKNLK4XOxHMSYKN6oiacfiRmUCrhRG42KOAJNpIRhND8ZCVTKC6I/JAtxLE43EscM0Ee0XKdDXUh+9tcRgZgZXSIlXV69h+eoKlfN1KuHrWFle5m56PFoSvd5fjOs0F16hZrpI8C63FONiWxnmCeBsYxn1QHYaKSGA5ZggiON1pRglgEOlOTSN5NBcnInWlEiaPgKQ4uNEKjOHlZYytKROQU38uOh9WWu67mlhhCBHS4TyLBFgb0Jlows3Y1U468nBkfqfnTopUFUC9mqS4CqeIoDizwGODTJzCzM4PzKEodFhXJiYQHtzAxiOJk5T+GZ73klS3z5y4kNkJvsJYJSNKkUYPtZbEnCvMQ43KwW4VhqOywXBmMnwxmiyB6Zz+VhpzoKZqjT+iQCyj5M6fkSkZG1ZCXiY66CD+vvchUEyjxXcv/86rq0ukDNfwLXFUVy/OBi96c3Rstr1gTLc7C3D1Y5iEbxZyn9TDQSoMgvn8lNxLieJ3DaTIg05MFu+1P9GSX2s6nqyaYxLj0JVpD+yglxFZeprZQQLdXnRW4sW6gpwNddHjLcjcghwTlQgMkI8EOvjgBCeBdzNNeGgJwt7LVIfKY9RkyIjkaIyPv0CYN/ICDM+t4D+4WEMjJzHMPW/mtJCmNHvOClG5UvQThzYK3Liw5QJT+/bjjI/IzxoS8TjLiHut6fiVk0UuW8kbtbFYrUiEovFfCxXxeNidRIUTx3Cj/5pEw7u3kmx6JBIzWqSp+BkoIoWyrkz/S3Udydx53VqISszuDTRgqWpFlyebKrd9Gi8fPneUAUZCIEpTUYvTRZ1wnhkC/wR7mxNauIi0deNjCAIHVSifblJGKA9mJeK7owYNMYFoCrKG3mBLkimsg134sLdSBOMljzcTNUQ7mKKdOqDBXGBKEoIQZbAG4l+joj2sAXfxRI+thy4EUQnfXnYqEmAoUDtRGZio3ZmhKv+HcDJeWbs8ir6Lkyg/8IFjExMIyc5AdqU+Y6zAHdvF5XvcVLOCZpK3LVOYzTJgeCl4d2hArzVLcS9pgQRwBuVEVgtC6dwTQArYzGez8cpagMv/+ifaYTcAXHqf8cIoOyxQ+CqSKE8NRrTAz2YOD+IC1NTGBzqxvhQI24uDeLmlaFLmx5N1GysdOSjKysCBZHeiCbncbUwhbGKArhqivDiGiHCxQ7CQE8CFYSWBD7aySC6Mmn+TQxDRbgn8vwdkO5pgxhnC3hRb7PXPAsXGs98uBoIYQwhcLFGqKs9fO2t4WxuAEfqg46GaqQ+bbhxteBpqUO/R4ugUx/UkCADoDijKfEC4OLyIrN25ybml5cwt3wVY9STgt2c6EXuw9Hd7Jvx23HygBhO7d8NRuU4uiNNsVbigyc96WQaBXg6kIuH5zLxRmMSbtdEk7FE4GqZACs18RjM9CcFv4ZXfvQj7N2xDeIH9+EEARQ/KEbj5TEk+jhipqcFUz2tGOo+h5bWFowOtuPutSm8fm1qY9OdwdKN/vxo5IR6INrTHgEOltTk9WGnS9nMnEMv3gZJvo7IJIUVkTlUhHuhISYAnamRaI4PRjnfHRme1oiyN6T5Vg/OuvJwNlSBq6EqnHSVSYUc2BnpQUtZEZKnT0OSQirbE3WlT8OU/sJO1AM9LCnOWGrDUeesKMq4G7AlLfUC4PJkPXN36RxWZttwbaEbnZRRuRoKOEMR5ij1PxachvghCCwUMJZij/VafzztiMKnI7n46UQxPh8vwccXivFRP+XCrkw8pJJ+gxT5elMSelO9cEJsGza/9M/Yt30rJA7tg8ThA6LxUPXMEXiZalLOpTFvuB2r40O4SD14aX4ad6gn3ru9urHpakfeF+dyY1CRJEBZciQqM+JQSoZQEBNIZeyBRC87CBxMafQyQhyPS5OEHRpiAyn3JVNsiUdbKh8lVKIxDqQ0ewN4EBAXIzIGIzU46SnD1ZgDJ0suNJTkCaIy7EwM4W5hCAeCa039ztGAgjTDob6pBWd9WfBIfeyZanud7xU42hR9cLghljfWHMuMtSUwKcHWjJbEEeb0vj3MkT1b6biDMZM9wnRGWTBrZT7MzUp/5kGzgHnSnc486RMyT7uFzKPuHOZRZzqz3pzC3KyPZdYqBcxymYCpETiIHv/qjzYxe3ZsZY4f3MNIH9nHHNzxKiN7/CBjrS7L1KQImJm2GmZxcoyZ7Otm5icply4uMhcvXrTcdKk5GyM12RikqaOPsl4XGUQLGUdxLJ+Ux0MogQvg6tHoZYpIBy4S3GxQQwqcosfM1RehPzsatZFeSOKZIJLCsTcbT4xowjDTpFnXEKn+HsiPj0JRcjxKhCkoSY+HUEAjmYcVqZQAGyrBnxTobaoCJwq/DpoScNWTZkGOeHwH8B96XWzJwVxbKQZq8lGXGUfKC0MaPxQR/v4IcHaEn5UpQmyNEU3umuzvSj2BJgq+N8ZI1hfrKAvmxqM+whMpTmbgW2jBnwCGkxJTfeyQ6eeEnGAPlMSGoCIlmkwkHJl8H8RRXwm04YjCdIy7CWLdTBBGs7Gf+fNM6EElzNOXeaHAf+g135T/xUJnKWrIGBK8XRDs6gADVRVIS8lCVkaezEQJHhSAA7gcZAS5oCwuCOUxQRjOicNsYSoGkgWopf6Y7mSKMDMNBBmT89rqI92HQQZFlQxvJ6S42yHB1ZrUa4tET5qFyYXT/OwJsjWSvCwQ42KCcEYfgZYacNM/K3r/mKcr9UKBV1OVD64mq/CWY9WZ5QRZZjlWmlmJVWAW2S2QZmb4EsxMsCwzI7oszYwFSzCT7PZnL7Nblpmk+0zyT4uus7eNie53mhkR3S7BjNB9B31kmUHRUZoZoT3oT9fpsSN8FWYkSpcZi9Vl1ur5zKdP1piNp3eZG1dmLTdN1mZsjNemIcXDhkYrLvgODvBzdIK7nR18He0R5+tBL9KJXrwdCslEmlLDUJ8QjLHsKExnx2IwMRxNoZ7IdbFApIUmwkxVEWOnhxxfun8AD5UCHzRQfGlMCkNzagSa0sLQSHNyeYQHstgTC95WBNcUcS7GCLbSgLu+DNwIoNMPeuC6UJd5t5CD9Uwt3E1XwZ0UBdxNU8GtJEWsJchhNUYGVyPlsRIrj+UoWSwKzuKyQBqL4fJ0mb0ug8uR9LMIdstiKVIGl2gvhNG4F34W83T/ef5ZzITIYor2ZJAspoPlMM1XwJRACbPRGphN0MOlVMqWnTH49WcP8OVnb+PhvasbmyaaczcGShNQGOREL1CAodoydFdXoLOsGPUZSSgJ90MBBd/CUC9UUoxpSAxFW3wAZvNiMJMTj2EK0W1Rfsj3sEY8mUE6gcj3tkCxry1KyHCK/RzQlBiM/uJkkZv1F8Shhx7blslHdbwv8kMdkUQlHM3jINqJfW9XA34m8nDVl34BcCVFi3kjUw/XEtVxNUEFV+KUsByniKUYOVyMlCUI0pgLlcZ8mAxmQp9DmA2RouNZTAbLYCxEAhN0edhXBj3ukuhwk8B5Pwlc8JfAsLc0HWUwGiiD876SGPKVwpAXe/kshgJoBylhlK+KkUhNTMQZ4WZDJH729g18+OR13F6d39g0XZe+PEAxZqwsHovtNOM2FKCjIA21qdT4w30hpBIUkplURQehnYJze3IYhtLDcLEkkXYqzZOx6IoNRgmbBb24ovN+hYHWKPK3QaG3NUpJiTUR7pQzozBSlobRshQMFSegJycKrTSPsu8hZ9F9Y50pMtnrgm+njUALFWobMi9KeI0FmGGA1QRNgqeClThVXI0hiNFKpCQ5Ebj5UFJRmDSpiCAGS2GOgI0TlMmgsxj0kUS2yRG4yOyA5YmdMD26C64Su5DHOYQ+lzM4781uCQx6EUBPGQz5EEA/OQyxO0ABw6EquBCpjel4I9xpiMKX79wigPdwa2V+edNae07jpdYczDdlkrNmoDsvEfXJESiiZp9NplFJBtDPnkytLSHjyMJwRgT1vjhcofsuV2djIicFPfFhqAxxRbafLfU9CySw7/9SOE6mvlga7IymaB+aWsIxXJyEsfI0jJSkoDcnGi2JgaiiUs4Pskeypyn4jA6CrDXJTKjvcr4HyCrwdSEHKwk6tNVpkwpjlUiB8iKAC1SubAkuUMnO8+UwFSSN6SAZTATIoM/tDPhKYtDd9xrU974Gzd1boLrjVShsfRk6215BpKwY+pzFMeghiQFXCQwQ0AF3KQx4SRN4gkkQL4QoY1SgQwCNcbcpDl+9fQcfiRQ427jpzYnK6CvteZivz8BEVQZ6yVVbyBi6sxJo5s3GdFMZ5ttqCG4BJgqTRW/CLNdmEMBMXKnLx1xJDoZSotAQ4U0gHKgczRBupSNy4wBTLaS6WaGSppXWxBAM0Nw8mM/CS0RnGvXDBAIYRZMMPS6BdWJ7HQRYqsHPguINR/ZFCa+lcKiEDbCWrItryVpYSSQVJijjapwC9TY5LFLfuySQo76niIsE8LkKZalkpRGndgTOZ48jkvJmtrU6Cq0USY1SiNE5Dbezh2F0YCsSVfej34WguRFEFp4rwXOXIUXKUknLYzREHWPh2phLMMZ6cyy+enIbHz2+h9tXZ6M3Xe8Uci/XpeJyM8WZhlxcpEhziT0T05AvOic421iEmdpCTBWnEbxU0Wn/tY5iXGsuxEpDMS6VF2I0kwI1ZcMi9rwfzxgpnjQ/u5gh0FyXJhRTpNLcmxPgjOqoADTRH6c5heAR0NpoX9GnF9iz1PHkxHw7Am9BALkqNE//ECCVsFAP11O1cCNFC9fi1Wmr4Wo8qTCKILJbQIYRQT2RDGOWynk6WB4tPEmEaEjjfLQ77pdH4l6RLy4n2GI83BiDQcboCjBGBk0vFoe3o8j4GIbcCJwblbwHKc/jLAY85TAoAqiMsUgtzJEC32ABsgokgLeWp7mblpuFYjc6cnCzpwQ3BupwtbsOS63luNhYjItNRQSxGLPVBJJ9f6QgCbOVQqxRr7zTXY3rLdVYqqvAZG4GOil4F1HuS3YxRS4ByfCyQqCZPuLsLZDp5YA0gphIUSY70BWlFKQryHhKaSzMCXJGqpct9UBj0ZtM3mbK8DRVhIuh7Pc9MJ0AZhvgeoouASQVxmuQoZAK45XJdVmnVSR3lccSlfAlctXZUHJQAlhhK4kymtE/7aYRrj4Bd/L9cTmZh7k4W4xFWGEkyhZdoVzwJA8hWG4vBl1kSImkQlJiP8Ec8JTFgLc8hgOUMR5BCowzJQUm4OcE8MMHN7C+PPb8U/23e/LX37hQjzvnW3CttwXLnQ1YbK3FQmM5KbEME2WZ6EzkoykyBJOFabjeVoL1gUbc6GjEcksLJosK0B0biRIK2Uk09mXSlBHDGMNXVxMJtlbI8nBAQZArsggwCzHOmf2kAhuy3ZBBLh3nakH9jwN/c3W4GyvAjTaPSviPAG+mGzFv5pqSAnWeKzCRBahBsUWZIgm5MUWYiwIFgqeAi2FyZCDy1AcVUGSngeGCHHw41ICVvAgsZQTRpikqwRUtPlzUeZmjm9QZpKcEH+kD6OdJoZcnQcYiiT53CYJ4Fn0eZCT+qpgI18NsDJUw9cBnBPDdO0vPT+mz6+GFisoHBPBmfyNWzrViuauVSrkBs7WlOF+QjAYygfJgd7TFRRK8CjwYasKbQ8242dZEAFsxX1uLwfRkFLs7QkBjn7eRJixlJeEkL49oCzNksGdz3Bnk+LkiL9Qb6eTqkYw5uS77QUw6OhgjyEYXXiY0RxsoiLaj/vclfDOdwzzMM8HNdD3a2riepInryQSQogyrwGWCKMp84eTKfEUs0J4OUkU1BfZbMxfxaOI8+uOD0Ecj55WcMPSHO6OKKqTKx56EEY4cLwYB8kfR7yiNHp40+l1Jie5UwtQHBzzYPkgKDNOhPGhEJZyAZ49v461rC5UieOx6MlPD3Buuxeq5SlymsrzYWI2J0jz0CeNQG+WD0lAXNMXz0S9MxS2C/GCsE2+OdmCNIF9qqMdkWSm6U+OQ7mABBxV5qB85BL2Tx+Cmpkrh2hTxdhaI4hpTTrRCkrsTYl3sEelkRZnPCN7GWgROE17mNIUYK5Py5EXbXpcFqP5dCRsRQC7uZBrhdhoHN1iAtFepB7ImskwKXAyTpz6oRBAVyImVKANqoCPCB+/cvI2vrl/BUGwg+sjtH5QnoD/MASX2HFRReY9kx6AlOgBBipLodZQlRyZozpT/3GRFvXDQQx7nfVQxFsYq0Aj3W5Px80e38HhlxloEj10by52br/eU/2qBet40+/mXkiz0ChNQLfBFLv2V6qjhDxYKMUylOpAvRGN8JLooZM/VUG4syaVsGEtThxc89TSgefQwOKdPwF5VGd4cHYRZmCCCVMg3N4HA1pKc1gr+NmbwsTBEgJUx/CwN4MxRAU9PGS5GynAkeIy+HGx1vgfIKvBxPhf3Mk1xO4MApmjjZpIWVuOUyTjIQCIUyYXZowIBVCaAqhSotbCcJsDTq8v45dI4bhfE425hPJ5SgpiM98FAMIMLUe6YE0ZiIIkPIZXxoJMi+qiM+5yo91E/fG4oyhSqNUiBBDDaFA86MvHV4zu/2vjhG+vsutJR1HmBDGKoKJ1CbywF3ABkuNM0EeaFvmKCV5yPwaxs1MfEIt3NHRke7qiLi0B9fARdd0GIuRlMpSWgQ/AsleTgqKkOX1OCZEEjGtcUQaTEIGtz+Fkbw52rDx4pz9PCAL42hnA304ITvQAnjjIYjgLsCKCVCOD3Jfw43xyvZ5rhTroBbiVrkwp1cC1WVQRtiTURKuHL4exRhUpYjSYTLTyk7Pne1VV88+AmPqrPxid9LXib0sUUzfPjAncsxPngSlEqhmnMrDRWxSBPCQM86nsEsN9ZTtQDBz1UMEIAp8I5WIg1xaOuHIoxb3SKoP1wzTXnmvTlJ4mmjarYAKS4WiLPn4dWYSJ6C3PRlZ2FzpRU1EXHkHuGk4tGoiKCjzJBMGIcGDhqqMLgzCmYy8vAXkuNJgkOAghYoJU5gghgoKUJvMiV3Uy04WSgAQv1s7DWVoKXNQfeVnoIsDWEo6EKHEiFdqRCK63vz8bcFFowTwotCaAp7mQYkwL1CaA21uLUcYX63zJBXKYYczmMSjpMhRRIEEkxDzPD8Pb0FH7/5D7u0Sj5OglhKoGPeg97tHraYSTMBUvFGRiMD0GbpRaGndXR50jgqJQHXOQx5E6bFDjqr4mpSGNcTLTCW71F+MXbj/7jT+63ZEQ9raRekeBiRTnOmuD4oy6BQnJyEloy0tFMuy4xAZVxMahIiEZpVDjSfLzINAxgJicDc0VZOOpowpfUFkz9LszRGuH2XISRYQRYkdIInoOuKizU5MA5Kw5rDXl4U1YMtjdBrI8deGQiPFNVMBR6uZo/BGjGPCq0wK1MEwJI82gKh0qYAMZSFoxUIohkJKS+JSrlJQELUI3GOj3cSPDEk752/OHN21iiUXI4OgyjNNO3eziiwMoETTSrX6aKm0gLJdVpUPnKUx98vgdcWXgKNOKp4kKQNmYiTShD2uHtgbL/+ONt7KqJD/FNdLMmZ+Qik/Jabqgfysl5G1KTaJMbCzPQSrGgrTgPrUU5qCVTyY0MQ5QbDyFOtgjlMUj09UQq9cNEb2ck0F86mswi2NIIPB1lGJ89A1X23/AeOgQT+bME2ByRPGukh9JjAnlwNFCBEwF0MFKFuZr0i3flWBN5M8+SHNgMt1INsUblezNJ77kCo1SobKn/sepj+2C4GpWwOpWwNsZJOdfLEoC31nGzJBNLKbFYy0jEYKAvOn1dyEzcCWAKhiN56HVSQA+pr9dJnuKMAgEkeB6KGPZRx3iIvqj/LSXZ4+3zNf/5ByzZb9qIcbJaT6cZODXAHYVxAtRkxKOZ5t32fArLxTnoqy7FQF05BilA91YUozU/G7WZKahKT0BNZjJq0xJRHBFCI50DQi2M4KqhDHNJCeieOgpF9p/HiolRn5SAgLJgeogLUmmGzo8PRLSbJez1FUTvkdjpK8NIUeoFwJV0MwJogxtp5pQFjbBC8FbjdbAaq0kOrIalMDITFiCV7vMeqE4mQgADNajsXPDTuWF8fqEfX3Q3YI1aUqcfD6PRXhhKCMAwGU2PnympToHgsb1PngI19UJ3ZQx7K2GEAE6EUv+L4eJKstM6+zFoEaz/bOWGuXIzQzyRRjsnOhjlNHa15SWhp1SIwcpcDFYVYqimhC6XoL+0CF0EsFWYjioq67ygYCTyeAg1oX6nrQ07OVk466nAWkUOjKEWuFqqdFkZITbmyCRwGUFOyKRRLifKG0F2hrDRUaD7qcCSjhwFiT8BeD/XBtfSTGkeNsRqggFWYvRoa9IopyaCxipwMUQVi2z5hqjRJKKGySBNTIdz8XZrCd4aHsCNhnIMUw+scOai3MMO5VRlndH+6PXSRzepr4dU2MfCc1GmWZgFqIILfmQgLMBoa6ykOf1l3/yRwXdfDqdf4kMuKeDZIsOfXDfAg+ZZHxQFB6I0JATFAUHI8/RDupMrEm2dIDDmIkjXAKGWlgg0MUZigDcS+b5IjfKHr6MNvO1t4WtgiHArC6T6OCKDSjYziDafVBjME729aaYpC0tdRZiqykBH5vs31ldSrJk3cuywlk7TSLIJVhONcCVOH1dJgUukwMUwAkfxZSFUHfPBpL4gdUwHaNPWxQzfGBtN+fh8rA/jRXnI5DmAz9FGorUZmvmeGEkMRD8L0F6WQrQShWh1ii8qVL6qGCX3veCvjclQA1KgzV/2zxzYlRftr5oUQEpiuPAzNoGtnDKU9h+C+onjcKdfHkdAhO7OKOaHINfPH8WxcSiJi0dpYjLa2xtQXZyN6spCVJTlIDU8EP5GRggzMqc8aIE4mkrYN+CFNANnBrsgjkY5gbs17CjGGJMzm2jKwUBBEroyZ14AXE63YO7l2uJ6mgXWUkywkmCI5Wg9XI7SpBijiYukukvhqiLlzQapYSqQFEjlO04Qe0J5uN9Qha+n+/CT/masFFI7CnJEd4QnxpMpziSHoMTRFo32muixJwNxVqYyViEFqlH5apID61IJG2Im2uK/9nUpZWmCgdIEykexUSgK5CPKzI6iii745LrFOakY6m3GNI1Ik1OjuDA+jgsT4xgc6sfgYBeay/JRm1eAqqQ0ZDq6Iplri1hzS5qPLRHrzL4P4oCcMA8yD3dE+zvAx84YZmqyMFSVBUdJBgbyUuDI/bAHEsA8OyphK6ymmOEKKfByLAeXovQpSGtRZKHN1xApbypAHZM0u076aaE33AO5SfWYz8rCb0fa8fOeerzbmIuZpABciPPDRLw/Rql9RJnRhOQZhg4HijE8FfRQnOl1VaMMSLO0jzYmgg3/a//Ui12NhUkn6rLinjVkxuFckRB1UbEo9+WjIECAgrBYNKRnoLu0AkONLWQqrRhpacf5phbqjw3oSMpEQ2AUKt0DIbRyQDwLz8YSEQyNdKzrUu/JCvdCUogbQimsOxlrw1BZGgbKMtCVkyQFSsNQTnqE+ycAGaykW+FqsgWuJhhjMcYQFwngvEAXs3wdzATrEjx6sYGG6PCxQhaPUkT+GFJr30CueziWqfe9WZyGOzSRsJPIYKgzevxtUW6jB08rajNhVYh18EKdrR65sSqVszp6XDTJTLSejYX8N79jpiE33qQlP+Hrc6WZaM9LRyc5bGdMHFpCYqh/JKIjoQg9wloM5jVjOL8Bo6XNGC1pQW9aCeoCo1HA80Q644BYGt8EdmaIoEzIRpu0YOqpNOGk8D3g72QBC21lUp409BWkCCCpj1SoL/unJrKea4fVVBssJXKxlGCGizHGmIs0wgxNCNMhuhgOs0dVcBCS/OPgw0uAtVUC4pvfR2rfFwgNKkcaxwjdHrYYDnFGfxCDDk8zgmeIUK4r3IIr4RhaDXNuOJyNPZFibYNampU7nDS+7nbV/Ou+MqqlIDW0m2bjc4UZ6GVTO8WZnqRkdAji0RaehraoHHQmFaEjPgcdcVl0PQ21oTEo8QtGjo8PUjxcEeFAWdDWBDEuNkgLcCGH94IwMgDxQa5wpYBtqHIW+jTM6ytKQ0/hLDiKtAngHxXI9sD1XHtcTbXFpURrXIqzwlyUBaYFFrhAUaQ6swqR0VVwdkkHw8uCs38NGCYDCU1P4J25CD1DPkL1LFGqr4EmR0O08AzRaKeDWF4AlJQdoM4JhLVPEYwto6FvEg4z4yDYGXoikOsULYLw1672ovTG3hIh+suzMFSWhUEaibopurRTj2uMjEWdIAZ1ETGopV0ZEYHcoEAK4n6I83IDnxo035aLKEdLUe9LDXBGBk04KYJAhHs60RPVhD6pjy1ffQKnrygHQ0UZGFKM+eMkspxiwdzNYbCcbIf5eBvMxTCYibDGUH4eMupW4JvYDxf/MljYJcDUOgEeYU2wccmDR+wg5HTDcVLBC1ZW1MtNzdBiqYtWGh1LbRmY2UThmKQNjp+ygJqhAByLOOgbh0HPmA99g+B/P+/+dxf7BQxtxRmX+ityMFiRi36C2ZlHfS4zFdUpiSiNi0Y+jXWpBC3RyxPxHu6I93FHlKsTqY5BLPupBipd9jxgJt8L2TF8xFAs8nOwhKm6vKh8jZRpvFOWh6GSHEyUyI0Vvx/lWIC3shhcTmQwG2uHmRh7TCQHo7H/BuLOvQMevxlMYBVsfcoIQCRs3YvAC2mApkUyNG2yoGmdDVXrPDjaRKPYzBSVFmbg2QggoeEPcWVPyOnwIacdAmW9cOiYxUDbKPwGh/M3/NIJdnVWCnd2lWe/1c9+IpX2uTxSH/XEquQ45McIkB0RitQgPyT7eCPV1xtp/t7I8POi7OiBLD8XOrrRaOiFgtgg5MaGIdyLBx5XH0Yq5L7KsqRARXJhZZgQRHNVeXBV5L4HmEAAhQwWkwhenCOmoxwxUpCL0pnPEdn6FuxCmmHuUwH7kBboWyTCzDEbvNAmWHqVwsBWCA6vAKrcNCiYpcPZVQgPawHOGkZBSjMYCpwIqJrGQd0sCbpWqdAyiflYjRN/QPSi/9ars0woda4s8+lARTYGaTLppTjTlpaA6sQYVJAKyynylMdEoSw6EpWRkaiKFKA2mo+ycH+U8P1QHBWAkqQwpEcEwsfRGha6qlSyMuCoEjyaUIxU2X94KA8LdUVYqisSwOfnAxdJgdcz7TGf4ICpGEeM0Rh2rnEYKeNfwK94BXbh58BxyoNtSCtMeDkwo23jWwW74CYYOebB0DGfVJgJZdNkGHrUQcs2Gwr6EVA0TYS6JUEzT4SeXTZ0bNI+1jJL+Z/54p0/rr4yoVhXRc7yYFU+zhcK0Zedij5y6O4ccuqsTLRlpqMtnXZqKloS4tCWFIfq6DDkB/sgL9wPebEhiKHJxtXSGCaqCtCjnmdAymO/OchUXQlmBNBKUxlWWiovAM6TAlfT7DAb74CJaB6Gk6JQNfoE/JbH8Cy4CvPQLmjY5cAytB2WPlUEsAD2oc0wc6O+6F8HXSYH+rxcqHOToWObA3WLNKjRZVVrIfQIpoFNDkHNuKFmKfyfUd6fr2Xqib1V+a0DRWQqBO5CSR4ulBdjpKwU54uKqbkXYCAvH52paWiKi0MV9cesYC+khnggMcgDQc62sCZH1Fek/qeqAkN19quXSH1qCiL12WirwEZL/QXAmVgz5nIKg8lYHiZiPSkNVCNr+kt4F1+De/4VGPo2QcuhEBbhfXCI6gXXq5qOPeC4lcDYsxz6pEKOcyG07DLBcciDCjcdilZp0HUrhaFrCXTsczv/5j3vL1ltBcLQnuysr88XsADLMVlejZGSCpwvLsNQcQkGcnLRmpyIyqgw0UeE46kXhjjbgWduCI6KEvSVqWw11GiEUyV4SuBqKInUZ6ejRgrUfAFwMtaCuZjkhLEYN4zFBaB54Doiuzfgkr0I++Rp6HjWQpNXBpuoQbhmz8GC3wm7mD4YeJTDwLUUOg750GbyoONYCHXqibLG8dByyoeJb+3XHI/yv01U+e+ursJck968vGcDhSUYLqnEIO3+4nL0FFE2FGahLpFm5MgwpFDZhvBs4GxqQPOuOilPlZSnDmMNdZhQ+ZqpKYo+k22tqQI7bQKqrjmirs4VARyJsmDmCeBojAeGU+JQfulnCKx8A44pszAJ64eWRy3UnMmFE8bhkrsE27hR2MWPwMC7Frru5dAkgCpW6dB2LIIqlaycZTr9vOoZx6furwvJf6s1UFl4oregYKAnrwjd+YXoYpVHvbCSemAuP5iynw+CHWhkM+HATEsLhpo6MKKjEQtQjVWfoqh0ragH2hBAWy01mOhyRtS5z79DlQU4m+CC4SgP9FR3IHP653AVLsE6eQa6Pi3Q8WuFmguVbeZl2AtJlRnzsI4ZhklAG/TcqqHFFEDZIh1qDuTG9vmQs0gb4PhU/mN8BegPV5NQqNqckb5cnxCPajYbRvKRFuiLIAc72BsawFxHD+b6RjCmo6G6GrmuChmJMsF7Xrq2pDz2w+022uow1DV5AXAsypqZTXTGoMATzSNvIG7oK1LYBOxS5qHpTCoL6CQV1oOXtwom4xJcC1ZhLhiASUgX9Mh11e2LoGyZATX7omU1x+J/vC+h/fNVEBHGzRfw1xP8fRDgwMDexARWhqawNDCh8tUmcOS61ANNKbY8h0eqI3j2BM9BTxNcLXVoaZkSwIjnAGPtmJl4Gv5jI1E6/zOEtX4MS8EQbBJnoe5cDVWXWijxysGlkrbPuEwA12DMpx7I74G6Wy1UncvWlWyE/xjl+pcuIYfzEs/c3NfW0OipBQ3y5rocGKqRspQUYaqsADOKMOZsz9NSBaOrAXt9LRE8B31tMhgVaHJsRzwinn8RtwhgrBPaC2qRNflLBNa+A5PATljGT0LTowEyJkJIW2bBJG4cToW3YJOyAL2gczCJGnyqzCvz/bs47N9yGauoGBgrK7UaKco9YycNC3JcSw1lKlU1MPqasOdowdFAG46G2mA4ulCTlYesmskIN+L5F3GzJTwd64ayxquIGfo1vIrvwcC/Fab8fuj6t0PSKAUyphRRwgbhVHLnV/r8/k5j/uD/7q+C/48Wh8PZbKmmxKWSLbXV1Vi3F0HThZORHpxN9OFsrAcbPW3ISZ6FnJr59wrkc5jxtEjkD34AfufnsEuYgD5FF45PM3RYgAaJ66e1oit1A3utOcI/+8TA/8vLmsMRczLjcD3MjKLdzY1reUYGy6aqqhvSp89uqBq6dkYIZ0UABwMMLSeiHTcacmo3wjJGli39GhpNfGqjOa7VXAU74d/xf4exadP/B/l/gb0yNrVRAAAAAElFTkSuQmCC";


        private string imagePart3Data = "iVBORw0KGgoAAAANSUhEUgAAAJIAAAAYCAYAAAARUKQwAAAACXBIWXMAAAsTAAALEwEAmpwYAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAHMSURBVHgB7drdTQJBFAXgO0P4WcIDWAGxAkMBxg6wA40NQGxAoAIqUEpQG1AaEDpwO9h94CeQwHjuRlBhIZtd3zhfQlgGeNqTO3dnxsgRQRBUS6VSyzl3hVcdQ3Whk4F7Hlprx7h8XiwWL7VazT/0WxM3iADVi8XikwZIiH4MlstlNy5QdndgNpu18vn8iCGiGLeFQuEDGWnvfvGnIk2n0we8dYToOIcpr+t5XnczsA2SViJUob4QJeOMMfflcjnKTBQk7Yl0OsMXVSFKLkDP1NCeKeqREKIOQ0Qp1PBQ9qgXRqsRGqhPIUrHoSqdWVSjayFKz6AQtdF826YQZYC26NLiSe1CiDLQXQ+DtSMnRNk4K0T/QIPkC1E2oUWj5AtRBsjQ2K5Wq6EQpefW6/WrLkhWsQ4QCFE6uoF7brFPEqI0vQtROgPP8/xo03Y+n9dRnka45H4bJYb1oyCXyzU0SNHjv15gsCtEyen6Y0+zox+260iVSqWPqsQwUVI9zczmw96Z7clk0kbPpCclOc3RHp3OZCdEKvbwv/ZMWBbQM0o3QvQNIXpDT3S3mc5+M8f+qIHCn5uY8vSoyQUPv50c7Z31NdQKhPsfHvrhF2LTr85r43/mAAAAAElFTkSuQmCC";

        private string imagePart4Data = "PHN2ZyB3aWR0aD0iMTIiIGhlaWdodD0iMTIiIHZpZXdCb3g9IjAgMCAxMiAxMiIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIiB4bWxuczp4bGluaz0iaHR0cDovL3d3dy53My5vcmcvMTk5OS94bGluayIgZmlsbD0ibm9uZSIgb3ZlcmZsb3c9ImhpZGRlbiI+PHBhdGggZD0iTTYgN0M2LjU1MjI5IDcgNyA2LjU1MjI5IDcgNiA3IDUuNDQ3NzIgNi41NTIyOSA1IDYgNSA1LjQ0NzcyIDUgNSA1LjQ0NzcyIDUgNiA1IDYuNTUyMjkgNS40NDc3MiA3IDYgN1pNMi41IDZDMi41IDQuMDY3IDQuMDY3IDIuNSA2IDIuNSA3LjkzMyAyLjUgOS41IDQuMDY3IDkuNSA2IDkuNSA3LjkzMyA3LjkzMyA5LjUgNiA5LjUgNC4wNjcgOS41IDIuNSA3LjkzMyAyLjUgNlpNNiAzLjVDNC42MTkyOSAzLjUgMy41IDQuNjE5MjkgMy41IDYgMy41IDcuMzgwNzEgNC42MTkyOSA4LjUgNiA4LjUgNy4zODA3MSA4LjUgOC41IDcuMzgwNzEgOC41IDYgOC41IDQuNjE5MjkgNy4zODA3MSAzLjUgNiAzLjVaTTAgNi4wMDA2NkMwIDIuNjg2NTggMi42ODY1OCAwIDYuMDAwNjYgMCA5LjMxNDczIDAgMTIuMDAxMyAyLjY4NjU4IDEyLjAwMTMgNi4wMDA2NiAxMi4wMDEzIDkuMzE0NzMgOS4zMTQ3MyAxMi4wMDEzIDYuMDAwNjYgMTIuMDAxMyAyLjY4NjU4IDEyLjAwMTMgMCA5LjMxNDczIDAgNi4wMDA2NlpNNi4wMDA2NiAxQzMuMjM4ODcgMSAxIDMuMjM4ODcgMSA2LjAwMDY2IDEgOC43NjI0NCAzLjIzODg3IDExLjAwMTMgNi4wMDA2NiAxMS4wMDEzIDguNzYyNDQgMTEuMDAxMyAxMS4wMDEzIDguNzYyNDQgMTEuMDAxMyA2LjAwMDY2IDExLjAwMTMgMy4yMzg4NyA4Ljc2MjQ0IDEgNi4wMDA2NiAxWiIgZmlsbD0iIzQyNDI0MiIvPjwvc3ZnPg==";

        private string imagePart5Data = "PHN2ZyB2aWV3Qm94PSIwIDAgOTYgOTYiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIgeG1sbnM6eGxpbms9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkveGxpbmsiIGlkPSJJY29uc19MaW5lU2xpZ2h0Q3VydmVfTSIgb3ZlcmZsb3c9ImhpZGRlbiI+PGcgaWQ9Ikljb25zIj48cGF0aCBkPSJNNzUuNzA3IDMyLjI5M0M3NS4zMDk3IDMxLjkwOTMgNzQuNjc2NyAzMS45MjAzIDc0LjI5MyAzMi4zMTc2IDczLjkxODcgMzIuNzA1MSA3My45MTg3IDMzLjMxOTUgNzQuMjkzIDMzLjcwN0w4Ny41NjkgNDYuOTgzQzg3LjU3MjkgNDYuOTg2OSA4Ny41NzI4IDQ2Ljk5MzMgODcuNTY4OSA0Ni45OTcxIDg3LjU2NyA0Ni45OTg5IDg3LjU2NDYgNDcgODcuNTYyIDQ3TDM0IDQ3QzI0LjQ5MTUgNDYuOTk5OSAxNS40NzM2IDQyLjc3ODggOS4zODMgMzUuNDc3TDkuMjgzIDM1LjM1OUM4LjkyODk5IDM0LjkzNDggOC4yOTgxNSAzNC44NzggNy44NzQgMzUuMjMyIDcuNDQ5ODUgMzUuNTg2IDcuMzkyOTkgMzYuMjE2OCA3Ljc0NyAzNi42NDFMNy44OTUgMzYuODE3QzE0LjM2NjEgNDQuNTM5NiAyMy45MjQ2IDQ5LjAwMDUgMzQgNDlMODcuNTYyIDQ5Qzg3LjU2NzUgNDkuMDAwMSA4Ny41NzE5IDQ5LjAwNDYgODcuNTcxOSA0OS4wMTAxIDg3LjU3MTggNDkuMDEyNyA4Ny41NzA4IDQ5LjAxNTIgODcuNTY5IDQ5LjAxN0w3NC4yOTMgNjIuMjkzQzczLjg5NTcgNjIuNjc2NyA3My44ODQ3IDYzLjMwOTcgNzQuMjY4NCA2My43MDcgNzQuNjUyMSA2NC4xMDQzIDc1LjI4NTIgNjQuMTE1MyA3NS42ODI0IDYzLjczMTYgNzUuNjkwOCA2My43MjM1IDc1LjY5OSA2My43MTUzIDc1LjcwNyA2My43MDdMOTAuNzA3IDQ4LjcwN0M5MS4wOTc0IDQ4LjMxNjUgOTEuMDk3NCA0Ny42ODM1IDkwLjcwNyA0Ny4yOTNaIi8+PC9nPjwvc3ZnPg==";

        private string imagePart6Data = "iVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAOw4AADsOAcy2oYMAAAWWSURBVGhDzZohUB1JEIYRERERiAhEREQEIiLiRMRVccATEREIBOKJJxARCEQE4gRVEQhExImIiIgIBCICERFxIhKBQEQgTiAQiCcQCAT3/Uvvq5nemd19j93j/qoOyfbfPT29PT0zS+a6wNLS0uPV1dU3Kysre8vLy1+Qv/n7L2SMXCFneoZ8FQdZk42ZPwwIYJ5ARgT1zYK8nVKusT3i5ya+nprb/qHMMfAOAyu7qcBmEU1mV0mxYfoBAw2Rf4KBu5YLZIuJPLIhu4FeMRlSXacGLeUczid+jlgPA2xeKqPIE2RRz6SD8xc/m5Jwis1zG/5+UCA4zA14SUB/wnll9NaQX9niQ1lP+mbSfxh9NuDgLY5SC/Qa2SOIe9csPrSmdvGXWlM3yKZRp4MF7x3eMtgPBu3m9QbA5wK+D1JjIu+M1g56vRilMr+PrtsF5mBl5ce9aV1OBPgUA1/zepUjo/QOxlpDfAIviO2ZUfIgA6lu0yp4BlhUBlVm2GgnVhCq7VOefZdOHKPXIlPCx7UVAEF93hvtmzoLOMrYSWDTJOKumXkWTFiL29tumToGM3uMMiodZbJuxuiewzsObaYUZfSFuUuCGA6dzSU21e4HUceDkHitAE1dwWAw+B1OrodPI7X9nhgW4PgWu2fqO2hGjaQAFnyqS92QCB3uhpo8ol1Yojc1NJ0aQsWOSQzMfQXY+c6k5C6YuqjhkSOkXxOwYFKZ/4XutdGyEEdcZyvRmMly4rnK+9zxt01dzFCZmSg1Y1NVgN7XvDKq/aH1+d4C2g98lHJslArQbYdcYvxRKMxZVA48S55t0KnbTHgmjV0qB9k6X5JkdyImvfmQd8OzefVb3aRCxbnZVIDOt0qVzcw3K0ueL6cTU1eA7tRxh3qoK97kIa/mk/EjMNhiyEOUgcaabwI+XslX6FtjmToCug8hr4iVP76ED5Hkrqt1EfL49zdTVUAAz9DrYHaJ6J7wVc9MXYH0zndyDaJbd7wjGUdHh1w7g6fjwYSHDE0VQYGiU+AhVzJGl9xX0G2EXMb6bqoI2KuDhT6PZRzVIKSXxo+Q4CWDYfDckViBHRgtgnw57qmpIsBTckLeuQKLNjBIyf6PzneqJ6aKgC6V/VKSDUK+HG9sqgoc73aawP63EzgLH+As1wHalpA/fE2kgxLSuSjkXfSxiBVM6px0hS55VEDXdhH/FvKQEwUWtTCkbRs9NFUFmoT08Mo2epALXkAftXKNZaoI6KKTALwjPfQb2UfjRyCAPjcyfeWY+NZYpo6Azm9knyuzQs6MXwG6Po4S3mfdUcJzR6WTthl4yMOc3wNUAXcfhVVLoZJ/7xSKBNA/1HF6K+QS409TFcrNUInoE0YyIJ6/QJ/q9Sqnxk+M4sD1pSDRUSN3oXmEPmr3TOC9qQuCvgVFZQQhe6nRHRZO8mqInbraBj79lXIDnbpNNI6JrpRvzH0F6KPsiy+fpr4Dzv0nDGUk+4sH7Rdw6nbdtjKuC54YdF+PrrDEWu2UGWJy5yyBjcqp788q+mQf2uQTi9K/qtpSKgGvlw9bcN4FNoXUxsPMtFj8tU3SOJiAfZefFlWifp2dYV/f7SBowfnavsLhW6P0Dgven6d0lkreVSrAQbLLkMFdo/QGxlHZpDrculHaAQO/NxTCJA7JROe/FsXnPL79gi3lg9GmA4Za1KlsjBlsp7EeWwAfWncaJ/eddbbgS1g55Zzra/Y2QWQ/AueAjc42CjzaYQPRGpiubHKwwZr6vbqXjrvr8F/Lxsxlv4DoMqJ2K05Ty1W3abdg2wKH5avuYvfNiUpTrfbepZkFzrVj6xKUOtPMKjpDfcT3f/p/JlRW2wysTSu10JtEQf9E3uNr6jXUKeyt6BcYnxHdLbRewu/5agIn0onD30fdZHtu7l+7mPH9zYWl+wAAAABJRU5ErkJggg==";

        private string imagePart7Data = "iVBORw0KGgoAAAANSUhEUgAAAYAAAAGACAYAAACkx7W/AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAOw4AADsOAcy2oYMAAAeKSURBVHhe7d0f0KZVGAfghSAIgoVgYWEhCIIgCIIgCIIgWFgIgmAhCBaaCYJgYSFYCIKFIAgWFoIgCBaCIAgWgoUgCIK6fzN9M++8ne993vdrazrnvq6Z30zSt03n2/t+/pxzP5cAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICjPVO5U/n+r3xaea4CwMJS/H+s/LGXXyqvVABY1EeV/eJ/ll8rmgDAoh5WRsX/LJoAwKJ+qIwK/240AYAF5YXvqOjvRxMAWMyVSl74jor+fjQBgMWkqKe4j4r+fjQBgMVoAgCNaQIAjWkCAI1pAovIPI/M+PipkuPe9yrPVwAO0QQmd7mSwj9arJcrAIdoAhPLlf9ooZKfK+4EgC2awKQy1nW0SGdJE7hWAThEE5jQ1pCnJHNAzP0GtmgCk7ldGS3OftIEnq0AHKIJTCRX9scu1jeVfBgC4BBNYCKnLNbXlacqAIdoAhN5o/J7ZbQ4+/miogkAWzSBibxVObYJfFYB2KIJTOTdymhhRskLZIAtmsBEblZGCzPKrQrAFk1gIh9XRgszygcVgC2awESOPSOQ3K14MQxs0QQmksmgo4UZ5UHFOQFgiyYwiVzVZ9vnaGFGyYnhqxWAQzSBSaQJ5BTwaGFGyQA5o6SBLZrAJPJo55QmkMXKuQKAQzSBSZzaBBI7hIAtmsAk8jjo88poYc6LHULAFk1gIh9WRgtzXuwQArZoAhO5Xjl2dlCSHUI+MQkcoglMJP/zH1dGizPKb5X3Kx4JAefRBCaSq/pc3Y8W57zkuwLOCwDn0QQmcrly6g6hLFqmjwKMaAITucgOoeSrypUKwD5NYDKn7hBK8h7BwTFgRBOYzKk7hM6SuUPPVgB2aQKTebHysDJaoEN5VMk3igF2aQKTyXuBjyoXuRvI+wTvBoBdmsCEMh301K2iSc4NfFLxWAg4owlM6OnKncpokbbySyWD5fIzADSBSb1a+akyWqit5FsDNypOEgOawKTySOezymihjkkeJ71ZAXrTBCaW3T6nzBLaT04f544C6EsTmFjGSJzy3eFRvqy8UAF60gQm93blx8powY7N/UpOFHtHAP1oApNL4X6vkpe9o0U7Nvn3P66YOAq9aAILyHbPzBTK9s/Rwp0SdwXQiyawiOwWul3JgbDR4p2S3BXkVLLTxbA+TWAhKdrZNnqRkRL7yc/ICOrsQHJXAOvSBBaTr49lx89oAS+SDJ67W8kjIh+th/VoAgvKbKF8VnK0iBdN7gzyM29VXqoAa9AEFpVDYHmcM1rIf5q8M8hjp2xPNYwO5qYJLCyPhjJo7tgFPjW5O8ipY3cHMC9NYHG5Un+/ctFhc8cmvxzfVtJ03qmkKZhUCv9/mkAD2d2TYXEZMfEktpAek9wl5Ato+TNzCO3dymuVaxXg/0MTaCSzhlKMc8U+WuD/KhlzkZfM9yq5c8hBt9w9pFHlFyynlneT/27g36EJNJSr8RwG+6czh0SkVzSBxWR6aF7qflcZLbiIyG7SBMwXW1BOGucx0YPKkzhtLCJrJgdRWVhOBL9eyYvcvDfQEETkLBlQSSO7DSFnAY59cSQi6yV//2ku7w+uVzKpVFMQ6ZOMl4e/SVPIls6blWzzzC9KPl7vEZLIGsmFXqYOwEnykjlzi9IgblRyBiB3D5k1lLMBOUCW08u7eRIfwxGRJ5MUf9tAARaQK/nHlVGx34/iD7CIFP9jvzuu+AMsQvEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8QdoSPEHaEjxB2hI8Qdo6ErlcWVU7Pej+AMs5PPKqNjvR/EHWMyjyqjg70bxB1jQ1uMfxR9gUV9WRoU/UfwBFpYdQCn0ij9AQ2kC9ysp+kn++WoFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFjGpUt/AtrE7R9IdGLVAAAAAElFTkSuQmCC";

        private string imagePart8Data = "iVBORw0KGgoAAAANSUhEUgAAAEwAAAAYCAYAAABQiBvKAAAACXBIWXMAAAsTAAALEwEAmpwYAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAGZSURBVHgB7ZlbTsJAFIb/OZAIPDHxyQRjuwNdge5Ad6CsAFwBdAfsQJcgO4AVwA6KwVdTnhATnPFMwQahlPJKz5dMMr29fJnb+auQQWijenmBFnfvuHnrdnpYzKAwLlm8KYX+RVVP9r2q0m6GX5FXBl6wElU4WMprCQjSxNH2jekiapUtRiioLIcFnpYKo/f5Z3v72b8RNv2OOjDoQtike1nVwd9FIsyNLFbbg7CDseb5qnYeu4mFxWuWm4YKdQi7KMzYz41b0+I1rGRMV2RlYFFfrjZBqPWOGEI4yLICTWTNA4Rc0Ny0iRTdQ8gFEd0Sz89rCHnxSBb7o/AIwlE4YRMI+eDzmAg7BosxGZghhFxwidRXYcSZVwURhIPwAd8nX+sZ9wcQMnEZWVJLcp3UjFNHIR124wJF142F+WxOEQIIqaiN9DU5hzUq2uU9Im0bg6BR00lOuJPpfyyitjXoFL4C4GnoRtamLMfenyAuI1NEjygmA94Rm2k/QVTWV07cGacZPy4CckX66Y66yboNOfPq+Urv3QB/AbmjgbKhvFmDAAAAAElFTkSuQmCC";

        private string imagePart9Data = "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAAsTAAALEwEAmpwYAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAJWSURBVHgBlVNNaxNRFL13MgW/wGwUoUID/gCzcyPYjYgrXXfVf5DuxJXpQsSVunCdKIouRNOFSrvQ+gMEFyKUoiZgIYaaTGKSmXnv3Xu9702qxY01cHlvMu+ce86Z+xD+43fo1rUKEDTEmSoaWwbrNksHBt+pV5D4LTh3Fq07LM4BOJLoIOBTrxuVSPgFMC9oQShiEUfwT4LTG09rJYm/SG6rCkJfQoFA94SxP7TUHdSUdYXBLbBkHV2bz+bPrM5vPK4K012epgpwIOTErx4Y9qokXvq2uwz+kFhgMMBiKiSmfvnr++HHrU8X2DF4v1pSdCfffVYEsaDU9oE1ZCspZzDm9AZYSpQcVIVoYEE6svjnAA4ECqyygphzJCWZUoYjl8OIsuMiXPZBgbOocr1kLNRQsVfCeGqGMBdHqLJhwrmMbIojUgIlgcxoQmrB2CAZ1bN28mthhSmJusNe24GRn5RCouDEZTB0mVAvQUA9r97Jhyh//MusdJ9EW5+3W+0f37GfTz0BjPIUeHeEpcFUZapvYzHt9dWNeleCohzM5qEZm974Xsdurxw7eQJQpyKe5iDWeKnI1sq4s4NsQgZF8rMVidqRRO9K7uV6El08X86T5BynGYAx6MYTyPsDmex0vf8g3YeHYXxnxVQ3T9bWcW/ijty/2dDLsexfhtD8am0o0Gf9T8AqWW6UyNbNo9aqx/2+TPbVm7W5S4uoZhe9zNDVOlGpqPPgCfUTaupM183D57f3cPj37Psrq5/pina8qt0qYFSuNR/YuM2j+fhB0mwl+8//Ak9GDB97HnnQAAAAAElFTkSuQmCC";


        private string extendedPart1Data = "PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiIHN0YW5kYWxvbmU9InllcyI/PjxjbGJsOmxhYmVsTGlzdCB4bWxuczpjbGJsPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS8yMDIwL21pcExhYmVsTWV0YWRhdGEiPjxjbGJsOmxhYmVsIGlkPSJ7ZjQyYWEzNDItODcwNi00Mjg4LWJkMTEtZWJiODU5OTUwMjhjfSIgZW5hYmxlZD0iMSIgbWV0aG9kPSJTdGFuZGFyZCIgc2l0ZUlkPSJ7NzJmOTg4YmYtODZmMS00MWFmLTkxYWItMmQ3Y2QwMTFkYjQ3fSIgY29udGVudEJpdHM9IjAiIHJlbW92ZWQ9IjAiIC8+PC9jbGJsOmxhYmVsTGlzdD4=";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        #endregion

    }
}