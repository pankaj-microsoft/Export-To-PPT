using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Diagnostics;

namespace ExportToPPT;

class Program
{
    static void Main()
    {
        // version of the SDK
        Console.WriteLine("Open XML SDK Version: " + typeof(PresentationDocument).Assembly.GetName().Version);

        // // create a new presentation
        // string presentationFilePath = "/Users/pankaj/Documents/VivaGoalsV2/ExportToPPT/MyPresentation.pptx";
        // PresentationDocument presentationDocument = PresentationDocument.Open(presentationFilePath, true);

        // // open the presentation
        // using (presentationDocument)
        // {

        //     try
        //     {
        //         GeneratedCode.GeneratedClass generatedClass = new GeneratedCode.GeneratedClass();
        //         PresentationPart presentationPart = presentationDocument.PresentationPart;


        //         // time taken to create 100 slides

        //         Stopwatch stopwatch = new Stopwatch();

        //         stopwatch.Start();
        //         for (int i = 0; i < 1; i++)
        //         {
        //             SlideId lastSlideId = presentationPart.Presentation.SlideIdList.LastChild as SlideId;

        //             uint slideIdValue = 1;
        //             if (lastSlideId != null)
        //             {
        //                 slideIdValue = lastSlideId.Id + 1;
        //             }

        //             // create a new slide 
        //             SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();

        //             // generate slide part id
        //             string slidePartId = presentationPart.GetIdOfPart(slidePart);

        //             SlideId newSlideId = new SlideId
        //             {
        //                 RelationshipId = slidePartId,
        //                 Id = slideIdValue
        //             };

        //             // generatedClass.CreateSlidePart(slidePart);
        //             BasicPresentation.GeneratedClass basicPresentation = new BasicPresentation.GeneratedClass();
        //             basicPresentation.CreateSlidePart(slidePart);

        //             presentationPart.Presentation.SlideIdList.InsertAfter(newSlideId, lastSlideId);
        //             OpenXmlValidator validator = new OpenXmlValidator();

        //             // Validate the document
        //             var errors = validator.Validate(presentationDocument);

        //             // Process the validation errors (if needed)
        //             foreach (ValidationErrorInfo errInfo in errors)
        //             {
        //                 Console.WriteLine($"Error: \"{errInfo.Description}\"");
        //                 Console.WriteLine($"XPath: {errInfo.Path.XPath}");
        //             }
        //         }




        //         // add the slide to the presentation
        //         System.Console.WriteLine("Table created successfully");

        //         stopwatch.Stop();

        //         TimeSpan timeTaken = stopwatch.Elapsed;

        //         Console.WriteLine("Time taken: " + timeTaken.ToString());
        //     }
        //     catch (System.Exception)
        //     {

        //         throw;
        //     }
        // }



        try
        {
            // create a new presentation

            string presentationFilePath = "/Users/pankaj/Documents/VivaGoalsV2/ExportToPPT/MyPresentation.pptx";



            // PresentationDocument presentationDoc = PresentationDocument.Create(presentationFilePath, PresentationDocumentType.Presentation);
            // PresentationPart presentationPart = presentationDoc.AddPresentationPart();

            // // generate the presentation part

            CoreGoals.CreatePresentationService createPresentation = new CoreGoals.CreatePresentationService();

            List<Goal.GoalDetail> goalsArray = createPresentation.GetGoalsArray();

            createPresentation.CreatePackage(presentationFilePath, goalsArray);


            // presentationPart.Presentation.Save();
        }
        catch (System.Exception)
        {

            throw;
        }



    }

}
