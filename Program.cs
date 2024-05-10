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



        try
        {

            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            // create a new presentation
            string presentationFilePath = "/Users/pankaj/Documents/VivaGoalsV2/ExportToPPT/MyPresentation.pptx";

            // PresentationDocument presentationDoc = PresentationDocument.Create(presentationFilePath, PresentationDocumentType.Presentation);
            // PresentationPart presentationPart = presentationDoc.AddPresentationPart();

            // generate the presentation part
            CoreGoals.CreatePresentationService createPresentation = new CoreGoals.CreatePresentationService();

            // List<Goal.GoalDetail> goalsArray = createPresentation.GetGoalsArray();

            // createPresentation.CreatePackage(presentationFilePath, goalsArray);

            CoreGoals.CreatePresentationServiceV2 createPresentationV2 = new CoreGoals.CreatePresentationServiceV2();


            List<Goal.GoalDetail> goalsArray = createPresentationV2.GetGoalsArray();

            createPresentationV2.CreatePackage(filePath: presentationFilePath, goalsArray: goalsArray);
            stopwatch.Stop();
            TimeSpan elapsedTime = stopwatch.Elapsed;
            Console.WriteLine($"Time taken for the process: {elapsedTime}");

            // presentationPart.Presentation.Save();

            // Print the memory usage after creating the presentation
           
        }
        catch (System.Exception)
        {
            throw;
        }
    }
}