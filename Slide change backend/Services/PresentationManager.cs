namespace Services
{
    using Microsoft.Office.Interop.PowerPoint;
    using System.Runtime.InteropServices;
    using Models;
    using Models.Responses;

    public class PresentationManager : IPresentationManager
    {
        private Dictionary<string, (Application App, Dictionary<string, Presentation> Presentations)> powerPointInstances;
        private const int THUMBNAIL_WIDTH = 240;
        private const int THUMBNAIL_HEIGHT = 180;

        public PresentationManager()
        {
            powerPointInstances = new Dictionary<string, (Application, Dictionary<string, Presentation>)>();
            RefreshAllInstances();
        }

        private void RefreshAllInstances()
        {
            // Previous implementation remains the same
            // Just update namespace references
        }

        public List<PresentationInfo> GetActivePresentations()
        {
            // Previous implementation remains the same
            // Just update namespace references
        }

        public OperationResult ChangeSlide(string presentationId, int slideChange)
        {
            foreach (var instance in powerPointInstances)
            {
                if (instance.Value.Presentations.TryGetValue(presentationId, out var presentation))
                {
                    var slideShow = presentation.SlideShowWindow?.View;
                    if (slideShow == null)
                    {
                        return OperationResult.CreateError("Presentation is not in slide show mode");
                    }

                    int newPosition = slideShow.CurrentShowPosition + slideChange;
                    
                    if (newPosition < 1 || newPosition > presentation.Slides.Count)
                    {
                        return OperationResult.CreateError($"Invalid slide number. Must be between 1 and {presentation.Slides.Count}");
                    }

                    slideShow.GotoSlide(newPosition, Microsoft.Office.Core.MsoTriState.msoTrue);
                    return OperationResult.CreateSuccess("Slide changed successfully");
                }
            }
            
            return OperationResult.CreateError("Presentation not found");
        }

        public OperationResult ControlSlideShow(string presentationId, bool start)
        {
            // Previous implementation remains the same
            // Just update return types to use OperationResult
        }

        public async Task<(byte[] imageData, string contentType)?> GetSlidePreview(string presentationId, int slideNumber)
        {
            // Previous implementation remains the same
        }

        public void Dispose()
        {
            // Previous implementation remains the same
        }
    }
}