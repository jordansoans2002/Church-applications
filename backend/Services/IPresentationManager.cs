namespace Services
{
    public interface IPresentationManager : IDisposable
    {
        List<Models.PresentationInfo> GetActivePresentations();
        Models.Responses.OperationResult ChangeSlide(string presentationId, int slideChange);
        Models.Responses.OperationResult ControlSlideShow(string presentationId, bool start);
        Task<(byte[] imageData, string contentType)?> GetSlidePreview(string presentationId, int slideNumber);
    }
}