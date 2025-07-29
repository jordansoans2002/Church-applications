namespace Models.Requests
{
    public class SlideChangeRequest
    {
        public List<string> PresentationIds { get; set; }
        public int SlideChange { get; set; }
    }
}