namespace Models
{
    public class PresentationInfo
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public int CurrentSlide { get; set; }
        public int TotalSlides { get; set; }
        public string WindowTitle { get; set; }
        public bool IsRunning { get; set; }
        public string ProcessId { get; set; }
    }
}