namespace Controllers
{
    using Microsoft.AspNetCore.Mvc;
    using Models;
    using Models.Requests;
    using Services;

    [ApiController]
    [Route("[controller]")]
    public class PresentationController : ControllerBase
    {
        private readonly IPresentationManager _presentationManager;

        public PresentationController(IPresentationManager presentationManager)
        {
            _presentationManager = presentationManager;
        }

        [HttpGet("list")]
        public ActionResult<List<PresentationInfo>> GetPresentations()
        {
            return _presentationManager.GetActivePresentations();
        }

        [HttpPost("change-slide")]
        public ActionResult ChangeSlide([FromBody] SlideChangeRequest request)
        {
            var results = new List<object>();
            
            foreach (var presId in request.PresentationIds)
            {
                var result = _presentationManager.ChangeSlide(presId, request.SlideChange);
                results.Add(new
                {
                    presentationId = presId,
                    result.Success,
                    result.Message
                });
            }

            return Ok(new { results });
        }

        [HttpPost("control-show")]
        public ActionResult ControlShow([FromBody] ShowControlRequest request)
        {
            var result = _presentationManager.ControlSlideShow(request.PresentationId, request.Start);
            return Ok(result);
        }

        [HttpGet("preview/{presentationId}/{slideNumber}")]
        public async Task<ActionResult> GetSlidePreview(string presentationId, int slideNumber)
        {
            var result = await _presentationManager.GetSlidePreview(presentationId, slideNumber);
            if (result == null)
            {
                return NotFound();
            }

            return File(result.Value.imageData, result.Value.contentType);
        }
    }
}