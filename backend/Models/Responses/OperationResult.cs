namespace Models.Responses
{
    public class OperationResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }

        public static OperationResult CreateSuccess(string message = null) =>
            new OperationResult { Success = true, Message = message };

        public static OperationResult CreateError(string message) =>
            new OperationResult { Success = false, Message = message };
    }
}