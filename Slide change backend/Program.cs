using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;

public class Program
{
    public static void Main(string[] args)
    {
        var builder = WebApplication.CreateBuilder(args);
        builder.Services.AddControllers();
        builder.Services.AddSingleton<Services.IPresentationManager, Services.PresentationManager>();

        var app = builder.Build();
        app.MapControllers();
        
        app.Run("http://localhost:7777");
    }
}