namespace PptWithVba.Models;

public sealed class SlideContent
{
    public string Title { get; set; } = string.Empty;
    public string ImagePath { get; set; } = string.Empty;
    public string[] Content { get; set; } = Array.Empty<string>();
}
