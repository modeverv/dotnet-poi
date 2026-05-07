# Slides

Create and manage slides in a pptx presentation.

## Creating a Presentation

```csharp
using DotnetPoi.XSLF.UserModel;

using var ppt = new XMLSlideShow();
var slide1 = ppt.createSlide();
var slide2 = ppt.createSlide();
```

## Slide Count

```csharp
var slides = ppt.getSlides();
Console.WriteLine($"Slide count: {slides.Count}");
```

## Slide Size

```csharp
// Set slide size (width and height in points)
ppt.setSlideSize(new Dimension(10240000, 7680000));  // 10.24M x 7.68M EMU = 4:3
ppt.setSlideSize(new Dimension(12192000, 6858000));  // 16:9 widescreen
```

Get current slide size:

```csharp
var pageSize = ppt.getPageSize();
Console.WriteLine($"Width: {pageSize.Width}, Height: {pageSize.Height}");
```

## Reading Existing Slides

```csharp
using var stream = File.OpenRead("input.pptx");
using var ppt = new XMLSlideShow(stream);

foreach (var slide in ppt.getSlides())
{
    Console.WriteLine($"Slide #{slide.getSlideNumber()}");
}
```

## Full Runnable Example

See `examples/Phase33PptxExample/` and `examples/UsageSamples/Program.cs` (`CreatePresentation`):

[examples/Phase33PptxExample](https://github.com/modeverv/dotnet-poi/tree/master/examples/Phase33PptxExample)

[examples/UsageSamples](https://github.com/modeverv/dotnet-poi/tree/master/examples/UsageSamples)
