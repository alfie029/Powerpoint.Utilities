using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace Powerpoint.DefaultFontSize;

public static class VerbMapper
{
    public static int GetFontSize(GetFontSizeOptions opts)
    {
        opts.Logger?.Information("Read default font-size for {fileName}", opts.FileName);

        try
        {
            using var pptx = PresentationDocument.Open(opts.FileName, false);
            var presentation = pptx.PresentationPart?.Presentation;
            var defaultTextStyle = presentation?.DefaultTextStyle;
            if (defaultTextStyle is null)
            {
                opts.Logger?.Information("\tno default text style, fallback to {fontsize} pt",
                    "18.0");
                return 0;
            }

            Enumerable.Range(1, 9)
                .Select(index => index switch
                {
                    1 => (index, elements: defaultTextStyle.Level1ParagraphProperties?.ChildElements),
                    2 => (index, elements: defaultTextStyle.Level2ParagraphProperties?.ChildElements),
                    3 => (index, elements: defaultTextStyle.Level3ParagraphProperties?.ChildElements),
                    4 => (index, elements: defaultTextStyle.Level4ParagraphProperties?.ChildElements),
                    5 => (index, elements: defaultTextStyle.Level5ParagraphProperties?.ChildElements),
                    6 => (index, elements: defaultTextStyle.Level6ParagraphProperties?.ChildElements),
                    7 => (index, elements: defaultTextStyle.Level7ParagraphProperties?.ChildElements),
                    8 => (index, elements: defaultTextStyle.Level8ParagraphProperties?.ChildElements),
                    9 => (index, elements: defaultTextStyle.Level9ParagraphProperties?.ChildElements),
                    _ => throw new ArgumentOutOfRangeException(nameof(index))
                })
                .Select(x => (x.index, run: x.elements?.First<DefaultRunProperties>()))
                .ToList()
                .ForEach(x => opts.Logger?.Information(
                    "\tLevel {level} default font-size: {fontsize} pt",
                    x.index,
                    ((x.run?.FontSize ?? 1800) / 100f).ToString("F1")));
        }
        catch (IOException ioe)
        {
            opts.Logger?.Error(ioe, "IO error (fileName={FileName})", opts.FileName);
            return 1;
        }

        return 0;
    }

    public static int SetFontSize(SetFontSizeOptions opts)
    {
        opts.Logger?.Information("Change default font-size to {fontsize} for {fileName}",
            opts.Fontsize,
            opts.FileName);

        try
        {
            using var pptx = PresentationDocument.Open(opts.FileName, true);
            if (pptx.PresentationPart is null)
            {
                opts.Logger?.Error("The {fileName} is not a valid Powerpoint document",
                    opts.FileName);
                return 2;
            }

            var presentation = pptx.PresentationPart.Presentation;
            var defaultTextStyle = presentation.DefaultTextStyle ??
                                   (presentation.DefaultTextStyle = new DefaultTextStyle());

            Enumerable.Range(1, 9)
                .Select(index => index switch
                {
                    1 => (index, elements: (defaultTextStyle.Level1ParagraphProperties ??
                                            (defaultTextStyle.Level1ParagraphProperties =
                                                new Level1ParagraphProperties())).ChildElements),
                    2 => (index, elements: (defaultTextStyle.Level2ParagraphProperties ??
                                            (defaultTextStyle.Level2ParagraphProperties =
                                                new Level2ParagraphProperties())).ChildElements),
                    3 => (index, elements: (defaultTextStyle.Level3ParagraphProperties ??
                                            (defaultTextStyle.Level3ParagraphProperties =
                                                new Level3ParagraphProperties())).ChildElements),
                    4 => (index, elements: (defaultTextStyle.Level4ParagraphProperties ??
                                            (defaultTextStyle.Level4ParagraphProperties =
                                                new Level4ParagraphProperties())).ChildElements),
                    5 => (index, elements: (defaultTextStyle.Level5ParagraphProperties ??
                                            (defaultTextStyle.Level5ParagraphProperties =
                                                new Level5ParagraphProperties())).ChildElements),
                    6 => (index, elements: (defaultTextStyle.Level6ParagraphProperties ??
                                            (defaultTextStyle.Level6ParagraphProperties =
                                                new Level6ParagraphProperties())).ChildElements),
                    7 => (index, elements: (defaultTextStyle.Level7ParagraphProperties ??
                                            (defaultTextStyle.Level7ParagraphProperties =
                                                new Level7ParagraphProperties())).ChildElements),
                    8 => (index, elements: (defaultTextStyle.Level8ParagraphProperties ??
                                            (defaultTextStyle.Level8ParagraphProperties =
                                                new Level8ParagraphProperties())).ChildElements),
                    9 => (index, elements: (defaultTextStyle.Level9ParagraphProperties ??
                                            (defaultTextStyle.Level9ParagraphProperties =
                                                new Level9ParagraphProperties())).ChildElements),
                    _ => throw new ArgumentOutOfRangeException(nameof(index))
                })
                .Select(x => (
                    x.index,
                    run: x.elements.First<DefaultRunProperties>() ??
                         x.elements.Append(new DefaultRunProperties()).Last() as DefaultRunProperties
                ))
                .ToList()
                .ForEach(x =>
                {
                    x.run!.FontSize = Convert.ToInt32(opts.Fontsize * 100);
                });
        }
        catch (IOException ioe)
        {
            opts.Logger?.Error(ioe, "IO error (fileName={FileName})", opts.FileName);
            return 1;
        }

        return 0;
    }
}
