using CommandLine;
using Serilog;

namespace Powerpoint.DefaultFontSize;

public class DefaultOptions
{
    [Option('f', "file", Required = true, HelpText = "Indicate the Powerpoint file to be operated")]
    public string FileName { get; set; } = string.Empty;

    internal ILogger? Logger { get; set; }
}

[Verb("get", isDefault: true, HelpText = "Read default font-size for specified document.")]
public class GetFontSizeOptions : DefaultOptions
{
}

[Verb("set", HelpText = "Change default font-size for specified document.")]
public class SetFontSizeOptions : DefaultOptions
{
    [Option('s', "size", Required = true, HelpText = "The new default font-size.")]
    public float Fontsize { get; set; }
}

