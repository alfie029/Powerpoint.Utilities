using CommandLine;
using Powerpoint.DefaultFontSize;
using Serilog;

Log.Logger = new LoggerConfiguration()
    .MinimumLevel.Debug()
    .WriteTo.Console()
    .CreateLogger();

Parser.Default.ParseArguments<GetFontSizeOptions, SetFontSizeOptions>(args)
    .MapResult<GetFontSizeOptions, SetFontSizeOptions, int>(
        opts =>
        {
            opts.Logger = Log.Logger;
            return VerbMapper.GetFontSize(opts);
        },
        opts =>
        {
            opts.Logger = Log.Logger;
            return VerbMapper.SetFontSize(opts);
        },
        errors =>
        {
            Log.Logger.Error("something went wrong: {Errors}", errors);
            return -1;
        }
    );

