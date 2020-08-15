using System;
using System.IO;
using System.Linq;
using AltiumSharp;
using AltiumSharp.Drawing;
using BxlSharp;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            var path = @"e:\shared\bxl\new\";//args[0];
            var files = Enumerable.Concat(
                Directory.GetFiles(path, "*.bxl"),
                Directory.GetFiles(path, "*.xlr")).ToList();
            var n = 0;
            using (var schLibWriter = new SchLibWriter())
            {
                foreach (var filename in files)
                {
                    Console.WriteLine($"{++n}/{files.Count} {Path.GetFileName(filename)}");
                    if (File.Exists($"{filename}.schlib")) continue;

                    var schLib = BxlConverter.ReadSymbolsFromFile(filename, out var logs);
                    foreach (var entry in logs)
                    {
                        switch (entry.Severity)
                        {
                            case LogSeverity.Information:
                                Console.ForegroundColor = ConsoleColor.Blue;
                                break;
                            case LogSeverity.Warning:
                                Console.ForegroundColor = ConsoleColor.Yellow;
                                break;
                            case LogSeverity.Error:
                                Console.ForegroundColor = ConsoleColor.Red;
                                break;
                        }

                        //if (entry.Severity != LogSeverity.Information)
                        {
                            Console.WriteLine(entry.Message);
                        }
                    }

                    Console.ResetColor();
                    Console.WriteLine();

                    schLibWriter.Write(schLib, $"{filename}.schlib", true);

                    var imgDirName = Path.Combine(Path.GetDirectoryName(filename), "img");
                    Directory.CreateDirectory(imgDirName);
                    var imgFilename = Path.Combine(imgDirName, Path.GetFileName(filename));

                    using (var schLibRenderer = new SchLibRenderer(schLib.Header, null))
                    {
                        foreach (var c in schLib.Items)
                        {
                            schLibRenderer.Component = c;
                            for (c.CurrentPartId = 1; c.CurrentPartId <= c.PartCount; ++c.CurrentPartId)
                            {
                                for (c.DisplayMode = 0; c.DisplayMode < c.DisplayModeCount; ++c.DisplayMode)
                                {
                                    try
                                    {
                                        var image = schLibRenderer.RenderAsImage(250, 250, true, true);
                                        image.Save($"{imgFilename}-{c.LibReference}[P={c.CurrentPartId}][M={c.DisplayMode}].jpg");
                                    }
                                    catch
                                    {
                                    }
                                }
                            }
                        }
                    }
                    
                    if (logs.Any(e => e.IsError))
                    {
                        Console.WriteLine("Error, press a key to continue");
                        Console.ReadKey();
                    }
                }
            }

            Console.WriteLine("DONE");
            Console.ReadKey();
        }
    }
}
