using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using AltiumSharp.BasicTypes;
using AltiumSharp.Records;
using BxlSharp;
using BxlSharp.Types;
using TextJustification = AltiumSharp.Records.TextJustification;

namespace AltiumSharp
{
    public static class BxlConverter
    {
        /// <summary>
        /// Converts the components and their schematic symbols from a BXL file.
        /// </summary>
        /// <example>
        ///    var progress = new Progress&lt;int&gt;(v =>
        ///    {
        ///        statusProgressBar.Value = v;
        ///        statusStrip.Update();
        ///    });
        ///    var schLib = await Task.Run(() =>
        ///        BxlConverter.ReadSymbolsFromFile(fileDialog.FileName, out logs, progress));
        ///    foreach (var entry in logs)
        ///    {
        ///        switch (entry.Severity)
        ///        {
        ///            case LogSeverity.Information:
        ///                Console.ForegroundColor = ConsoleColor.Blue;
        ///                break;
        ///            case LogSeverity.Warning:
        ///                Console.ForegroundColor = ConsoleColor.Yellow;
        ///                break;
        ///            case LogSeverity.Error:
        ///                Console.ForegroundColor = ConsoleColor.Red;
        ///                break;
        ///        }
        ///        Console.WriteLine(entry.Message);
        ///    }
        ///    Console.ResetColor();
        /// </example>
        /// <param name="fileName">BXL file to be read from.</param>
        /// <param name="logs">Logs of the BXL parsing.</param>
        /// <param name="progress">Progress reporting.</param>
        /// <returns>Returns an AltiumSharp schematic library.</returns>
        public static SchLib ReadSymbolsFromFile(string fileName, out Logs logs, IProgress<int> progress = null)
        {
            var document = BxlDocument.ReadFromFile(fileName, BxlFileType.FromExtension, out logs, progress);
            return ConvertSymbols(document);
        }

        /// <summary>
        /// Converts the components and their schematic symbols from a BXL file.
        /// </summary>
        /// <example>
        ///    var progress = new Progress&lt;int&gt;(v =>
        ///    {
        ///        statusProgressBar.Value = v;
        ///        statusStrip.Update();
        ///    });
        ///    var schLib = await Task.Run(() =>
        ///        BxlConverter.ReadSymbolsFromFile(fileDialog.FileName, out _, progress));
        /// </example>
        /// <param name="fileName">BXL file to be read from.</param>
        /// <param name="progress">Progress reporting.</param>
        /// <returns>Returns an AltiumSharp schematic library.</returns>
        public static SchLib ReadSymbolsFromFile(string fileName, IProgress<int> progress = null)
        {
            return ReadSymbolsFromFile(fileName, out _, progress);
        }

        /// <summary>
        /// Converts the components and their schematic symbols from a BxlSharp document
        /// to the equivalent AltiumSharp schematic library.
        /// <remarks>
        /// This ignores the components' footprints.
        /// </remarks>
        /// </summary>
        /// <param name="document">BxlSharp document to have their schematic library data converted</param>
        /// <param name="hideAllParameters">
        /// Many BXL documents have by default <c>RefDes</c> and <c>Value</c> attributes marked as
        /// being visible. If this argument is <c>true</c> then the visibility is ignored when
        /// creating the respective AltiumSharp component schematic parameters.</param>
        /// <param name="overrideRectangleLineWidth">
        /// Optional value to be used when creating the rectangles in the document, instead of the
        /// values provided in the BxlSharp document. 
        /// </param>
        /// <returns></returns>
        public static SchLib ConvertSymbols(BxlDocument document, bool hideAllParameters = false,
            LineWidth? overrideRectangleLineWidth = null, bool convertOverlines = true)
        {
            var schlib = new SchLib();
            foreach (var component in document.Components)
            {
                var schComponent = new SchComponent
                {
                    LibReference = component.Name,
                    Designator = { Text = $"{component.RefDesPrefix}?" }
                };

                // Read the general parameters for the component
                foreach (var libitem in component.Data)
                {
                    SchPrimitive schPrimitive = null;
                    switch (libitem)
                    {
                        case LibAttribute libAttribute:
                            schPrimitive = new SchParameter
                            {
                                Location = ConvertPoint(libAttribute.Origin),
                                Name = libAttribute.Name,
                                Text = libAttribute.Text,
                                FontId = GetFontId(schlib, document, libAttribute.TextStyle),
                                Justification = ConvertJustification(libAttribute.Justify),
                                Orientation = ConvertRotatation(libAttribute.Rotate),
                                IsMirrored = libAttribute.IsFlipped,
                                IsHidden = true,
                            };
                            if (libAttribute.Name?.Equals("DESCRIPTION",
                                StringComparison.InvariantCultureIgnoreCase) == true)
                            {
                                schComponent.ComponentDescription = libAttribute.Text;
                            }
                            break;
                        case LibWizard libWizard:
                            schPrimitive = new SchParameter
                            {
                                Location = ConvertPoint(libWizard.Origin),
                                Name = libWizard.VarName,
                                Text = libWizard.VarData
                            };
                            break;
                    }

                    if (schPrimitive != null)
                    {
                        schComponent.Add(schPrimitive);
                    }
                }

                // needed for fetching some parameter values
                var footprint = document.GetFootprint(component.PatternName);

                // each component has a set of schematic symbols equivalent to AD display modes and part numbers,
                // so we group them by part number and then iterate over each part's display modes
                var attachedSymbolsByPartNum =
                    component.AttachedSymbols.Where(s => !string.IsNullOrEmpty(s.SymbolName))
                        .GroupBy(s => s.PartNum);

                foreach (var group in attachedSymbolsByPartNum)
                {
                    var partNum = group.Key;
                    if (partNum > schComponent.PartCount)
                    {
                        schComponent.AddPart();
                    }

                    var displayMode = 0;
                    foreach (var attachedSymbol in group)
                    {
                        var symbol = document.GetSymbol(attachedSymbol.SymbolName);
                        if (symbol == null) continue;

                        if (displayMode >= schComponent.DisplayModeCount)
                        {
                            schComponent.AddDisplayMode();
                        }

                        var ignored = new HashSet<LibItem>();

                        // detect rectangles first so the are drawn in the background
                        for (int j = 0; j < symbol.Data.Count; ++j)
                        {
                            if (DetectRectangle(symbol.Data, j, out var rectangle))
                            {
                                ignored.Add(symbol.Data[j]);
                                ignored.Add(symbol.Data[j + 1]);
                                ignored.Add(symbol.Data[j + 2]);
                                ignored.Add(symbol.Data[j + 3]);
                                j += 3;

                                rectangle.LineWidth = overrideRectangleLineWidth ?? rectangle.LineWidth;
                                schComponent.Add(rectangle);
                            }
                        }

                        foreach (var libitem in symbol.Data.Where(li => !ignored.Contains(li)))
                        {
                            SchPrimitive schPrimitive = null;
                            switch (libitem)
                            {
                                case LibLine libLine:
                                    schPrimitive = new SchLine
                                    {
                                        Location = ConvertPoint(libLine.Origin),
                                        Corner = ConvertPoint(libLine.EndPoint),
                                        LineWidth = ConvertWidth(libLine.Width)
                                    };
                                    break;
                                case LibPoly libPoly:
                                    schPrimitive = new SchPolyline
                                    {
                                        Vertices = ConvertPoints(libPoly.Points),
                                        LineWidth = ConvertWidth(libPoly.Width)
                                    };
                                    break;
                                case LibArc libArc:
                                    schPrimitive = new SchArc
                                    {
                                        Location = ConvertPoint(libArc.Origin),
                                        Radius = Coord.FromMils(libArc.Radius),
                                        StartAngle = libArc.StartAngle,
                                        EndAngle = libArc.StartAngle + libArc.SweepAngle,
                                        LineWidth = ConvertWidth(libArc.Width)
                                    };
                                    break;
                                case LibText libText:
                                    schPrimitive = new SchLabel
                                    {
                                        Location = ConvertPoint(libText.Origin),
                                        Text = libText.Text,
                                        FontId = GetFontId(schlib, document, libText.TextStyle),
                                        Justification = ConvertJustification(libText.Justify),
                                        Orientation = ConvertRotatation(libText.Rotate),
                                        IsMirrored = libText.IsFlipped,
                                        IsHidden = !libText.IsVisible,
                                    };
                                    break;
                                case LibPin libPin:
                                    schPrimitive = new SchPin
                                    {
                                        Location = ConvertPoint(libPin.Origin),
                                        Designator = libPin.Designator.Text,
                                        Name = ConvertPinName(libPin.Name.Text, convertOverlines),
                                        NameVisible = libPin.Name.IsVisible,
                                        Orientation = ConvertRotatation(libPin.Rotate),
                                        DesignatorVisible = libPin.Designator.IsVisible,
                                        PinLength = Coord.FromMils(libPin.PinLength)
                                    };
                                    break;
                                case LibAttribute libAttribute:
                                    if (schComponent.GetParameter(libAttribute.Name) != null)
                                    {
                                        break;
                                    }
                                    var parameterValue = libAttribute.Text;
                                    if (string.IsNullOrEmpty(parameterValue) && footprint != null)
                                    {
                                        // when the value is empty, and not present in the component, Ultimate Librarian
                                        // gets the parameter value from the fooprint when exporting to Altium
                                        parameterValue = footprint.GetAttribute(libAttribute.Name)?.Text;
                                    }
                                    if (string.IsNullOrEmpty(parameterValue))
                                    {
                                        parameterValue = component.GetAttribute(libAttribute.Name)?.Text;
                                    }

                                    schPrimitive = new SchParameter
                                    {
                                        Location = ConvertPoint(libAttribute.Origin),
                                        Name = libAttribute.Name,
                                        Text = parameterValue ?? string.Empty,
                                        FontId = GetFontId(schlib, document, libAttribute.TextStyle),
                                        Justification = ConvertJustification(libAttribute.Justify),
                                        Orientation = ConvertRotatation(libAttribute.Rotate),
                                        IsMirrored = libAttribute.IsFlipped,
                                        IsHidden = hideAllParameters || !libAttribute.IsVisible,
                                    };
                                    break;
                                case LibWizard libWizard:
                                    if (schComponent.GetParameter(libWizard.VarName) != null)
                                    {
                                        break;
                                    }
                                    schPrimitive = new SchParameter
                                    {
                                        Location = ConvertPoint(libWizard.Origin),
                                        Name = libWizard.VarName,
                                        Text = libWizard.VarData
                                    };
                                    break;
                                case null:
                                    break;
                                default:
                                    throw new ArgumentOutOfRangeException(nameof(libitem));
                            }

                            if (schPrimitive != null)
                            {
                                schPrimitive.OwnerPartDisplayMode = displayMode;
                                schComponent.Add(schPrimitive);
                            }
                        }

                        displayMode++;
                    }
                }

                schComponent.DisplayMode = 0;
                schComponent.CurrentPartId = 1;
                schlib.Add(schComponent);
            }

            return schlib;
        }

        /// <summary>
        /// Converts the BxlSharp floating point width (possibly in mils) to
        /// the discrete values supported by AltiumSharp.
        /// </summary>
        private static LineWidth ConvertWidth(double width)
        {
            if (width < 10)
            {
                return LineWidth.Small;
            }

            // weird but a width of exactly 10, according to Ultra Librarian, should be the default (Smallest)

            if (width > 10)
            {
                return LineWidth.Medium;
            }

            if (width > 12)
            {
                return LineWidth.Large;
            }

            return LineWidth.Smallest;
        }

        /// <summary>
        /// Converts the BxlSharp text justification to AltiumSharp values.
        /// </summary>
        private static TextJustification ConvertJustification(BxlSharp.Types.TextJustification justify)
        {
            switch (justify)
            {
                case BxlSharp.Types.TextJustification.UpperLeft:
                    return TextJustification.TopLeft;
                case BxlSharp.Types.TextJustification.UpperCenter:
                    return TextJustification.TopCenter;
                case BxlSharp.Types.TextJustification.UpperRight:
                    return TextJustification.TopRight;
                case BxlSharp.Types.TextJustification.Left:
                    return TextJustification.MiddleLeft;
                case BxlSharp.Types.TextJustification.Center:
                    return TextJustification.MiddleCenter;
                case BxlSharp.Types.TextJustification.Right:
                    return TextJustification.MiddleRight;
                case BxlSharp.Types.TextJustification.LowerLeft:
                    return TextJustification.BottomLeft;
                case BxlSharp.Types.TextJustification.LowerCenter:
                    return TextJustification.BottomCenter;
                case BxlSharp.Types.TextJustification.LowerRight:
                    return TextJustification.BottomRight;
                default:
                    throw new ArgumentOutOfRangeException(nameof(justify), justify, null);
            }
        }

        /// <summary>
        /// Converts BxlSharp rotation in degrees to AltiumSharp discrete values.
        /// </summary>
        private static TextOrientations ConvertRotatation(double rotate)
        {
            rotate %= 360;
            if (rotate < 45)
            {
                return TextOrientations.None;
            }

            if (rotate < 135)
            {
                return TextOrientations.Rotated;
            }

            if (rotate < 225)
            {
                return TextOrientations.Flipped;
            }

            if (rotate < 315)
            {
                return TextOrientations.Rotated | TextOrientations.Flipped;
            }

            throw new ArgumentOutOfRangeException(nameof(rotate), rotate, null);
        }

        /// <summary>
        /// Converts a BxlSharp point to AltiumSharp
        /// </summary>
        private static CoordPoint ConvertPoint(Point point)
        {
            return CoordPoint.FromMils(point.X, point.Y);
        }

        /// <summary>
        /// Converts a sequence of BxlSharp points to the equivalent list of AltiumSharp CoordPoints.
        /// </summary>
        private static List<CoordPoint> ConvertPoints(IEnumerable<Point> points)
        {
            return points.Select(ConvertPoint).ToList();
        }

        /// <summary>
        /// Fixes the pin names, and can detect active low parts in pin names (like <c>!TST!/GPO</c>)
        /// and convert them to their equivalent (like <c>T\S\T\/GPO</c>).
        /// </summary>
        /// <param name="text">Text of the pin name to be converted.</param>
        /// <param name="convertOverlines">
        /// <para>
        /// Converts pin name parts that start or end with any of the active low conventional markers
        /// replacing them with the proper AD syntax using back-slashes.
        /// </para>
        /// <para>
        /// Currently supports <c>\</c>, <c>~</c>, <c>*</c> and <c>!</c>.
        /// </para>
        /// </param>
        /// <returns></returns>
        private static string ConvertPinName(string text, bool convertOverlines)
        {
            text = text?.Replace('\\', '|') ?? string.Empty;
            if (!convertOverlines) return text;

            var activeLowMarkers = new[] { '\'', '~', '*', '!' };
            var reALMarkers = Regex.Escape(string.Concat(activeLowMarkers));
            var reText = @"[a-zA-Z_\-\d]+";
            text = Regex.Replace(text ?? "", $"{reText}[{reALMarkers}]|([{reALMarkers}]){reText}\\1?", match =>
            {
                var matchText = match.Value.Trim(activeLowMarkers);
                return Regex.Replace(matchText, @"\w", @"$0\");
            });
            return text;
        }

        /// <summary>
        /// Detects rectangles from 4 consecutive lines that are connected.
        /// </summary>
        /// <param name="data">List of items in the schematic symbol.</param>
        /// <param name="dataIndex">Index of the first line to be tested for being part of a rectangle.</param>
        /// <param name="rectangle">Resulting rectangle.</param>
        /// <returns>Returns if a rectangle was found starting at index <paramref name="dataIndex"/></returns>
        private static bool DetectRectangle(List<LibItem> data, int dataIndex, out SchRectangle rectangle)
        {
            rectangle = null;

            bool LineMatches(Point testPoint, LibLine line, out Point otherPoint)
            {
                // tests if the point given matches the origin or end-point of the given line, returning the other point of the line
                if (testPoint == line.Origin)
                {
                    otherPoint = line.EndPoint;
                    return true;
                }
                else if (testPoint == line.EndPoint)
                {
                    otherPoint = line.Origin;
                    return true;
                }
                else
                {
                    otherPoint = default;
                    return false;
                }
            }

            bool AreLinesConnected(params LibLine[] lines)
            {
                // see if starting from the origin of the first line we can find a sequence of points that go through all the 4 lines.
                var visited = new HashSet<LibLine>();
                var currentPoint = lines[0].Origin;
                for (int n = 0, i = 0; n < lines.Length - 1; ++n)
                {
                    visited.Add(lines[i]);
                    Point nextPoint = default;
                    i = Array.FindIndex(lines, l =>
                        !visited.Contains(l) && LineMatches(currentPoint, l, out nextPoint));
                    if (i == -1)
                    {
                        return false;
                    }
                    currentPoint = nextPoint;
                }

                return true;
            }

            (CoordPoint, CoordPoint) ExtremePoints(params LibLine[] lines)
            {
                var points = lines.Select(l1 => l1.Origin).Concat(lines.Select(l2 => l2.EndPoint)).ToList();
                var c0 = CoordPoint.FromMils(points.Min(p => p.X), points.Min(p => p.Y));
                var c1 = CoordPoint.FromMils(points.Max(p => p.X), points.Max(p => p.Y));
                return (c0, c1);
            }

            if ((dataIndex < data.Count - 3) &&
                data[dataIndex + 0] is LibLine libLine0 && data[dataIndex + 1] is LibLine libLine1 &&
                data[dataIndex + 2] is LibLine libLine2 && data[dataIndex + 3] is LibLine libLine3 &&
                AreLinesConnected(libLine0, libLine1, libLine2, libLine3))
            {
                var (location, corner) = ExtremePoints(libLine0, libLine1, libLine2, libLine3);

                rectangle = new SchRectangle
                {
                    Location = location,
                    Corner = corner,
                    LineWidth = ConvertWidth(libLine0.Width),
                };
                return true;
            }

            return false;
        }

        private static int GetFontId(SchLib schLib, BxlDocument document, string textStyle)
        {
            return 1; // always use the default font
        }
    }
}
