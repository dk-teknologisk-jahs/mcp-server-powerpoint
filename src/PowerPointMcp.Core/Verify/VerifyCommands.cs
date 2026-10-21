using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Verify;

/// <inheritdoc cref="IVerifyCommands"/>
public sealed class VerifyCommands : IVerifyCommands
{
    private const int MsoTrue = -1;
    private const int MsoFillSolid = 1;
    private const int MsoTextBox = 17;
    private const int MsoPicture = 13;
    private const int MsoLinkedPicture = 11;
    private const int MsoColorTypeRGB = 1;
    private const double BoundsTolerancePt = 2.0;
    private const double OverlapFraction = 0.12;
    private const double CoverFraction = 0.70;
    private const double FullBleedFraction = 0.60;
    private const double MinContrastRatio = 3.0;
    private const double MinTitleBodyRatio = 1.75;

    /// <inheritdoc/>
    public VerifyOperationResult LintDeck(IPresentationBatch batch, int minBodyPt = 14, int hardMinPt = 10, double marginPt = 36)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (minBodyPt <= 0 || hardMinPt <= 0 || hardMinPt > minBodyPt)
        {
            return new VerifyOperationResult { Success = false, ErrorMessage = "Font floors must be positive with hardMinPt <= minBodyPt." };
        }

        return batch.Execute((ctx, ct) =>
        {
            var issues = new List<VerifyOperationResult.VerifyIssue>();
            double slideWidth = ctx.Presentation.PageSetup.SlideWidth;
            double slideHeight = ctx.Presentation.PageSetup.SlideHeight;
            int slideCount = ctx.Presentation.Slides.Count;

            for (int slideIndex = 1; slideIndex <= slideCount; slideIndex++)
            {
                ct.ThrowIfCancellationRequested();
                LintSlide(ctx.Presentation, slideIndex, slideWidth, slideHeight, minBodyPt, hardMinPt, marginPt, issues);
            }

            return new VerifyOperationResult
            {
                Success = true,
                SlideCount = slideCount,
                SlideWidth = slideWidth,
                SlideHeight = slideHeight,
                Errors = issues.Count(i => i.Severity == "error"),
                Warnings = issues.Count(i => i.Severity == "warning"),
                Issues = issues,
            };
        });
    }

    private sealed class ShapeInfo
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
        public double Left { get; set; }
        public double Top { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public double Rotation { get; set; }
        public int Type { get; set; }
        public bool IsPlaceholder { get; set; }
        public bool HasText { get; set; }
        public string Text { get; set; } = string.Empty;
        public List<double> FontSizes { get; } = [];
        public (byte R, byte G, byte B)? TextColor { get; set; }
        public (byte R, byte G, byte B)? FillColor { get; set; }
        public double Area => Width * Height;
    }

    private static void LintSlide(
        PowerPoint.Presentation presentation,
        int slideIndex,
        double slideWidth,
        double slideHeight,
        int minBodyPt,
        int hardMinPt,
        double marginPt,
        List<VerifyOperationResult.VerifyIssue> issues)
    {
        var slide = presentation.Slides[slideIndex];
        var shapes = slide.Shapes;
        int count = shapes.Count;
        var infos = new List<ShapeInfo>(count);

        for (int i = 1; i <= count; i++)
        {
            try
            {
                var shape = shapes[i];
                var info = ReadShape(shape, i);
                if (info == null)
                {
                    continue; // invisible
                }

                if (info.Width <= 1 || info.Height <= 1)
                {
                    Add(issues, "zero-size", "error", slideIndex, i,
                        $"{info.Name} is {info.Width:0.#}x{info.Height:0.#}pt", "resize or delete the shape");
                }

                infos.Add(info);
            }
            catch
            {
                // unreadable shape
            }
        }

        LintGeometry(infos, slideIndex, slideWidth, slideHeight, minBodyPt, hardMinPt, marginPt, issues);
    }

    private static ShapeInfo? ReadShape(PowerPoint.Shape shape, int index)
    {
        var info = new ShapeInfo { Index = index, Name = shape.Name, Type = (int)shape.Type };
        if ((int)shape.Visible != MsoTrue)
        {
            return null;
        }

        info.Left = shape.Left;
        info.Top = shape.Top;
        info.Width = shape.Width;
        info.Height = shape.Height;
        info.Rotation = shape.Rotation;

        try
        {
            if ((int)shape.PlaceholderFormat.Type != 0)
            {
                info.IsPlaceholder = true;
            }
        }
        catch
        {
            // not a placeholder
        }

        try
        {
            if ((int)shape.HasTextFrame == MsoTrue && (int)shape.TextFrame.HasText == MsoTrue)
            {
                info.HasText = true;
                try { info.Text = shape.TextFrame.TextRange.Text ?? string.Empty; } catch { }
                var tr = shape.TextFrame.TextRange;
                int pc = tr.Paragraphs().Count;
                for (int p = 1; p <= pc; p++)
                {
                    try
                    {
                        var para = tr.Paragraphs(p, p);
                        double size = para.Font.Size;
                        if (size <= 0)
                        {
                            try
                            {
                                size = para.Runs(1, 1).Font.Size;
                            }
                            catch
                            {
                                // no runs
                            }
                        }

                        if (size > 0)
                        {
                            info.FontSizes.Add(size);
                        }
                    }
                    catch
                    {
                        // unreadable paragraph
                    }
                }

                try
                {
                    if ((int)tr.Font.Color.Type == MsoColorTypeRGB)
                    {
                        info.TextColor = SplitRgb((int)tr.Font.Color.RGB);
                    }
                }
                catch
                {
                    // mixed or themed color
                }
            }
        }
        catch
        {
            // text frame unreadable
        }

        try
        {
            var fill = shape.Fill;
            if ((int)fill.Visible == MsoTrue && (int)fill.Type == MsoFillSolid)
            {
                info.FillColor = SplitRgb((int)fill.ForeColor.RGB);
            }
        }
        catch
        {
            // fill unreadable
        }

        return info;
    }

    private static void LintGeometry(
        List<ShapeInfo> infos,
        int slideIndex,
        double slideWidth,
        double slideHeight,
        int minBodyPt,
        int hardMinPt,
        double marginPt,
        List<VerifyOperationResult.VerifyIssue> issues)
    {
        foreach (var e in infos)
        {
            // bounds
            if (e.Left < -BoundsTolerancePt || e.Top < -BoundsTolerancePt ||
                e.Left + e.Width > slideWidth + BoundsTolerancePt ||
                e.Top + e.Height > slideHeight + BoundsTolerancePt)
            {
                Add(issues, "off-slide", "error", slideIndex, e.Index,
                    $"{e.Name} extends outside the {slideWidth:0}x{slideHeight:0} slide (L={e.Left:0.#} T={e.Top:0.#} W={e.Width:0.#} H={e.Height:0.#})",
                    "move/resize inside slide bounds");
            }

            // font floors
            foreach (double size in e.FontSizes.Distinct())
            {
                if (size < hardMinPt)
                {
                    Add(issues, "font-floor", "error", slideIndex, e.Index,
                        $"{e.Name}: {size:0.#}pt text (hard minimum is {hardMinPt}pt)",
                        $"raise to >= {minBodyPt}pt (>= {hardMinPt}pt only for template-bound text)");
                }
                else if (size < minBodyPt)
                {
                    Add(issues, "font-min", "warning", slideIndex, e.Index,
                        $"{e.Name}: {size:0.#}pt text below the {minBodyPt}pt body floor",
                        $"raise to >= {minBodyPt}pt unless a deliberate caption");
                }
            }

            // placeholder hygiene
            if (e.IsPlaceholder && !e.HasText)
            {
                Add(issues, "unused-placeholder", "warning", slideIndex, e.Index,
                    $"{e.Name} is an empty placeholder", "fill it or delete it");
            }

            if (!e.HasText && e.FillColor == null && e.Type == MsoTextBox)
            {
                Add(issues, "empty-textbox", "warning", slideIndex, e.Index,
                    $"{e.Name} is a text box with no text and no fill", "delete it");
            }

            // safe margins for unfilled text shapes
            if (e.HasText && e.FillColor == null)
            {
                double half = marginPt / 2.0;
                if (e.Left < marginPt || slideWidth - (e.Left + e.Width) < marginPt ||
                    e.Top < half || slideHeight - (e.Top + e.Height) < half)
                {
                    Add(issues, "tight-margin", "warning", slideIndex, e.Index,
                        $"{e.Name} text sits too close to a slide edge",
                        $"move inward (>= {marginPt:0.#}pt left/right, >= {half:0.#}pt top/bottom)");
                }
            }

            // overflow risk (heuristic): estimated wrapped lines vs box height
            if (e.HasText && e.Text.Trim().Length > 0 && e.FontSizes.Count > 0 && e.Width > 10 && e.Height > 5 && e.Rotation == 0)
            {
                double maxSize = e.FontSizes.Max();
                // average glyph width ~0.52em for Calibri-class sans; 1.28em line height
                double charsPerLine = Math.Max(1, e.Width / (maxSize * 0.52));
                int longestWord = e.Text.Split(' ', '\n', '\r').Max(w => w.Length);
                int lines = (int)Math.Ceiling(e.Text.Trim().Length / charsPerLine) + (longestWord > charsPerLine ? 1 : 0);
                double needed = lines * maxSize * 1.28;
                double available = e.Height + maxSize * 0.4; // allow one partial line of slack + box insets
                if (needed > available * 1.15)
                {
                    Add(issues, "overflow-risk", "warning", slideIndex, e.Index,
                        $"{e.Name}: ~{lines} wrapped line(s) of {maxSize:0.#}pt text need ~{needed:0}pt in a {e.Height:0}pt box",
                        "shorten the text, lower the font size, or grow the shape");
                }
            }

            // contrast against effective background
            if (e.HasText && e.TextColor.HasValue)
            {
                var bg = EffectiveBackground(infos, e, slideWidth, slideHeight);
                double ratio = ContrastRatio(e.TextColor.Value, bg);
                if (ratio < MinContrastRatio)
                {
                    Add(issues, "low-contrast", "warning", slideIndex, e.Index,
                        $"{e.Name}: text contrast {ratio:0.00}:1 against its background",
                        "darken/lighten text or change background (need >= 3:1)");
                }
            }
        }

        // overlaps between text and picture shapes
        var content = infos.Where(s => s.HasText || s.Type == MsoPicture || s.Type == MsoLinkedPicture).ToList();
        for (int a = 0; a < content.Count; a++)
        {
            for (int b = a + 1; b < content.Count; b++)
            {
                var ea = content[a];
                var eb = content[b];
                if (Math.Abs(ea.Rotation) > 0.01 || Math.Abs(eb.Rotation) > 0.01)
                {
                    continue;
                }

                double min = Math.Min(ea.Area, eb.Area);
                if (min <= 0)
                {
                    continue;
                }

                double inter = IntersectionArea(ea, eb);
                if (inter / min > OverlapFraction)
                {
                    Add(issues, "overlap", ea.HasText && eb.HasText ? "error" : "warning", slideIndex, ea.Index,
                        $"{ea.Name} overlaps {eb.Name} by {100 * inter / min:0}% of the smaller shape",
                        "separate them, regroup, or make one an intentional background");
                }
            }
        }

        // title/body hierarchy — only when title and body live in distinct shapes.
        // If every sized shape is a placeholder, the ratio is the template's own design choice
        // (rulebook: match the template), so the warning is suppressed.
        var titleShapes = infos
            .Where(s => s.HasText && s.Top < slideHeight * 0.4 && s.FontSizes.Count > 0 && s.FontSizes.Max() >= 24)
            .ToList();
        var otherTextShapes = infos
            .Where(s => s.HasText && s.FontSizes.Count > 0 && !titleShapes.Contains(s))
            .ToList();
        bool allTemplateBound = infos.Where(s => s.HasText && s.FontSizes.Count > 0).All(s => s.IsPlaceholder);
        if (titleShapes.Count > 0 && otherTextShapes.Count > 0 && !allTemplateBound)
        {
            double titleMax = titleShapes.SelectMany(s => s.FontSizes).Max();
            double bodyMin = otherTextShapes.SelectMany(s => s.FontSizes).Min();
            var inv = System.Globalization.CultureInfo.InvariantCulture;
            if (bodyMin > 0 && titleMax / bodyMin < MinTitleBodyRatio)
            {
                Add(issues, "hierarchy", "warning", slideIndex, 0,
                    $"largest title {titleMax.ToString("0.#", inv)}pt vs body {bodyMin.ToString("0.#", inv)}pt (< {MinTitleBodyRatio.ToString("0.##", inv)}x)",
                    "titles 32-40pt bold; keep >= 1.75x body size contrast");
            }
        }
    }

    private static (byte R, byte G, byte B) EffectiveBackground(List<ShapeInfo> infos, ShapeInfo target, double slideWidth, double slideHeight)
    {
        (byte, byte, byte)? best = null;
        foreach (var cand in infos)
        {
            if (cand.Index >= target.Index)
            {
                break;
            }

            if (cand.HasText || cand.FillColor == null)
            {
                continue;
            }

            double intersection = IntersectionArea(cand, target);
            bool fullBleed = cand.Area >= FullBleedFraction * slideWidth * slideHeight && intersection > 0;
            bool coversTarget = target.Area > 0 && intersection >= CoverFraction * target.Area;
            if (fullBleed || coversTarget)
            {
                best = cand.FillColor; // later (higher z) wins
            }
        }

        return best ?? target.FillColor ?? (255, 255, 255);
    }

    private static double IntersectionArea(ShapeInfo a, ShapeInfo b)
    {
        double x = Math.Max(0, Math.Min(a.Left + a.Width, b.Left + b.Width) - Math.Max(a.Left, b.Left));
        double y = Math.Max(0, Math.Min(a.Top + a.Height, b.Top + b.Height) - Math.Max(a.Top, b.Top));
        return x * y;
    }

    private static (byte R, byte G, byte B) SplitRgb(int oleColor) =>
        ((byte)(oleColor & 0xFF), (byte)((oleColor >> 8) & 0xFF), (byte)((oleColor >> 16) & 0xFF));

    private static double ContrastRatio((byte R, byte G, byte B) c1, (byte R, byte G, byte B) c2)
    {
        double l1 = Luminance(c1);
        double l2 = Luminance(c2);
        double hi = Math.Max(l1, l2);
        double lo = Math.Min(l1, l2);
        return (hi + 0.05) / (lo + 0.05);
    }

    private static double Luminance((byte R, byte G, byte B) c)
    {
        static double Channel(byte v)
        {
            double s = v / 255.0;
            return s <= 0.03928 ? s / 12.92 : Math.Pow((s + 0.055) / 1.055, 2.4);
        }

        return 0.2126 * Channel(c.R) + 0.7152 * Channel(c.G) + 0.0722 * Channel(c.B);
    }

    private static void Add(
        List<VerifyOperationResult.VerifyIssue> issues,
        string rule,
        string severity,
        int slideIndex,
        int shapeIndex,
        string detail,
        string fix) =>
        issues.Add(new VerifyOperationResult.VerifyIssue
        {
            Rule = rule,
            Severity = severity,
            SlideIndex = slideIndex,
            ShapeIndex = shapeIndex == 0 ? null : shapeIndex,
            Detail = detail,
            Fix = fix,
        });
}
