extern alias OfficeInterop;

using System.Text.Json;
using System.Text.Json.Serialization;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Design;

/// <inheritdoc cref="IDesignCommands"/>
public sealed class DesignCommands : IDesignCommands
{
    // office.dll types used as raw ints via dynamic per the project convention (see ShapeCommands).
    private const int MsoShapeRectangle = 1;
    private const int MsoShapeRoundedRectangle = 5;
    private const int MsoTextOrientationHorizontal = 1;

    private const double DefaultWidth = 960;
    private const double DefaultHeight = 540;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true,
    };

    /// <inheritdoc/>
    public DesignOperationResult BuildFromSpec(IPresentationBatch batch, string specJson, bool clearExisting = false)
    {
        ArgumentNullException.ThrowIfNull(batch);
        DeckSpec? spec;
        try
        {
            spec = JsonSerializer.Deserialize<DeckSpec>(specJson, JsonOptions);
        }
        catch (Exception ex)
        {
            return new DesignOperationResult { Success = false, ErrorMessage = $"specJson is not valid JSON: {ex.Message}" };
        }

        if (spec?.Slides == null || spec.Slides.Count == 0)
        {
            return new DesignOperationResult { Success = false, ErrorMessage = "spec.slides must contain at least one slide." };
        }

        double w = spec.SlideWidth is > 0 ? spec.SlideWidth.Value : DefaultWidth;
        double h = spec.SlideHeight is > 0 ? spec.SlideHeight.Value : DefaultHeight;
        var theme = spec.Theme ?? new ThemeSpec();
        var palette = BuildPalette(theme);
        var warnings = new List<string>();

        return batch.Execute((ctx, ct) =>
        {
            ctx.Presentation.PageSetup.SlideWidth = (float)w;
            ctx.Presentation.PageSetup.SlideHeight = (float)h;

            if (clearExisting)
            {
                while (ctx.Presentation.Slides.Count > 0)
                {
                    ctx.Presentation.Slides[ctx.Presentation.Slides.Count].Delete();
                }
            }

            int added = 0;
            foreach (var s in spec.Slides!)
            {
                ct.ThrowIfCancellationRequested();
                var slide = ctx.Presentation.Slides.Add(ctx.Presentation.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                BuildSlide(slide, s, w, palette, theme, warnings);
                if (!string.IsNullOrWhiteSpace(s.Notes))
                {
                    try
                    {
                        slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text = s.Notes;
                    }
                    catch
                    {
                        // notes placeholder unavailable
                    }
                }

                added++;
            }

            return new DesignOperationResult
            {
                Success = true,
                SlidesAdded = added,
                SlideCountAfter = ctx.Presentation.Slides.Count,
                Warnings = warnings.Count > 0 ? warnings : null,
            };
        });
    }

    private sealed record Palette(int Navy, int Accent, int Ink, int Muted);

    private static Palette BuildPalette(ThemeSpec theme)
    {
        return new(HexToOle(theme.Banner, "1E2761"), HexToOle(theme.Accent, "028090"), HexToOle(theme.Ink, "212121"), HexToOle(theme.Muted, "6B7280"));
    }

    private static void BuildSlide(
        PowerPoint.Slide slide,
        SlideSpec s,
        double w,
        Palette p,
        ThemeSpec theme,
        List<string> warnings)
    {
        string layout = (s.Layout ?? "banner").ToLowerInvariant();
        string? title = s.Title;
        string? body = s.Body;
        string headFont = theme.HeadFont ?? "Calibri";
        string bodyFont = theme.BodyFont ?? "Calibri";

        switch (layout)
        {
            case "banner":
                {
                    AddBar(slide, 0, 0, w, 90, p.Navy);
                    AddText(slide, 60, 26, w - 120, 44, title, 28, 0xFFFFFF, true, true, headFont);
                    double y = 130;
                    if (!string.IsNullOrWhiteSpace(body))
                    {
                        AddText(slide, 60, y, w - 120, 300, body, 18, p.Ink, false, false, bodyFont);
                        y += 40 + (body!.Length / 90) * 24;
                    }

                    if (s.Bullets != null)
                    {
                        for (int n = 0; n < s.Bullets.Count; n++)
                        {
                            AddBar(slide, 60, y + (n * 40) + 8, 10, 10, p.Accent);
                            AddText(slide, 84, y + (n * 40), w - 160, 32, s.Bullets[n], 16, p.Ink, false, false, bodyFont);
                        }
                    }

                    break;
                }

            case "split":
                {
                    AddText(slide, 60, 40, w - 120, 50, title, 32, p.Navy, true, false, headFont);
                    if (!string.IsNullOrWhiteSpace(body))
                    {
                        AddText(slide, 60, 120, (w - 140) / 2, 340, body, 16, p.Ink, false, false, bodyFont);
                    }

                    if (s.Bullets != null)
                    {
                        for (int n = 0; n < s.Bullets.Count; n++)
                        {
                            string text = s.Bullets[n];
                            int bar = text.IndexOf('|');
                            string head = bar >= 0 ? text[..bar] : text;
                            string cardBody = bar >= 0 ? text[(bar + 1)..] : string.Empty;
                            AddCard(slide, (w / 2) + 20, 120 + (n * 90), (w - 160) / 2, 76, head, cardBody, p, headFont, bodyFont);
                        }
                    }

                    if (s.Images != null)
                    {
                        foreach (string img in s.Images)
                        {
                            if (!File.Exists(img))
                            {
                                warnings.Add($"image not found, skipped: {img}");
                                continue;
                            }

                            dynamic dynShapes = slide.Shapes;
                            dynShapes.AddPicture(img, 0, -1, (w / 2) + 20, 120.0, (w - 160) / 2, 200.0);
                            ComUtilities.Release(ref dynShapes!);
                        }
                    }

                    break;
                }

            case "stat":
                {
                    AddText(slide, 60, 40, w - 120, 50, title, 32, p.Navy, true, false, headFont);
                    AddText(slide, 60, 150, w - 120, 140, s.Stat?.Value, 72, p.Accent, true, true, headFont);
                    AddText(slide, 60, 300, w - 120, 60, s.Stat?.Label, 18, p.Muted, false, true, bodyFont);
                    if (!string.IsNullOrWhiteSpace(body))
                    {
                        AddText(slide, 60, 360, w - 120, 120, body, 16, p.Ink, false, false, bodyFont);
                    }

                    break;
                }

            case "cards":
                {
                    AddText(slide, 60, 40, w - 120, 50, title, 32, p.Navy, true, false, headFont);
                    var items = s.Cards ?? [];
                    if (items.Count == 0)
                    {
                        warnings.Add("cards layout with no cards");
                        break;
                    }

                    int cols = Math.Min(3, items.Count);
                    int rows = (int)Math.Ceiling(items.Count / (double)cols);
                    double cw = (w - 120 - ((cols - 1) * 24)) / cols;
                    int maxLines = 1;
                    foreach (var c in items)
                    {
                        int est = (int)Math.Ceiling((c.Text?.Length ?? 0) / ((cw - 28) * 0.16));
                        maxLines = Math.Max(maxLines, est);
                    }

                    double ch = rows > 1 ? (380 - ((rows - 1) * 20)) / rows : Math.Max(110, Math.Min(400, 78 + (maxLines * 22)));
                    for (int n = 0; n < items.Count; n++)
                    {
                        int col = n % cols;
                        int row = n / cols;
                        AddCard(slide, 60 + (col * (cw + 24)), 120 + (row * (ch + 20)), cw, ch, items[n].Head ?? string.Empty, items[n].Text ?? string.Empty, p, headFont, bodyFont);
                    }

                    break;
                }

            case "quote":
                {
                    AddBar(slide, 60, 120, 8, 260, p.Accent);
                    AddText(slide, 100, 130, w - 220, 200, "\"" + body + "\"", 30, p.Navy, false, false, headFont);
                    AddText(slide, 100, 360, w - 220, 40, "— " + title, 16, p.Muted, false, false, bodyFont);
                    break;
                }

            default:
                {
                    warnings.Add($"unknown layout '{layout}', built as plain");
                    AddText(slide, 60, 40, w - 120, 50, title, 32, p.Navy, true, false, headFont);
                    if (!string.IsNullOrWhiteSpace(body))
                    {
                        AddText(slide, 60, 120, w - 120, 340, body, 16, p.Ink, false, false, bodyFont);
                    }

                    break;
                }
        }
    }

    private static void AddText(PowerPoint.Slide slide, double l, double t, double width, double height, string? text, double size, int color, bool bold, bool center, string font)
    {
        if (string.IsNullOrEmpty(text))
        {
            return;
        }

        dynamic dynShapes = slide.Shapes;
        PowerPoint.Shape tb = dynShapes.AddTextbox(MsoTextOrientationHorizontal, (float)l, (float)t, (float)width, (float)height);
        ComUtilities.Release(ref dynShapes!);
        var tr = tb.TextFrame.TextRange;
        tr.Text = text;
        tr.Font.Size = (float)size;
        tr.Font.Color.RGB = color;
        tr.Font.Bold = bold ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
        tr.Font.Name = font;
        tr.ParagraphFormat.Alignment = center ? PowerPoint.PpParagraphAlignment.ppAlignCenter : PowerPoint.PpParagraphAlignment.ppAlignLeft;
        tb.TextFrame.MarginLeft = 0;
        tb.TextFrame.MarginRight = 0;
        tb.TextFrame.MarginTop = 0;
        tb.TextFrame.MarginBottom = 0;
        tb.TextFrame.WordWrap = Office.MsoTriState.msoTrue;
    }

    private static void AddBar(PowerPoint.Slide slide, double l, double t, double width, double height, int rgb)
    {
        dynamic dynShapes = slide.Shapes;
        dynShapes.AddShape(MsoShapeRectangle, (float)l, (float)t, (float)width, (float)height);
        int newIndex = slide.Shapes.Count;
        ComUtilities.Release(ref dynShapes!);
        var shape = slide.Shapes[newIndex];
        shape.Fill.ForeColor.RGB = rgb;
        shape.Line.Visible = Office.MsoTriState.msoFalse;
    }

    private static void AddCard(PowerPoint.Slide slide, double l, double t, double width, double height, string head, string text, Palette p, string headFont, string bodyFont)
    {
        dynamic dynShapes = slide.Shapes;
        dynShapes.AddShape(MsoShapeRoundedRectangle, (float)l, (float)t, (float)width, (float)height);
        int newIndex = slide.Shapes.Count;
        ComUtilities.Release(ref dynShapes!);
        var card = slide.Shapes[newIndex];
        card.Fill.ForeColor.RGB = HexToOle(null, "F3F4F6");
        card.Line.Visible = Office.MsoTriState.msoFalse;
        card.TextFrame.TextRange.Text = string.Empty;
        AddText(slide, l + 14, t + 12, width - 28, 24, head, 18, p.Navy, true, false, headFont);
        if (height > 60)
        {
            AddText(slide, l + 14, t + 42, width - 28, height - 54, text, 14, p.Ink, false, false, bodyFont);
        }
    }

    private static int HexToOle(string? hex, string fallback)
    {
        string clean = (hex ?? fallback).Trim().TrimStart('#');
        if (clean.Length != 6 || !int.TryParse(clean, System.Globalization.NumberStyles.HexNumber, null, out int rgb24))
        {
            int.TryParse(fallback, System.Globalization.NumberStyles.HexNumber, null, out rgb24);
        }

        // 0xRRGGBB → OLE COLORREF 0x00BBGGRR
        int r = (rgb24 >> 16) & 0xFF;
        int g = (rgb24 >> 8) & 0xFF;
        int b = rgb24 & 0xFF;
        return (b << 16) | (g << 8) | r;
    }

    private sealed record DeckSpec
    {
        [JsonPropertyName("slideWidth")]
        public double? SlideWidth { get; init; }

        [JsonPropertyName("slideHeight")]
        public double? SlideHeight { get; init; }

        [JsonPropertyName("theme")]
        public ThemeSpec? Theme { get; init; }

        [JsonPropertyName("slides")]
        public List<SlideSpec>? Slides { get; init; }
    }

    private sealed record ThemeSpec
    {
        public string? Banner { get; init; }
        public string? Accent { get; init; }
        public string? Ink { get; init; }
        public string? Paper { get; init; }
        public string? Muted { get; init; }
        public string? HeadFont { get; init; }
        public string? BodyFont { get; init; }
    }

    private sealed record SlideSpec
    {
        public string? Layout { get; init; }
        public string? Title { get; init; }
        public string? Body { get; init; }
        public List<string>? Bullets { get; init; }
        public List<CardSpec>? Cards { get; init; }
        public StatSpec? Stat { get; init; }
        public List<string>? Images { get; init; }
        public string? Notes { get; init; }
    }

    private sealed record CardSpec
    {
        public string? Head { get; init; }
        public string? Text { get; init; }
    }

    private sealed record StatSpec
    {
        public string? Value { get; init; }
        public string? Label { get; init; }
    }
}
