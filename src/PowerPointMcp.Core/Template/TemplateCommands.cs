using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.Template;

/// <inheritdoc cref="ITemplateCommands"/>
public sealed class TemplateCommands : ITemplateCommands
{
    /// <inheritdoc/>
    public TemplateOperationResult Survey(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = ctx.Presentation;
            var masters = new List<TemplateOperationResult.MasterInfo>();
            try
            {
                dynamic master = pres.SlideMaster;
                var mi = new TemplateOperationResult.MasterInfo { Index = 1, Name = (string)master.Name };
                dynamic layouts = master.CustomLayouts;
                int layoutCount = (int)layouts.Count;
                for (int l = 1; l <= layoutCount; l++)
                {
                    dynamic layout = layouts[l];
                    var li = new TemplateOperationResult.LayoutInfo { Index = l, Name = (string)layout.Name };
                    try
                    {
                        dynamic placeholders = layout.Shapes.Placeholders;
                        int phCount = (int)placeholders.Count;
                        for (int ph = 1; ph <= phCount; ph++)
                        {
                            dynamic p = placeholders[ph];
                            li.Placeholders.Add(new TemplateOperationResult.PlaceholderInfo
                            {
                                Ordinal = ph,
                                Type = (int)p.PlaceholderFormat.Type,
                                Name = (string)p.Name,
                                Left = (double)p.Left,
                                Top = (double)p.Top,
                                Width = (double)p.Width,
                                Height = (double)p.Height,
                            });
                        }
                    }
                    catch
                    {
                        // layout placeholders unreadable
                    }

                    mi.Layouts.Add(li);
                }

                masters.Add(mi);
            }
            catch (Exception ex)
            {
                return new TemplateOperationResult { Success = false, ErrorMessage = $"master unreadable: {ex.Message}" };
            }

            var colors = new Dictionary<string, string>();
            string? headingFont = null;
            string? bodyFont = null;
            try
            {
                dynamic theme = ((dynamic)pres.SlideMaster).Theme;
                string[] slots = ["dark1", "light1", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
                for (int i = 0; i < slots.Length; i++)
                {
                    try
                    {
                        int ole = (int)theme.ThemeColorScheme.Item(i + 1).RGB;
                        colors[slots[i]] = $"{(ole >> 16) & 0xFF:X2}{(ole >> 8) & 0xFF:X2}{ole & 0xFF:X2}";
                    }
                    catch
                    {
                        // slot unreadable
                    }
                }

                try
                {
                    headingFont = (string)theme.ThemeFontScheme.MajorFont.Item(1).Name;
                    bodyFont = (string)theme.ThemeFontScheme.MinorFont.Item(1).Name;
                }
                catch
                {
                    // fonts unreadable
                }
            }
            catch
            {
                // theme unreadable
            }

            return new TemplateOperationResult
            {
                Success = true,
                SlideWidth = ctx.Presentation.PageSetup.SlideWidth,
                SlideHeight = ctx.Presentation.PageSetup.SlideHeight,
                ExistingSlides = ctx.Presentation.Slides.Count,
                Masters = masters,
                ThemeColors = colors.Count > 0 ? colors : null,
                ThemeFonts = headingFont != null || bodyFont != null
                    ? new TemplateOperationResult.ThemeFontsInfo { Heading = headingFont, Body = bodyFont }
                    : null,
            };
        });
    }
}
