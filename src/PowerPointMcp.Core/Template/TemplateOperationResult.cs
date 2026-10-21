namespace Sbroenne.PowerPointMcp.Core.Template;

/// <summary>Result of a template survey.</summary>
public sealed class TemplateOperationResult
{
    /// <summary>Whether the survey succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when <see cref="Success"/> is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Canvas width in points.</summary>
    public double SlideWidth { get; init; }

    /// <summary>Canvas height in points.</summary>
    public double SlideHeight { get; init; }

    /// <summary>Slides currently in the document.</summary>
    public int ExistingSlides { get; init; }

    /// <summary>Masters with their layouts.</summary>
    public IReadOnlyList<MasterInfo>? Masters { get; init; }

    /// <summary>Theme color slots (dark1, light1, accent1..6) as RRGGBB hex.</summary>
    public IReadOnlyDictionary<string, string>? ThemeColors { get; init; }

    /// <summary>Theme heading/body font names.</summary>
    public ThemeFontsInfo? ThemeFonts { get; init; }

    /// <summary>One slide master and its custom layouts.</summary>
    public sealed class MasterInfo
    {
        /// <summary>1-based master index.</summary>
        public int Index { get; set; }

        /// <summary>Master name.</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>Layouts available on this master.</summary>
        public List<LayoutInfo> Layouts { get; set; } = [];
    }

    /// <summary>One custom layout with its placeholders.</summary>
    public sealed class LayoutInfo
    {
        /// <summary>1-based layout index within the master.</summary>
        public int Index { get; set; }

        /// <summary>Layout name (match target for design layoutName).</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>Placeholders defined by the layout.</summary>
        public List<PlaceholderInfo> Placeholders { get; set; } = [];
    }

    /// <summary>One placeholder definition.</summary>
    public sealed class PlaceholderInfo
    {
        /// <summary>1-based position in the layout's placeholder collection (match target for design placeholders[].ordinal).</summary>
        public int Ordinal { get; set; }

        /// <summary>PpPlaceholderType value (2 date, 3 footer, 4 subtitle, 6 body, 7 object, 13 title, 14 centered title).</summary>
        public int Type { get; set; }

        /// <summary>Shape name (match target for design placeholders[].name).</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>Left in points.</summary>
        public double Left { get; set; }

        /// <summary>Top in points.</summary>
        public double Top { get; set; }

        /// <summary>Width in points.</summary>
        public double Width { get; set; }

        /// <summary>Height in points.</summary>
        public double Height { get; set; }
    }

    /// <summary>Theme major/minor font pair.</summary>
    public sealed class ThemeFontsInfo
    {
        /// <summary>Heading (major) font.</summary>
        public string? Heading { get; set; }

        /// <summary>Body (minor) font.</summary>
        public string? Body { get; set; }
    }
}
