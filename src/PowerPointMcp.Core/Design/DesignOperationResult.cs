namespace Sbroenne.PowerPointMcp.Core.Design;

/// <summary>Result of a spec-driven design build.</summary>
public sealed class DesignOperationResult
{
    /// <summary>Whether the build succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when <see cref="Success"/> is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Slides appended by this call.</summary>
    public int SlidesAdded { get; init; }

    /// <summary>Total slide count after the build.</summary>
    public int SlideCountAfter { get; init; }

    /// <summary>Non-fatal observations (e.g. unknown layout names that fell back to plain).</summary>
    public IReadOnlyList<string>? Warnings { get; init; }
}
