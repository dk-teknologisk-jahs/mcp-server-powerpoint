namespace Sbroenne.PowerPointMcp.Core.Verify;

/// <summary>Result of a deterministic deck-verification lint pass.</summary>
public sealed class VerifyOperationResult
{
    /// <summary>Whether the lint ran successfully.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when <see cref="Success"/> is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Number of slides inspected.</summary>
    public int SlideCount { get; init; }

    /// <summary>Slide canvas width in points.</summary>
    public double SlideWidth { get; init; }

    /// <summary>Slide canvas height in points.</summary>
    public double SlideHeight { get; init; }

    /// <summary>Count of issues with severity "error" (must fix).</summary>
    public int Errors { get; init; }

    /// <summary>Count of issues with severity "warning" (judgment calls).</summary>
    public int Warnings { get; init; }

    /// <summary>All detected issues, ordered by slide then shape.</summary>
    public IReadOnlyList<VerifyIssue>? Issues { get; init; }

    /// <summary>One deterministic design issue.</summary>
    public sealed class VerifyIssue
    {
        /// <summary>Stable machine-readable rule code (e.g. off-slide, font-floor, overlap).</summary>
        public required string Rule { get; init; }

        /// <summary>"error" (must fix) or "warning" (judgment call).</summary>
        public required string Severity { get; init; }

        /// <summary>1-based slide index; 0 for slide-level findings such as hierarchy.</summary>
        public int SlideIndex { get; init; }

        /// <summary>1-based shape index, when the issue targets a shape.</summary>
        public int? ShapeIndex { get; init; }

        /// <summary>Human-readable description with concrete values.</summary>
        public required string Detail { get; init; }

        /// <summary>Actionable suggested fix.</summary>
        public required string Fix { get; init; }
    }
}
