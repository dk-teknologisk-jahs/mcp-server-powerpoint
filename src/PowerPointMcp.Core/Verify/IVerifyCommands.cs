using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Verify;

/// <summary>Deterministic design verification (structural lint) of a presentation.</summary>
[ServiceCategory("verify", "Verify")]
[McpTool("verify", Title = "Deck Verification", Destructive = false, Category = "quality",
    Description = "Deterministic design lint of an open presentation: off-slide bounds, zero-size shapes, " +
        "text overlaps, font floors, text-background contrast, safe margins, empty placeholders, title/body hierarchy.")]
[McpReadOnlyActions("lint-deck")]
public interface IVerifyCommands
{
    /// <summary>
    /// Lints every slide of the open presentation against the design rulebook.
    /// Reads the LIVE in-memory document (unsaved edits included); never mutates.
    /// </summary>
    /// <param name="batch">The open presentation batch to verify.</param>
    /// <param name="minBodyPt">Recommended body font floor in points (warning below).</param>
    /// <param name="hardMinPt">Hard font floor in points (error below).</param>
    /// <param name="marginPt">Minimum left/right margin for text in points; top/bottom use half of it.</param>
    VerifyOperationResult LintDeck(
        IPresentationBatch batch,
        int minBodyPt = 14,
        int hardMinPt = 10,
        double marginPt = 36);
}
