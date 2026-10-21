using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Design;

/// <summary>Spec-driven deck construction: builds a full presentation in one call from a JSON deck spec.</summary>
[ServiceCategory("design", "Design")]
[McpTool("design", Title = "Deck Design", Destructive = true, Category = "content",
    Description = "Build slides from a structured deck spec (single pass): banner, split, stat, cards, quote, plain layouts with themed typography, chips and speaker notes.")]
public interface IDesignCommands
{
    /// <summary>
    /// Appends slides to the open presentation from a JSON spec. The spec shape:
    /// { slideWidth?, slideHeight?, theme?: {banner,accent,ink,paper,muted,headFont,bodyFont},
    ///   slides: [ { layout: banner|split|stat|cards|quote|plain, title?, body?, bullets?: [string],
    ///     cards?: [{head,text}], stat?: {value,label}, images?: [path], notes? } ] }.
    /// Colors are 6-hex-digit strings (no #). Geometry is points (default canvas 960x540).
    /// </summary>
    /// <param name="batch">The open presentation batch to build into (typically a fresh create).</param>
    /// <param name="specJson">The deck spec, as a JSON string.</param>
    /// <param name="clearExisting">Remove all existing slides before building. Default false.</param>
    DesignOperationResult BuildFromSpec(IPresentationBatch batch, string specJson, bool clearExisting = false);
}
