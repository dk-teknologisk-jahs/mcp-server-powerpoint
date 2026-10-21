using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Template;

/// <summary>Template inspection: masters, layouts, placeholder geometry, theme colors and fonts.</summary>
[ServiceCategory("template", "Template")]
[McpTool("template", Title = "Template Operations", Destructive = false, Category = "content",
    Description = "Inventory an opened presentation or template (.pptx/.potx): masters with layouts, placeholder indexes/types/positions, theme color palette and font pairs, slide size.")]
[McpReadOnlyActions("survey")]
public interface ITemplateCommands
{
    /// <summary>
    /// Surveys the open document: every master's layouts with placeholder geometry, plus theme
    /// colors/fonts and canvas size. Use before template-aware builds (design build-from-spec
    /// with layoutName / placeholders) so content maps onto real layouts.
    /// </summary>
    TemplateOperationResult Survey(IPresentationBatch batch);
}
