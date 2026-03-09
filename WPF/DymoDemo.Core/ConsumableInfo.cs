namespace DymoDemo.Core;

/// <summary>
/// Holds consumable/roll status information for a Dymo printer.
/// </summary>
public class ConsumableInfo
{
    public string Status { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string LabelsRemaining { get; set; } = string.Empty;
}
