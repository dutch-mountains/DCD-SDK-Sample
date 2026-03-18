using DymoSDK.Implementations;
using DymoSDK.Interfaces;
using Microsoft.Win32;

namespace DymoDemo.Core;

/// <summary>
/// Provides Dymo printer and label operations.
/// This service wraps the DYMO SDK to allow reuse across WPF, console, and other host applications.
/// </summary>
public class DymoService
{
    private readonly IDymoLabel _label;

    public DymoService()
    {
        DymoSDK.App.Init();
        _label = DymoLabel.LabelSharedInstance;
    }

    /// <summary>
    /// Checks whether DYMO Connect software is installed on this machine.
    /// </summary>
    public static bool IsDymoConnectInstalled()
    {
        using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\DYMO");
        return key != null;
    }

    /// <summary>
    /// Returns the list of Dymo printers installed on the system.
    /// Note: Network printers may not appear immediately after Init(). Use GetPrintersAsync for reliable discovery.
    /// </summary>
    public IEnumerable<IPrinter> GetPrinters()
    {
        return DymoPrinter.Instance.GetPrinters();
    }

    /// <summary>
    /// Discovers Dymo printers, waiting up to <paramref name="timeoutMs"/> milliseconds for network printers to appear.
    /// </summary>
    public async Task<List<IPrinter>> GetPrintersAsync(int timeoutMs = 5000, int pollIntervalMs = 500)
    {
        var printers = DymoPrinter.Instance.GetPrinters().ToList();

        var elapsed = 0;
        while (printers.Count == 0 && elapsed < timeoutMs)
        {
            await Task.Delay(pollIntervalMs);
            elapsed += pollIntervalMs;
            printers = DymoPrinter.Instance.GetPrinters().ToList();
        }

        return printers;
    }

    /// <summary>
    /// Loads a Dymo label file (.label or .dymo) from the specified path.
    /// </summary>
    public void LoadLabel(string filePath)
    {
        _label.LoadLabelFromFilePath(filePath);
    }

    /// <summary>
    /// Generates and returns a preview image (as a byte array) of the currently loaded label.
    /// </summary>
    public byte[] GetPreviewImage()
    {
        _label.GetPreviewLabel();
        return _label.Preview;
    }

    /// <summary>
    /// Returns the list of editable objects in the currently loaded label.
    /// </summary>
    public List<ILabelObject> GetLabelObjects()
    {
        return _label.GetLabelObjects().ToList();
    }

    /// <summary>
    /// Updates the value of a label object.
    /// </summary>
    public void UpdateLabelObject(ILabelObject labelObject, string value)
    {
        _label.UpdateLabelObject(labelObject, value);
    }

    /// <summary>
    /// Prints the currently loaded label on the specified printer.
    /// Returns true if the label contains counter objects that were updated (caller should refresh preview).
    /// </summary>
    public bool PrintLabel(string printerName, int copies, int? rollSelected = null)
    {
        if (rollSelected.HasValue)
            DymoPrinter.Instance.PrintLabel(_label, printerName, copies, rollSelected: rollSelected.Value);
        else
            DymoPrinter.Instance.PrintLabel(_label, printerName, copies);

        // Update counter objects if present
        var counterObjs = _label.GetLabelObjects()
            .Where(w => w.Type == TypeObject.CounterObject).ToList();

        if (counterObjs.Count > 0)
        {
            foreach (var obj in counterObjs)
                _label.UpdateLabelObject(obj, copies.ToString());
            return true;
        }

        return false;
    }

    /// <summary>
    /// Checks whether the printer supports roll/consumable status queries.
    /// </summary>
    public bool IsRollStatusSupported(string printerName)
    {
        return DymoPrinter.Instance.IsRollStatusSupported(printerName);
    }

    /// <summary>
    /// Retrieves consumable/roll information for the specified printer.
    /// Returns null if the printer does not support roll status or is not yet connected.
    /// </summary>
    public async Task<ConsumableInfo?> GetConsumableInfoAsync(string printerName)
    {
        if (!DymoPrinter.Instance.IsRollStatusSupported(printerName))
            return null;

        var rollStatus = await DymoPrinter.Instance.GetRollStatusInPrinter(printerName);
        if (rollStatus == null)
            return null;

        return new ConsumableInfo
        {
            Status = $"{rollStatus.RollStatus}",
            Name = $"{rollStatus.Name}",
            LabelsRemaining = $"{rollStatus.LabelsRemaining}"
        };
    }
}
