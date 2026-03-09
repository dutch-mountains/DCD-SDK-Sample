using DymoDemo.Core;
using DymoSDK.Interfaces;
using Microsoft.Win32;
using System.IO;
using System.Windows.Input;
using System.Windows.Media.Imaging;

namespace DymoDemo.Wpf.ViewModels;

public class MainViewModel : BaseViewModel
{
    #region Commands

    private ICommand? _openFileCommand;
    public ICommand OpenFileCommand =>
        _openFileCommand ??= new CommandHandler(() => OpenFileAction(), true);

    private ICommand? _printLabelCommand;
    public ICommand PrintLabelCommand =>
        _printLabelCommand ??= new CommandHandler(() => PrintLabelAction(), true);

    private ICommand? _updateLabelCommand;
    public ICommand UpdateLabelCommand =>
        _updateLabelCommand ??= new CommandHandler(() => UpdateValueAction(), true);

    private ICommand? _updatePreviewCommand;
    public ICommand UpdatePreviewCommand =>
        _updatePreviewCommand ??= new CommandHandler(() => UpdatePreviewAction(), true);

    #endregion

    #region Properties

    private IEnumerable<IPrinter> _printers = [];
    public IEnumerable<IPrinter> Printers
    {
        get => _printers;
        set
        {
            _printers = value;
            NotifyPropertyChanged();
            NotifyPropertyChanged(nameof(PrintersFound));
        }
    }

    public int PrintersFound => Printers.Count();

    private string _fileName = string.Empty;
    public string FileName
    {
        get => string.IsNullOrEmpty(_fileName) ? "No file selected" : _fileName;
        set
        {
            _fileName = value;
            NotifyPropertyChanged();
        }
    }

    private BitmapImage? _imageSourcePreview;
    public BitmapImage? ImageSourcePreview
    {
        get => _imageSourcePreview;
        set
        {
            _imageSourcePreview = value;
            NotifyPropertyChanged();
        }
    }

    private List<ILabelObject> _labelObjects = [];
    public List<ILabelObject> LabelObjects
    {
        get => _labelObjects;
        set
        {
            _labelObjects = value;
            NotifyPropertyChanged();
        }
    }

    private ILabelObject? _selectedLabelObject;
    public ILabelObject? SelectedLabelObject
    {
        get => _selectedLabelObject;
        set
        {
            _selectedLabelObject = value;
            NotifyPropertyChanged();
        }
    }

    private string _objectValue = string.Empty;
    public string ObjectValue
    {
        get => _objectValue;
        set
        {
            _objectValue = value;
            NotifyPropertyChanged();
        }
    }

    private IPrinter? _selectedPrinter;
    public IPrinter? SelectedPrinter
    {
        get => _selectedPrinter;
        set
        {
            _selectedPrinter = value;
            NotifyPropertyChanged();
            _ = DisplayConsumableInformation();
        }
    }

    private List<string> _twinTurboRolls = [];
    public List<string> TwinTurboRolls
    {
        get => _twinTurboRolls;
        set
        {
            _twinTurboRolls = value;
            NotifyPropertyChanged();
        }
    }

    private string? _selectedRoll;
    public string? SelectedRoll
    {
        get => _selectedRoll;
        set
        {
            _selectedRoll = value;
            NotifyPropertyChanged();
        }
    }

    private string _consumableInfoText = string.Empty;
    public string ConsumableInfoText
    {
        get => _consumableInfoText;
        set
        {
            _consumableInfoText = value;
            NotifyPropertyChanged();
        }
    }

    #endregion

    private readonly DymoService _dymoService;

    public MainViewModel()
    {
        _dymoService = new DymoService();
        TwinTurboRolls = ["Auto", "Left", "Right"];
        _ = LoadPrintersAsync();
    }

    private async Task LoadPrintersAsync()
    {
        Printers = await _dymoService.GetPrintersAsync();
    }

    private void OpenFileAction()
    {
        var openFileDialog = new OpenFileDialog
        {
            Filter = "DYMO files |*.label;*.dymo|All files|*.*"
        };

        if (openFileDialog.ShowDialog() == true)
        {
            FileName = openFileDialog.FileName;
            _dymoService.LoadLabel(FileName);
            ImageSourcePreview = LoadImage(_dymoService.GetPreviewImage());
            LabelObjects = _dymoService.GetLabelObjects();
        }
    }

    private void PrintLabelAction()
    {
        int copies = 1;
        if (SelectedPrinter != null)
        {
            int? rollSel = null;
            if (SelectedPrinter.Name.Contains("Twin Turbo"))
                rollSel = SelectedRoll == "Auto" ? 0 : SelectedRoll == "Left" ? 1 : 2;

            bool countersUpdated = _dymoService.PrintLabel(SelectedPrinter.Name, copies, rollSel);
            if (countersUpdated)
                UpdatePreviewAction();
        }
    }

    private void UpdateValueAction()
    {
        if (SelectedLabelObject != null)
        {
            _dymoService.UpdateLabelObject(SelectedLabelObject, ObjectValue);
            UpdatePreviewAction();
        }
    }

    private void UpdatePreviewAction()
    {
        ImageSourcePreview = LoadImage(_dymoService.GetPreviewImage());
    }

    private static BitmapImage LoadImage(byte[] array)
    {
        using var ms = new MemoryStream(array);
        var image = new BitmapImage();
        image.BeginInit();
        image.CacheOption = BitmapCacheOption.OnLoad;
        image.StreamSource = ms;
        image.EndInit();
        return image;
    }

    private async Task DisplayConsumableInformation()
    {
        ConsumableInfoText = string.Empty;
        if (SelectedPrinter != null)
        {
            var info = await _dymoService.GetConsumableInfoAsync(SelectedPrinter.DriverName);
            if (info != null)
            {
                ConsumableInfoText = $"Status: {info.Status} \nConsumable: {info.Name} \nLabels remaining: {info.LabelsRemaining}";
            }
        }
    }
}
