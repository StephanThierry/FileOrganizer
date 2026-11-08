using Avalonia.Controls;
using Avalonia.Platform.Storage;
using FileOrganizer.Desktop.ViewModels;

namespace FileOrganizer.Desktop.Views;

public partial class ScanView : UserControl
{
    public ScanView()
    {
        InitializeComponent();
        DataContextChanged += (_, _) =>
        {
            if (DataContext is not ScanViewModel vm) return;
            vm.PickFolder = async () =>
            {
                var top = TopLevel.GetTopLevel(this);
                if (top == null) return null;
                var folders = await top.StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions { Title = "Folder to organize", AllowMultiple = false });
                return folders.Count > 0 ? folders[0].TryGetLocalPath() : null;
            };
        };
    }
}
