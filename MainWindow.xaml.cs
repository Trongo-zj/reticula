using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Interop;
using Microsoft.Win32;

namespace ReticulaApp;

public partial class MainWindow : Window
{
    private const int Columns = 6;
    private const int Rows = 5;
    private const string TileIndexDataFormat = "ReticulaApp.TileIndex";

    private Point? _dragStartPoint;
    private AppTile? _dragSourceTile;

    public ObservableCollection<AppTile> Tiles { get; }

    public MainWindow()
    {
        Tiles = new ObservableCollection<AppTile>(Enumerable.Range(0, Columns * Rows).Select(_ => new AppTile()));
        InitializeComponent();
        DataContext = this;
    }

    private void TileButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button button)
        {
            return;
        }

        if (button.DataContext is not AppTile tile || !tile.HasApp)
        {
            return;
        }

        try
        {
            if (string.IsNullOrWhiteSpace(tile.CommandPath))
            {
                return;
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = tile.CommandPath,
                UseShellExecute = true
            };

            if (!string.IsNullOrWhiteSpace(tile.WorkingDirectory))
            {
                startInfo.WorkingDirectory = tile.WorkingDirectory;
            }

            bool isShortcut = tile.CommandPath.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase);
            if (!isShortcut && !string.IsNullOrWhiteSpace(tile.Arguments))
            {
                startInfo.Arguments = tile.Arguments;
            }

            Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"No se pudo iniciar la aplicación.\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void TileButton_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is FrameworkElement element && element.DataContext is AppTile tile)
        {
            _dragSourceTile = tile;
            _dragStartPoint = e.GetPosition(null);
        }
    }

    private void TileButton_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (e.LeftButton != MouseButtonState.Pressed || _dragStartPoint is null || _dragSourceTile is null)
        {
            return;
        }

        var currentPosition = e.GetPosition(null);
        var diff = currentPosition - _dragStartPoint.Value;
        if (Math.Abs(diff.X) < SystemParameters.MinimumHorizontalDragDistance &&
            Math.Abs(diff.Y) < SystemParameters.MinimumVerticalDragDistance)
        {
            return;
        }

        int sourceIndex = Tiles.IndexOf(_dragSourceTile);
        if (sourceIndex >= 0 && _dragSourceTile.HasApp)
        {
            var dataObject = new DataObject();
            dataObject.SetData(TileIndexDataFormat, sourceIndex);
            DragDrop.DoDragDrop((DependencyObject)sender, dataObject, DragDropEffects.Move);
        }

        _dragStartPoint = null;
        _dragSourceTile = null;
    }

    private void TileButton_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (sender is not FrameworkElement element || element.DataContext is not AppTile tile)
        {
            return;
        }

        if (tile.HasApp)
        {
            // Let the click handler process the launch without opening the picker.
            return;
        }

        e.Handled = true;
        _dragStartPoint = null;
        _dragSourceTile = null;

        int index = Tiles.IndexOf(tile);
        if (index >= 0)
        {
            PromptForApplicationSelection(index);
        }
    }

    private void TileButton_DragEnter(object sender, DragEventArgs e) => UpdateDragEffects(e);

    private void TileButton_DragOver(object sender, DragEventArgs e) => UpdateDragEffects(e);

    private void TileButton_Drop(object sender, DragEventArgs e)
    {
        e.Handled = true;

        if (sender is not FrameworkElement element || element.DataContext is not AppTile targetTile)
        {
            return;
        }

        int targetIndex = Tiles.IndexOf(targetTile);
        if (targetIndex < 0)
        {
            return;
        }

        if (e.Data.GetDataPresent(TileIndexDataFormat))
        {
            int sourceIndex = (int)e.Data.GetData(TileIndexDataFormat);
            MoveTile(sourceIndex, targetIndex);
            return;
        }

        var state = ExtractAppState(e.Data);
        if (state.HasValue)
        {
            AssignAppToSlot(state.Value, targetIndex);
        }
    }

    private void GridContainer_DragEnter(object sender, DragEventArgs e) => UpdateDragEffects(e);

    private void GridContainer_DragOver(object sender, DragEventArgs e) => UpdateDragEffects(e);

    private void GridContainer_Drop(object sender, DragEventArgs e)
    {
        HandleBackgroundDrop(e);
    }

    private void GridContainer_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount != 2)
        {
            return;
        }

        if (FindAncestor<Button>(e.OriginalSource as DependencyObject) is Button button &&
            button.DataContext is AppTile tile &&
            tile.HasApp)
        {
            // Double click sobre un botón con aplicación: no abrir el selector aquí.
            return;
        }

        e.Handled = true;
        PromptForApplicationSelection(null);
    }

    private void Window_DragEnter(object sender, DragEventArgs e) => UpdateDragEffects(e);

    private void Window_DragOver(object sender, DragEventArgs e) => UpdateDragEffects(e);

    private void Window_Drop(object sender, DragEventArgs e)
    {
        HandleBackgroundDrop(e);
    }

    private void HandleBackgroundDrop(DragEventArgs e)
    {
        e.Handled = true;

        if (e.Data.GetDataPresent(TileIndexDataFormat))
        {
            int sourceIndex = (int)e.Data.GetData(TileIndexDataFormat);
            if (IsValidIndex(sourceIndex))
            {
                Tiles[sourceIndex].Clear();
            }
            return;
        }

        var state = ExtractAppState(e.Data);
        if (!state.HasValue)
        {
            return;
        }

        int freeSlot = FindFirstFreeSlot();
        if (freeSlot == -1)
        {
            MessageBox.Show(this, "No hay espacios libres en la retícula.", "Sin espacio", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        AssignAppToSlot(state.Value, freeSlot);
    }

    private void PromptForApplicationSelection(int? targetIndex)
    {
        var dialog = CreateApplicationPickerDialog();
        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        string selectedPath = dialog.FileName;
        if (!TryCreateStateFromPath(selectedPath, out var state))
        {
            MessageBox.Show(this, "No se pudo cargar la aplicación seleccionada.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        if (targetIndex.HasValue)
        {
            AssignAppToSlot(state, targetIndex.Value, replaceExisting: true);
            return;
        }

        int freeSlot = FindFirstFreeSlot();
        if (freeSlot == -1)
        {
            MessageBox.Show(this, "No hay espacios libres en la retícula.", "Sin espacio", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        AssignAppToSlot(state, freeSlot);
    }

    private static OpenFileDialog CreateApplicationPickerDialog()
    {
        string startMenuPath = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu);
        string programsPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms);

        return new OpenFileDialog
        {
            Title = "Selecciona una aplicación",
            Filter = "Aplicaciones|*.exe;*.lnk;*.bat;*.cmd;*.ps1|Accesos directos (*.lnk)|*.lnk|Aplicaciones (*.exe)|*.exe|Todos los archivos|*.*",
            Multiselect = false,
            CheckFileExists = true,
            CheckPathExists = true,
            DereferenceLinks = false,
            RestoreDirectory = true,
            InitialDirectory = Directory.Exists(startMenuPath) ? startMenuPath : (Directory.Exists(programsPath) ? programsPath : Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles))
        };
    }

    private void UpdateDragEffects(DragEventArgs e)
    {
        if (e.Data.GetDataPresent(TileIndexDataFormat))
        {
            e.Effects = e.AllowedEffects.HasFlag(DragDropEffects.Move) ? DragDropEffects.Move : DragDropEffects.None;
        }
        else if (HasExternalAppData(e.Data))
        {
            e.Effects = GetExternalDropEffect(e.AllowedEffects);
        }
        else
        {
            e.Effects = DragDropEffects.None;
        }

        e.Handled = true;
    }

    private static DragDropEffects GetExternalDropEffect(DragDropEffects allowedEffects)
    {
        if (allowedEffects.HasFlag(DragDropEffects.Copy))
        {
            return DragDropEffects.Copy;
        }

        if (allowedEffects.HasFlag(DragDropEffects.Link))
        {
            return DragDropEffects.Link;
        }

        if (allowedEffects.HasFlag(DragDropEffects.Move))
        {
            return DragDropEffects.Move;
        }

        return DragDropEffects.None;
    }

    private void MoveTile(int fromIndex, int toIndex)
    {
        if (!IsValidIndex(fromIndex) || !IsValidIndex(toIndex) || fromIndex == toIndex)
        {
            return;
        }

        var sourceTile = Tiles[fromIndex];
        if (!sourceTile.HasApp)
        {
            return;
        }

        var sourceState = sourceTile.CreateSnapshot();

        if (!TryMakeRoom(toIndex, fromIndex, out var displacedState, out var displacedDestination))
        {
            return;
        }

        Tiles[toIndex].Apply(sourceState);

        if (displacedDestination.HasValue && displacedDestination.Value == fromIndex && displacedState.HasValue)
        {
            Tiles[fromIndex].Apply(displacedState.Value);
        }
        else
        {
            Tiles[fromIndex].Clear();
        }
    }

    private void AssignAppToSlot(AppTileState state, int targetIndex, bool replaceExisting = false)
    {
        if (!IsValidIndex(targetIndex) || !state.HasApp)
        {
            return;
        }

        if (!replaceExisting)
        {
            if (!TryMakeRoom(targetIndex, null, out _, out _))
            {
                return;
            }
        }
        else if (Tiles[targetIndex].HasApp)
        {
            Tiles[targetIndex].Clear();
        }

        Tiles[targetIndex].Apply(state);
    }

    private bool TryMakeRoom(int targetIndex, int? reservedIndex, out AppTileState? displacedState, out int? displacedDestination)
    {
        displacedState = null;
        displacedDestination = null;

        if (!IsValidIndex(targetIndex))
        {
            return false;
        }

        var targetTile = Tiles[targetIndex];
        if (!targetTile.HasApp)
        {
            return true;
        }

        int destination = FindNearestFreeSlot(targetIndex, reservedIndex);
        if (destination == -1)
        {
            MessageBox.Show(this, "No hay un lugar libre cercano para reubicar la aplicación.", "Sin espacio", MessageBoxButton.OK, MessageBoxImage.Information);
            return false;
        }

        displacedState = targetTile.CreateSnapshot();
        displacedDestination = destination;

        if (!reservedIndex.HasValue || destination != reservedIndex.Value)
        {
            Tiles[destination].Apply(displacedState.Value);
        }

        return true;
    }

    private AppTileState? ExtractAppState(IDataObject data)
    {
        foreach (var path in GetCandidatePaths(data))
        {
            if (TryCreateStateFromPath(path, out var state))
            {
                return state;
            }
        }

        return null;
    }

    private bool TryCreateStateFromPath(string path, out AppTileState state)
    {
        state = default;
        if (string.IsNullOrWhiteSpace(path))
        {
            return false;
        }

        try
        {
            if (path.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase))
            {
                var shortcut = ResolveShortcut(path);
                if (shortcut.HasValue)
                {
                    var shortcutInfo = shortcut.Value;
                    var icon = LoadIcon(shortcutInfo.IconLocation, shortcutInfo.TargetPath ?? shortcutInfo.CommandPath);
                    state = new AppTileState(
                        shortcutInfo.DisplayName,
                        shortcutInfo.CommandPath,
                        shortcutInfo.TargetPath ?? shortcutInfo.CommandPath,
                        shortcutInfo.Arguments,
                        shortcutInfo.WorkingDirectory,
                        icon);
                    return true;
                }
            }

            if (IsShellReference(path))
            {
                var shellState = CreateShellReferenceState(path);
                if (shellState.HasValue)
                {
                    state = shellState.Value;
                    return true;
                }
            }

            if (File.Exists(path))
            {
                string displayName = Path.GetFileNameWithoutExtension(path);
                string? workingDirectory = Path.GetDirectoryName(path);
                var icon = LoadIcon(path, path);
                state = new AppTileState(displayName, path, path, null, workingDirectory, icon);
                return true;
            }
        }
        catch
        {
            // Ignorar entradas inválidas
        }

        return false;
    }

    private IEnumerable<string> GetCandidatePaths(IDataObject data)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (data.GetDataPresent(DataFormats.FileDrop))
        {
            var fileDropData = data.GetData(DataFormats.FileDrop);
            if (fileDropData is string[] files)
            {
                foreach (var file in files)
                {
                    if (!string.IsNullOrWhiteSpace(file) && seen.Add(file))
                    {
                        yield return file;
                    }
                }
            }
            else if (fileDropData is StringCollection fileCollection)
            {
                foreach (string? entry in fileCollection)
                {
                    if (!string.IsNullOrWhiteSpace(entry) && seen.Add(entry))
                    {
                        yield return entry;
                    }
                }
            }
        }

        if (data.GetDataPresent("Shell IDList Array"))
        {
            foreach (var file in GetPathsFromShellIdList(data))
            {
                if (!string.IsNullOrWhiteSpace(file) && seen.Add(file))
                {
                    yield return file;
                }
            }
        }

        if (data.GetDataPresent(DataFormats.UnicodeText))
        {
            if (TryNormalizeStringCandidate(data.GetData(DataFormats.UnicodeText) as string, out var candidate) && seen.Add(candidate))
            {
                yield return candidate;
            }
        }

        if (data.GetDataPresent(DataFormats.StringFormat))
        {
            if (TryNormalizeStringCandidate(data.GetData(DataFormats.StringFormat) as string, out var candidate) && seen.Add(candidate))
            {
                yield return candidate;
            }
        }

        if (data.GetDataPresent(DataFormats.Text))
        {
            if (TryNormalizeStringCandidate(data.GetData(DataFormats.Text) as string, out var candidate) && seen.Add(candidate))
            {
                yield return candidate;
            }
        }
    }

    private static bool TryNormalizeStringCandidate(string? raw, out string normalized)
    {
        normalized = string.Empty;
        if (string.IsNullOrWhiteSpace(raw))
        {
            return false;
        }

        var text = raw.Trim().Trim('"').Trim('\0');
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        if (File.Exists(text) || IsShellReference(text))
        {
            normalized = text;
            return true;
        }

        return false;
    }

    private IEnumerable<string> GetPathsFromShellIdList(IDataObject data)
    {
        if (data.GetData("Shell IDList Array") is not MemoryStream stream)
        {
            yield break;
        }

        byte[] buffer = stream.ToArray();
        if (buffer.Length == 0)
        {
            yield break;
        }

        var handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
        try
        {
            IntPtr basePtr = handle.AddrOfPinnedObject();
            int count = Marshal.ReadInt32(basePtr);
            if (count <= 0)
            {
                yield break;
            }

            var offsets = new int[count + 1];
            for (int i = 0; i <= count; i++)
            {
                offsets[i] = Marshal.ReadInt32(basePtr, sizeof(int) * (i + 1));
            }

            IntPtr parentPidl = IntPtr.Add(basePtr, offsets[0]);

            for (int i = 1; i <= count; i++)
            {
                IntPtr itemPidl = IntPtr.Add(basePtr, offsets[i]);
                IntPtr fullPidl = ILCombine(parentPidl, itemPidl);
                if (fullPidl == IntPtr.Zero)
                {
                    continue;
                }

                try
                {
                    var pathBuilder = new StringBuilder(512);
                    bool pathFound = false;
                    if (SHGetPathFromIDList(fullPidl, pathBuilder))
                    {
                        var path = pathBuilder.ToString();
                        if (!string.IsNullOrWhiteSpace(path))
                        {
                            pathFound = true;
                            yield return path;
                        }
                    }

                    if (!pathFound && TryGetShellParsingName(fullPidl, out var parsingName))
                    {
                        yield return parsingName;
                    }
                }
                finally
                {
                    ILFree(fullPidl);
                }
            }
        }
        finally
        {
            handle.Free();
        }
    }

    private static bool TryGetShellParsingName(IntPtr pidl, out string parsingName)
    {
        parsingName = GetNameFromIdList(pidl, SIGDN.SIGDN_DESKTOPABSOLUTEPARSING) ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(parsingName))
        {
            return true;
        }

        parsingName = GetNameFromIdList(pidl, SIGDN.SIGDN_URL) ?? string.Empty;
        return !string.IsNullOrWhiteSpace(parsingName);
    }

    private bool HasExternalAppData(IDataObject data)
    {
        if (data.GetDataPresent(DataFormats.FileDrop))
        {
            return true;
        }

        if (data.GetDataPresent("Shell IDList Array"))
        {
            return true;
        }

        if (data.GetDataPresent(DataFormats.UnicodeText) && TryNormalizeStringCandidate(data.GetData(DataFormats.UnicodeText) as string, out _))
        {
            return true;
        }

        if (data.GetDataPresent(DataFormats.StringFormat) && TryNormalizeStringCandidate(data.GetData(DataFormats.StringFormat) as string, out _))
        {
            return true;
        }

        if (data.GetDataPresent(DataFormats.Text) && TryNormalizeStringCandidate(data.GetData(DataFormats.Text) as string, out _))
        {
            return true;
        }

        return false;
    }

    private static bool IsShellReference(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        var trimmed = value.Trim();
        return trimmed.StartsWith("shell:", StringComparison.OrdinalIgnoreCase) ||
               trimmed.StartsWith("::{", StringComparison.OrdinalIgnoreCase);
    }

    private static string GetFallbackNameFromShellReference(string reference)
    {
        if (string.IsNullOrWhiteSpace(reference))
        {
            return string.Empty;
        }

        var trimmed = reference.Trim().TrimEnd('\\');
        int separatorIndex = trimmed.LastIndexOf('\\');
        if (separatorIndex >= 0 && separatorIndex < trimmed.Length - 1)
        {
            return trimmed.Substring(separatorIndex + 1);
        }

        return trimmed;
    }

    private AppTileState? CreateShellReferenceState(string reference)
    {
        if (!IsShellReference(reference))
        {
            return null;
        }

        string explorerPath = Path.Combine(Environment.SystemDirectory, "explorer.exe");
        if (!File.Exists(explorerPath))
        {
            explorerPath = "explorer.exe";
        }

        if (TryGetShellItemInfo(reference, out var displayName, out var icon))
        {
            return new AppTileState(displayName, explorerPath, reference, reference, null, icon);
        }

        string fallbackName = GetFallbackNameFromShellReference(reference);
        return new AppTileState(string.IsNullOrWhiteSpace(fallbackName) ? reference : fallbackName, explorerPath, reference, reference, null, null);
    }

    private static bool TryGetShellItemInfo(string reference, out string displayName, out ImageSource? icon)
    {
        displayName = string.Empty;
        icon = null;

        if (SHParseDisplayName(reference, IntPtr.Zero, out var pidl, 0, out _) != 0 || pidl == IntPtr.Zero)
        {
            return false;
        }

        try
        {
            displayName = GetNameFromIdList(pidl, SIGDN.SIGDN_NORMALDISPLAY) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(displayName))
            {
                displayName = GetFallbackNameFromShellReference(reference);
            }

            icon = ExtractIconFromPidl(pidl);
            return true;
        }
        finally
        {
            ILFree(pidl);
        }
    }

    private static string? GetNameFromIdList(IntPtr pidl, SIGDN sigdn)
    {
        if (SHGetNameFromIDList(pidl, sigdn, out var namePtr) != 0 || namePtr == IntPtr.Zero)
        {
            return null;
        }

        try
        {
            return Marshal.PtrToStringUni(namePtr);
        }
        finally
        {
            CoTaskMemFree(namePtr);
        }
    }

    private static ImageSource? ExtractIconFromPidl(IntPtr pidl)
    {
        const uint SHGFI_PIDL = 0x000000008;
        const uint SHGFI_ICON = 0x000000100;
        const uint SHGFI_LARGEICON = 0x000000000;

        if (SHGetFileInfo(pidl, 0, out var info, (uint)Marshal.SizeOf<SHFILEINFO>(), SHGFI_PIDL | SHGFI_ICON | SHGFI_LARGEICON) == IntPtr.Zero || info.hIcon == IntPtr.Zero)
        {
            return null;
        }

        try
        {
            var source = Imaging.CreateBitmapSourceFromHIcon(info.hIcon, Int32Rect.Empty, BitmapSizeOptions.FromWidthAndHeight(64, 64));
            source.Freeze();
            return source;
        }
        finally
        {
            DestroyIcon(info.hIcon);
        }
    }

    private static T? FindAncestor<T>(DependencyObject? current) where T : DependencyObject
    {
        while (current is not null)
        {
            if (current is T matched)
            {
                return matched;
            }

            current = GetVisualOrLogicalParent(current);
        }

        return null;
    }

    private static DependencyObject? GetVisualOrLogicalParent(DependencyObject? child)
    {
        if (child is null)
        {
            return null;
        }

        if (child is FrameworkElement frameworkElement && frameworkElement.Parent is not null)
        {
            return frameworkElement.Parent;
        }

        if (child is FrameworkContentElement frameworkContentElement && frameworkContentElement.Parent is not null)
        {
            return frameworkContentElement.Parent;
        }

        return VisualTreeHelper.GetParent(child);
    }

    private int FindFirstFreeSlot()
    {
        for (int i = 0; i < Tiles.Count; i++)
        {
            if (!Tiles[i].HasApp)
            {
                return i;
            }
        }

        return -1;
    }

    private int FindNearestFreeSlot(int originIndex, int? reservedIndex)
    {
        int bestIndex = -1;
        int bestDistance = int.MaxValue;

        for (int i = 0; i < Tiles.Count; i++)
        {
            bool isReserved = reservedIndex.HasValue && i == reservedIndex.Value;
            bool isFree = !Tiles[i].HasApp || isReserved;
            if (!isFree || i == originIndex)
            {
                continue;
            }

            int distance = CalculateGridDistance(originIndex, i);
            if (distance < bestDistance || (distance == bestDistance && i < bestIndex))
            {
                bestDistance = distance;
                bestIndex = i;
            }
        }

        return bestIndex;
    }

    private static int CalculateGridDistance(int a, int b)
    {
        var (rowA, colA) = GetGridPosition(a);
        var (rowB, colB) = GetGridPosition(b);
        return Math.Abs(rowA - rowB) + Math.Abs(colA - colB);
    }

    private static (int Row, int Column) GetGridPosition(int index)
    {
        int row = index / Columns;
        int column = index % Columns;
        return (row, column);
    }

    private bool IsValidIndex(int index) => index >= 0 && index < Tiles.Count;

    private static ImageSource? LoadIcon(string? iconLocation, string? fallbackPath)
    {
        foreach (var candidate in GetIconCandidates(iconLocation, fallbackPath))
        {
            var image = ExtractIconImage(candidate.Path, candidate.Index);
            if (image is not null)
            {
                return image;
            }
        }

        return null;
    }

    private static IEnumerable<(string Path, int Index)> GetIconCandidates(string? iconLocation, string? fallbackPath)
    {
        if (!string.IsNullOrWhiteSpace(iconLocation))
        {
            var parts = iconLocation.Split(',');
            var iconPath = parts[0].Trim().Trim('"');
            int index = 0;
            if (parts.Length > 1 && int.TryParse(parts[1].Trim(), out var parsed))
            {
                index = parsed;
            }

            if (!string.IsNullOrWhiteSpace(iconPath))
            {
                yield return (iconPath, index);
            }
        }

        if (!string.IsNullOrWhiteSpace(fallbackPath))
        {
            yield return (fallbackPath, 0);
        }
    }

    private static ImageSource? ExtractIconImage(string path, int index)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
        {
            return null;
        }

        try
        {
            IntPtr[] smallIcons = new IntPtr[1];
            uint extracted = ExtractIconEx(path, index, null, smallIcons, 1);
            if (extracted > 0 && smallIcons[0] != IntPtr.Zero)
            {
                try
                {
                    var source = Imaging.CreateBitmapSourceFromHIcon(smallIcons[0], Int32Rect.Empty, BitmapSizeOptions.FromWidthAndHeight(64, 64));
                    source.Freeze();
                    return source;
                }
                finally
                {
                    DestroyIcon(smallIcons[0]);
                }
            }
        }
        catch
        {
            // Intentar método alternativo
        }

        try
        {
            using System.Drawing.Icon? icon = System.Drawing.Icon.ExtractAssociatedIcon(path);
            if (icon is not null)
            {
                var source = Imaging.CreateBitmapSourceFromHIcon(icon.Handle, Int32Rect.Empty, BitmapSizeOptions.FromWidthAndHeight(icon.Width, icon.Height));
                source.Freeze();
                return source;
            }
        }
        catch
        {
            // Ignorar
        }

        return null;
    }

    private ShortcutInfo? ResolveShortcut(string shortcutPath)
    {
        if (!File.Exists(shortcutPath))
        {
            return null;
        }

        object? shell = null;
        object? shortcutObj = null;

        try
        {
            var shellType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellType is null)
            {
                return null;
            }

            shell = Activator.CreateInstance(shellType);
            if (shell is null)
            {
                return null;
            }

            shortcutObj = shellType.InvokeMember("CreateShortcut", System.Reflection.BindingFlags.InvokeMethod, null, shell, new object[] { shortcutPath });
            if (shortcutObj is null)
            {
                return null;
            }

            dynamic shortcut = shortcutObj;
            string commandPath = shortcutPath;
            string targetPath = shortcut.TargetPath;
            string arguments = shortcut.Arguments;
            string workingDirectory = shortcut.WorkingDirectory;
            string iconLocation = shortcut.IconLocation;
            string description = shortcut.Description;

            return new ShortcutInfo(
                commandPath,
                string.IsNullOrWhiteSpace(targetPath) ? null : targetPath,
                string.IsNullOrWhiteSpace(arguments) ? null : arguments,
                string.IsNullOrWhiteSpace(workingDirectory) ? null : workingDirectory,
                string.IsNullOrWhiteSpace(iconLocation) ? null : iconLocation,
                !string.IsNullOrWhiteSpace(description) ? description : Path.GetFileNameWithoutExtension(shortcutPath));
        }
        catch
        {
            // Ignorar y usar fallback
        }
        finally
        {
            if (shortcutObj is not null && Marshal.IsComObject(shortcutObj))
            {
                Marshal.FinalReleaseComObject(shortcutObj);
            }

            if (shell is not null && Marshal.IsComObject(shell))
            {
                Marshal.FinalReleaseComObject(shell);
            }
        }

        return new ShortcutInfo(shortcutPath, null, null, null, null, Path.GetFileNameWithoutExtension(shortcutPath));
    }

    private readonly struct ShortcutInfo
    {
        public ShortcutInfo(string commandPath, string? targetPath, string? arguments, string? workingDirectory, string? iconLocation, string displayName)
        {
            CommandPath = commandPath;
            TargetPath = targetPath;
            Arguments = arguments;
            WorkingDirectory = workingDirectory;
            IconLocation = iconLocation;
            DisplayName = displayName;
        }

        public string CommandPath { get; }
        public string? TargetPath { get; }
        public string? Arguments { get; }
        public string? WorkingDirectory { get; }
        public string? IconLocation { get; }
        public string DisplayName { get; }
    }

    public readonly struct AppTileState
    {
        public AppTileState(string displayName, string commandPath, string? launchPath, string? arguments, string? workingDirectory, ImageSource? icon)
        {
            DisplayName = displayName;
            CommandPath = commandPath;
            LaunchPath = launchPath;
            Arguments = arguments;
            WorkingDirectory = workingDirectory;
            Icon = icon;
        }

        public string DisplayName { get; }
        public string CommandPath { get; }
        public string? LaunchPath { get; }
        public string? Arguments { get; }
        public string? WorkingDirectory { get; }
        public ImageSource? Icon { get; }

        public bool HasApp => !string.IsNullOrWhiteSpace(CommandPath);
    }

    public class AppTile : INotifyPropertyChanged
    {
        private string? _displayName;
        private string? _commandPath;
        private string? _launchPath;
        private string? _arguments;
        private string? _workingDirectory;
        private ImageSource? _icon;

        public Guid Id { get; } = Guid.NewGuid();

        public string? DisplayName
        {
            get => _displayName;
            private set => SetField(ref _displayName, value);
        }

        public string? CommandPath
        {
            get => _commandPath;
            private set => SetField(ref _commandPath, value, nameof(CommandPath), true);
        }

        public string? LaunchPath
        {
            get => _launchPath;
            private set => SetField(ref _launchPath, value);
        }

        public string? Arguments
        {
            get => _arguments;
            private set => SetField(ref _arguments, value);
        }

        public string? WorkingDirectory
        {
            get => _workingDirectory;
            private set => SetField(ref _workingDirectory, value);
        }

        public ImageSource? Icon
        {
            get => _icon;
            private set => SetField(ref _icon, value);
        }

        public bool HasApp => !string.IsNullOrWhiteSpace(CommandPath);

        public event PropertyChangedEventHandler? PropertyChanged;

        public void Apply(AppTileState state)
        {
            DisplayName = state.DisplayName;
            CommandPath = state.CommandPath;
            LaunchPath = state.LaunchPath;
            Arguments = state.Arguments;
            WorkingDirectory = state.WorkingDirectory;
            Icon = state.Icon;
            OnPropertyChanged(nameof(HasApp));
        }

        public void Clear()
        {
            DisplayName = null;
            CommandPath = null;
            LaunchPath = null;
            Arguments = null;
            WorkingDirectory = null;
            Icon = null;
            OnPropertyChanged(nameof(HasApp));
        }

        public AppTileState CreateSnapshot() => new AppTileState(DisplayName ?? string.Empty, CommandPath ?? string.Empty, LaunchPath, Arguments, WorkingDirectory, Icon);

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null, bool notifyHasApp = false)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
            {
                return;
            }

            field = value;
            if (!string.IsNullOrEmpty(propertyName))
            {
                OnPropertyChanged(propertyName);
            }
            if (notifyHasApp)
            {
                OnPropertyChanged(nameof(HasApp));
            }
        }
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern bool SHGetPathFromIDList(IntPtr pidl, StringBuilder pszPath);

    [DllImport("shell32.dll")]
    private static extern IntPtr ILCombine(IntPtr pidl1, IntPtr pidl2);

    [DllImport("shell32.dll")]
    private static extern void ILFree(IntPtr pidl);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHParseDisplayName(string name, IntPtr bindingContext, out IntPtr pidl, uint sfgaoIn, out uint psfgaoOut);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHGetNameFromIDList(IntPtr pidl, SIGDN sigdnName, out IntPtr ppszName);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SHGetFileInfo(IntPtr pszPath, uint dwFileAttributes, out SHFILEINFO psfi, uint cbFileInfo, uint uFlags);

    [DllImport("ole32.dll")]
    private static extern void CoTaskMemFree(IntPtr pv);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern uint ExtractIconEx(string lpszFile, int nIconIndex, IntPtr[]? phiconLarge, IntPtr[]? phiconSmall, uint nIcons);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool DestroyIcon(IntPtr hIcon);

    private enum SIGDN : uint
    {
        SIGDN_NORMALDISPLAY = 0x00000000,
        SIGDN_DESKTOPABSOLUTEPARSING = 0x80028000,
        SIGDN_DESKTOPABSOLUTEEDITING = 0x8004C000,
        SIGDN_FILESYSPATH = 0x80058000,
        SIGDN_URL = 0x80068000
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct SHFILEINFO
    {
        public IntPtr hIcon;
        public int iIcon;
        public uint dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
        public string szTypeName;
    }
}
