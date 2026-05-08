using System.Collections.Concurrent;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;

namespace OfficeToPDF;

public partial class MainWindow : Window
{
    private bool _isRunning;
    private volatile bool _isCancelled;
    private int _totalJobs;
    private int _finishedTasks;
    private int _errorCount;
    private DateTime _startTime;
    private ConcurrentQueue<string>? _fileQueue;
    private List<Thread>? _workerThreads;
    private readonly string _configPath;
    private double _progressBarMaxWidth;

    public MainWindow()
    {
        InitializeComponent();

        _configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.ini");
        LoadConfig();

        Loaded += (_, _) =>
        {
            _progressBarMaxWidth = ProgressFill.Parent is Border parent ? parent.ActualWidth : ActualWidth - 56;
        };

        SizeChanged += (_, _) =>
        {
            if (ProgressFill.Parent is Border parent)
            {
                _progressBarMaxWidth = parent.ActualWidth;
                UpdateProgressBar();
            }
        };
    }

    private void LoadConfig()
    {
        if (!File.Exists(_configPath)) return;
        var lines = File.ReadAllLines(_configPath);
        foreach (var line in lines)
        {
            if (line.StartsWith("input=", StringComparison.OrdinalIgnoreCase))
            {
                var path = line["input=".Length..].Trim();
                if (Directory.Exists(path))
                {
                    InputPathBox.Text = path;
                    UpdateOutputPath();
                }
            }
        }
    }

    private void SaveConfig(string inputPath)
    {
        File.WriteAllText(_configPath, $"input={inputPath}");
    }

    // ═══ 拖曳 ═══
    private void Window_Drop(object sender, DragEventArgs e)
    {
        ResetDropVisual();
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

        var paths = (string[])e.Data.GetData(DataFormats.FileDrop)!;
        var folder = paths.FirstOrDefault(Directory.Exists);
        if (folder != null)
        {
            InputPathBox.Text = folder;
            SaveConfig(folder);
        }
        else
        {
            MessageBox.Show("請拖入資料夾，而非檔案", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private void Window_DragEnter(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            DropBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(0x4A, 0xDE, 0x80));
            DropBorder.Background = new SolidColorBrush(Color.FromRgb(0x14, 0x53, 0x2D));
        }
    }

    private void Window_DragLeave(object sender, DragEventArgs e)
    {
        ResetDropVisual();
    }

    private void ResetDropVisual()
    {
        DropBorder.BorderBrush = (SolidColorBrush)FindResource("BorderBrush");
        DropBorder.Background = (SolidColorBrush)FindResource("DropBgBrush");
    }

    private void DropArea_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        BrowseForFolder();
    }

    // ═══ 路徑選擇 ═══
    private void BrowseInput_Click(object sender, RoutedEventArgs e)
    {
        BrowseForFolder();
    }

    private void BrowseForFolder()
    {
        var dialog = new OpenFolderDialog
        {
            Title = "選擇輸入資料夾"
        };

        if (!string.IsNullOrEmpty(InputPathBox.Text) && Directory.Exists(InputPathBox.Text))
            dialog.InitialDirectory = InputPathBox.Text;

        if (dialog.ShowDialog() == true)
        {
            InputPathBox.Text = dialog.FolderName;
            SaveConfig(dialog.FolderName);
        }
    }

    private void InputPathBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        UpdateOutputPath();
    }

    private void UpdateOutputPath()
    {
        var input = InputPathBox.Text.Trim();
        if (string.IsNullOrEmpty(input) || !Directory.Exists(input))
        {
            OutputPathBox.Text = "";
            return;
        }
        OutputPathBox.Text = GetOutputFolder(input);
    }

    private static string GetOutputFolder(string inputFolder)
    {
        var parent = Path.GetDirectoryName(inputFolder) ?? inputFolder;
        var folderName = Path.GetFileName(inputFolder);
        var baseOutput = Path.Combine(parent, folderName + "-pdf");

        if (!Directory.Exists(baseOutput) || !Directory.EnumerateFileSystemEntries(baseOutput).Any())
            return baseOutput;

        var counter = 1;
        while (true)
        {
            var numbered = Path.Combine(parent, $"{folderName}-pdf-{counter}");
            if (!Directory.Exists(numbered) || !Directory.EnumerateFileSystemEntries(numbered).Any())
                return numbered;
            counter++;
        }
    }

    // ═══ 執行轉換 ═══
    private void Execute_Click(object sender, RoutedEventArgs e)
    {
        if (!_isRunning)
            StartConversion();
        else
            StopConversion();
    }

    private void StartConversion()
    {
        var inputFolder = InputPathBox.Text.Trim();
        if (string.IsNullOrEmpty(inputFolder) || !Directory.Exists(inputFolder))
        {
            MessageBox.Show("請先選擇有效的輸入資料夾", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var outputFolder = GetOutputFolder(inputFolder);
        OutputPathBox.Text = outputFolder;

        // 收集檔案
        var files = CollectFiles(inputFolder, outputFolder);
        if (files.Count == 0)
        {
            MessageBox.Show("輸入資料夾中沒有可轉換的檔案，或未勾選任何格式", "提示",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        _fileQueue = new ConcurrentQueue<string>(files.Select(f => $"{f.Input}|{f.Output}"));
        _totalJobs = files.Count;
        _finishedTasks = 0;
        _errorCount = 0;
        _isCancelled = false;
        _isRunning = true;
        _startTime = DateTime.Now;

        // UI 更新
        ExecuteBtn.Content = "\u25A0  停止";
        ExecuteBtn.Style = (Style)FindResource("StopButton");
        BrowseButton.IsEnabled = false;
        InputPathBox.IsReadOnly = true;
        LogBox.Text = "";
        SetProgress(0);
        FileCountText.Text = $"共 {_totalJobs} 個檔案";

        // 取得執行緒數
        var threadCount = (ThreadCombo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "1";
        var threads = int.Parse(threadCount);

        // 啟動工作執行緒 — 每個執行緒會重複使用同一個 Office 實例
        _workerThreads = new List<Thread>();
        for (var i = 0; i < threads; i++)
        {
            var t = new Thread(ConvertWorkerPooled) { IsBackground = true };
            t.Start();
            _workerThreads.Add(t);
        }

        // 進度監控
        var monitor = new Thread(MonitorProgress) { IsBackground = true };
        monitor.Start();
    }

    private void StopConversion()
    {
        _isCancelled = true;
        AppendLog("停止中...等待剩餘任務完成...");
    }

    private HashSet<string> GetEnabledExtensions()
    {
        var exts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (ChkDoc.IsChecked == true) exts.Add(".doc");
        if (ChkDocx.IsChecked == true) exts.Add(".docx");
        if (ChkXls.IsChecked == true) exts.Add(".xls");
        if (ChkXlsx.IsChecked == true) exts.Add(".xlsx");
        if (ChkPptx.IsChecked == true) exts.Add(".pptx");
        return exts;
    }

    private List<(string Input, string Output)> CollectFiles(string inputFolder, string outputFolder)
    {
        var result = new List<(string, string)>();
        var extensions = GetEnabledExtensions();

        if (extensions.Count == 0)
            return result;

        foreach (var file in Directory.EnumerateFiles(inputFolder, "*.*", SearchOption.AllDirectories))
        {
            var fileName = Path.GetFileName(file);
            if (fileName.StartsWith('.') || fileName.StartsWith('~'))
                continue;

            var ext = Path.GetExtension(file).ToLowerInvariant();
            if (!extensions.Contains(ext))
                continue;

            var relativePath = Path.GetRelativePath(inputFolder, file);
            var outputFile = Path.Combine(outputFolder,
                Path.ChangeExtension(relativePath, ".pdf"));

            var outputDir = Path.GetDirectoryName(outputFile)!;
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            result.Add((file, outputFile));
        }

        return result;
    }

    // ═══ COM 轉換 Worker (pooled) ═══
    // 每個執行緒保持 Office 實例存活，遇到同類型檔案直接重用，省去重複啟動開銷
    private void ConvertWorkerPooled()
    {
        dynamic? wordApp = null;
        dynamic? excelApp = null;
        dynamic? pptApp = null;

        try
        {
            while (!_isCancelled && _fileQueue != null && _fileQueue.TryDequeue(out var item))
            {
                var parts = item.Split('|', 2);
                var inputFile = parts[0];
                var outputFile = parts[1];
                var ext = Path.GetExtension(inputFile).ToLowerInvariant();

                try
                {
                    var timestamp = DateTime.Now.ToString("HH:mm:ss");
                    AppendLog($"[{timestamp}] {Path.GetFileName(inputFile)} -> pdf");

                    if (ext is ".xls" or ".xlsx")
                    {
                        if (excelApp == null)
                        {
                            var t = Type.GetTypeFromProgID("Excel.Application")
                                ?? throw new InvalidOperationException("Excel is not installed");
                            excelApp = Activator.CreateInstance(t);
                            excelApp!.Visible = false;
                            excelApp.DisplayAlerts = false;
                        }
                        var workbook = excelApp.Workbooks.Open(inputFile);
                        workbook.ExportAsFixedFormat(0, outputFile); // xlTypePDF = 0
                        workbook.Close(false);
                    }
                    else if (ext is ".pptx")
                    {
                        if (pptApp == null)
                        {
                            var t = Type.GetTypeFromProgID("PowerPoint.Application")
                                ?? throw new InvalidOperationException("PowerPoint is not installed");
                            pptApp = Activator.CreateInstance(t);
                            pptApp!.DisplayAlerts = 0; // ppAlertsNone
                        }
                        var presentation = pptApp.Presentations.Open(inputFile, ReadOnly: true, Untitled: false, WithWindow: false);
                        presentation.SaveAs(outputFile, 32); // ppSaveAsPDF = 32
                        presentation.Close();
                    }
                    else // .doc, .docx
                    {
                        if (wordApp == null)
                        {
                            var t = Type.GetTypeFromProgID("Word.Application")
                                ?? throw new InvalidOperationException("Word is not installed");
                            wordApp = Activator.CreateInstance(t);
                            wordApp!.Visible = false;
                            wordApp.DisplayAlerts = 0; // wdAlertsNone
                        }
                        var doc = wordApp.Documents.Open(inputFile);
                        doc.SaveAs2(outputFile, 17); // wdFormatPDF = 17
                        doc.Close(false);
                    }
                }
                catch (Exception ex)
                {
                    Interlocked.Increment(ref _errorCount);
                    AppendLog($"[錯誤] {Path.GetFileName(inputFile)}: {ex.Message}");

                    // COM 出錯後實例可能已損壞，丟棄讓下次重建
                    if (ext is ".xls" or ".xlsx") { SafeQuit(ref excelApp); }
                    else if (ext is ".pptx") { SafeQuit(ref pptApp); }
                    else { SafeQuit(ref wordApp); }
                }
                finally
                {
                    Interlocked.Increment(ref _finishedTasks);
                }
            }
        }
        finally
        {
            // 整批完成後才關閉 Office
            SafeQuit(ref wordApp);
            SafeQuit(ref excelApp);
            SafeQuit(ref pptApp);
        }
    }

    private static void SafeQuit(ref dynamic? app)
    {
        if (app == null) return;
        try { app.Quit(); } catch { }
        try { Marshal.ReleaseComObject(app); } catch { }
        app = null;
    }

    // ═══ 進度監控 ═══
    private void MonitorProgress()
    {
        while (_isRunning)
        {
            Thread.Sleep(500);

            var finished = _finishedTasks;
            var errors = _errorCount;
            var allDone = finished >= _totalJobs;

            Dispatcher.Invoke(() =>
            {
                StatusText.Text = allDone
                    ? $"完成 ({finished}/{_totalJobs})"
                    : $"執行中 ({finished}/{_totalJobs})";
                ErrorText.Text = errors > 0 ? $"錯誤: {errors}" : "";
                SetProgress(allDone ? 1.0 : (double)finished / _totalJobs);
                TimeText.Text = $"耗時: {DateTime.Now - _startTime:hh\\:mm\\:ss}";
            });

            if (allDone || (_isCancelled && AllWorkersDone()))
            {
                Dispatcher.Invoke(() =>
                {
                    AppendLog(_isCancelled ? "任務已停止" : "轉換完成！");
                    ResetUIState();
                });
                break;
            }
        }
    }

    private bool AllWorkersDone()
    {
        return _workerThreads == null || _workerThreads.All(t => !t.IsAlive);
    }

    // ═══ UI Helpers ═══
    private void SetProgress(double ratio)
    {
        if (ProgressFill.Parent is Border parent)
            _progressBarMaxWidth = parent.ActualWidth;

        ProgressFill.Width = Math.Max(0, _progressBarMaxWidth * Math.Clamp(ratio, 0, 1));
    }

    private void UpdateProgressBar()
    {
        if (!_isRunning || _totalJobs == 0) return;
        var ratio = (double)_finishedTasks / _totalJobs;
        SetProgress(ratio);
    }

    private void AppendLog(string message)
    {
        Dispatcher.Invoke(() =>
        {
            LogBox.AppendText(message + Environment.NewLine);
            LogBox.ScrollToEnd();
        });
    }

    private void ResetUIState()
    {
        _isRunning = false;
        ExecuteBtn.Content = "\u25B6  開始轉換";
        ExecuteBtn.Style = (Style)FindResource("ExecuteButton");
        BrowseButton.IsEnabled = true;
        InputPathBox.IsReadOnly = false;
    }

    // ═══ 關閉視窗 ═══
    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        if (!_isRunning) return;

        var result = MessageBox.Show("還有任務在執行中，確定要關閉嗎？", "警告",
            MessageBoxButton.YesNo, MessageBoxImage.Warning);

        if (result == MessageBoxResult.No)
        {
            e.Cancel = true;
            return;
        }

        _isCancelled = true;
        // 等待 worker 結束
        if (_workerThreads != null)
        {
            foreach (var t in _workerThreads)
                t.Join(3000);
        }
    }
}
