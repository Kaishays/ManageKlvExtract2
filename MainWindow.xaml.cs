using System.Diagnostics;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Threading;
using System.Linq;

namespace ManageKlvExtract;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private const string KLV_EXTRACTOR_PATH = @"S:\Projects\AltiCam Vision\Software (MSP & GUI)\KLV_Metadata_Extraction\Release2\KlvExtractor.exe";
    private readonly DispatcherTimer _monitoringTimer;
    private readonly PerformanceCounter _cpuCounter;
    private readonly PerformanceCounter _memoryCounter;
    private bool _countersInitialized = false;
    private const double LOW_USAGE_THRESHOLD = 60.0;
    private const double HIGH_USAGE_THRESHOLD = 85.0;
    private const int MIN_INSTANCES = 1;
    private const int LOW_MEMORY_MAX_INSTANCES = 5;
    private const int HIGH_MEMORY_MAX_INSTANCES = 13;
    private const int MEMORY_THRESHOLD_GB = 16;
    private readonly int MAX_INSTANCES;
    private bool CanStartMoreExtractions = true;

    public MainWindow()
    {
        InitializeComponent();
        
        MAX_INSTANCES = CalculateMaxInstancesBasedOnSystemMemory();
        
        _cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total");
        _memoryCounter = new PerformanceCounter("Memory", "Available MBytes");
        
        _monitoringTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(2)
        };
        _monitoringTimer.Tick += MonitoringTimer_Tick;
        
        Loaded += MainWindow_Loaded;
        Closing += MainWindow_Closing;
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        await InitializeApplication();
    }

    private int CalculateMaxInstancesBasedOnSystemMemory()
    {
        var totalMemoryMB = GetTotalPhysicalMemory();
        
        if (totalMemoryMB == 0)
        {
            return LOW_MEMORY_MAX_INSTANCES;
        }
        
        var totalMemoryGB = totalMemoryMB / 1024;
        
        return totalMemoryGB <= MEMORY_THRESHOLD_GB 
            ? LOW_MEMORY_MAX_INSTANCES 
            : HIGH_MEMORY_MAX_INSTANCES;
    }


    private async void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        CanStartMoreExtractions = false;
        // Cancel the close event temporarily to allow async cleanup
        e.Cancel = true;
        
        await TerminateAllExtractorProcessesOnShutdown();
        
        _monitoringTimer.Stop();
        _cpuCounter.Dispose();
        _memoryCounter.Dispose();
        
        // Now allow the window to close
        e.Cancel = false;
        Application.Current.Shutdown();
    }

    private async Task InitializeApplication()
    {
        UpdateStatus("Initializing application...");
        await ValidateExtractorPath();
        await InitializePerformanceCounters();
        await StartThreeExtractorInstances();
        StartSystemMonitoring();
        UpdateStatus("Application initialized successfully");
    }

    private async Task InitializePerformanceCounters()
    {
        UpdateStatus("Initializing performance counters...");
        
        await Task.Run(() =>
        {
            try
            {
                _cpuCounter.NextValue();
                _memoryCounter.NextValue();
                
                Thread.Sleep(1000);
                
                _cpuCounter.NextValue();
                _memoryCounter.NextValue();
                
                _countersInitialized = true;
                
                Dispatcher.Invoke(() => UpdateStatus("Performance counters initialized"));
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => UpdateStatus($"Warning: Performance counter initialization failed - {ex.Message}"));
            }
        });
    }

    private Task ValidateExtractorPath()
    {
        if (!File.Exists(KLV_EXTRACTOR_PATH))
        {
            var errorMessage = $"KlvExtractor.exe not found at: {KLV_EXTRACTOR_PATH}";
            UpdateStatus(errorMessage);
            MessageBox.Show(errorMessage, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            return Task.CompletedTask;
        }
        UpdateStatus("KlvExtractor.exe path validated");
        return Task.CompletedTask;
    }

    private async Task StartThreeExtractorInstances()
    {
        UpdateStatus("Starting 3 KLV Extractor instances...");
        
        for (int i = 1; i <= 3; i++)
        {
            await LaunchExtractorInstance(i);
            await Task.Delay(500);
        }
        
        await RefreshProcessInformation();
        UpdateStatus("All extractor instances started");
    }

    private Task LaunchExtractorInstance(int instanceNumber)
    {
        try
        {
            if (CanStartMoreExtractions)
            {
                var startInfo = CreateProcessStartInfo();
                var process = Process.Start(startInfo);

                if (process != null)
                {
                    UpdateStatus($"Started extractor instance {instanceNumber} (PID: {process.Id})");
                }
            }
        }
        catch (Exception ex)
        {
            var errorMessage = $"Failed to start extractor instance {instanceNumber}: {ex.Message}";
            UpdateStatus(errorMessage);
            MessageBox.Show(errorMessage, "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        return Task.CompletedTask;
    }

    private ProcessStartInfo CreateProcessStartInfo()
    {
        return new ProcessStartInfo
        {
            FileName = KLV_EXTRACTOR_PATH,
            UseShellExecute = true,
            CreateNoWindow = false
        };
    }

    private void StartSystemMonitoring()
    {
        UpdateSystemInformation();
        PerformInitialSystemResourceCheck();
        _monitoringTimer.Start();
        UpdateStatus("System monitoring started");
    }

    private void PerformInitialSystemResourceCheck()
    {
        var cpuUsage = GetCpuUsage();
        var memoryInfo = GetMemoryInformation();
        
        UpdateCpuDisplay(cpuUsage);
        UpdateMemoryDisplay(memoryInfo);
        
        if (_countersInitialized && (cpuUsage > 0 || memoryInfo.totalMemoryMB > 0))
        {
            UpdateStatus($"Performance monitoring active - CPU: {cpuUsage}%, Memory: {memoryInfo.usagePercentage}%");
        }
        else
        {
            UpdateStatus("Warning: Performance counters may not be functioning properly");
        }
    }

    private async Task StartOrStopKlvExtractor(double cpuUsage, double memoryUsagePercentage)
    {
        var resourceUsages = new double[] { cpuUsage, memoryUsagePercentage};
        
        if (ShouldStartNewExtractorInstance(resourceUsages))
        {
            await StartNewExtractorInstanceIfNeeded();
        }
        else if (ShouldTerminateExtractorInstance(resourceUsages))
        {
            await TerminateOldestExtractorInstance();
        }
    }

    private bool ShouldStartNewExtractorInstance(double[] resourceUsages)
    {
        var runningProcessCount = GetKlvExtractorProcesses().Count;
        
        return AreAllResourcesBelowThreshold(resourceUsages, LOW_USAGE_THRESHOLD) 
               && runningProcessCount < MAX_INSTANCES;
    }

    private bool ShouldTerminateExtractorInstance(double[] resourceUsages)
    {
        var runningProcessCount = GetKlvExtractorProcesses().Count;
        
        return IsAnyResourceAboveThreshold(resourceUsages, HIGH_USAGE_THRESHOLD) 
               && runningProcessCount > MIN_INSTANCES;
    }

    private bool AreAllResourcesBelowThreshold(double[] resourceUsages, double threshold)
    {
        return resourceUsages.All(usage => usage < threshold && usage > 0);
    }

    private bool IsAnyResourceAboveThreshold(double[] resourceUsages, double threshold)
    {
        return resourceUsages.Any(usage => usage > threshold);
    }

    private async Task StartNewExtractorInstanceIfNeeded()
    {
        var currentInstanceCount = GetKlvExtractorProcesses().Count;
        var nextInstanceNumber = currentInstanceCount + 1;
        
        UpdateStatus($"System resources available - starting extractor instance {nextInstanceNumber}");
        await LaunchExtractorInstance(nextInstanceNumber);
        await RefreshProcessInformation();
    }

    private async Task TerminateOldestExtractorInstance()
    {
        var runningProcesses = GetKlvExtractorProcesses();
        var oldestProcess = FindOldestProcess(runningProcesses);
        
        if (oldestProcess != null)
        {
            await TerminateSpecificProcess(oldestProcess);
        }
    }

    private Process? FindOldestProcess(List<Process> processes)
    {
        if (!processes.Any())
        {
            return null;
        }
        
        return processes.OrderBy(p => GetProcessStartTime(p)).FirstOrDefault();
    }

    private DateTime GetProcessStartTime(Process process)
    {
        try
        {
            return process.StartTime;
        }
        catch
        {
            return DateTime.MaxValue;
        }
    }

    private async Task TerminateSpecificProcess(Process process)
    {
        UpdateStatus($"High resource usage detected - terminating process {process.Id}");
        
        if (await AttemptGracefulProcessTermination(process))
        {
            await RefreshProcessInformation();
            UpdateStatus($"Process {process.Id} terminated successfully");
        }
    }

    private async void MonitoringTimer_Tick(object? sender, EventArgs e)
    {
        await RefreshProcessInformation();
        await UpdateSystemResourceUsage();
    }

    private Task RefreshProcessInformation()
    {
        var runningProcesses = GetKlvExtractorProcesses();
        
        Dispatcher.Invoke(() =>
        {
            ProcessCountLabel.Text = $"Running Processes: {runningProcesses.Count}";
            UpdateProcessList(runningProcesses);
        });
        
        return Task.CompletedTask;
    }

    private List<Process> GetKlvExtractorProcesses()
    {
        var extractorProcesses = new List<Process>();
        var klvProcesses = Process.GetProcessesByName("KlvExtractor");
        
        foreach (var process in klvProcesses)
        {
            try
            {
                if (IsKlvExtractorFromTargetPath(process))
                {
                    extractorProcesses.Add(process);
                }
            }
            catch (Exception)
            {
                continue;
            }
        }
        
        return extractorProcesses;
    }

    private bool IsKlvExtractorFromTargetPath(Process process)
    {
        try
        {
            var processPath = process.MainModule?.FileName;
            return string.Equals(processPath, KLV_EXTRACTOR_PATH, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private void UpdateProcessList(List<Process> processes)
    {
        ProcessListBox.Items.Clear();
        
        foreach (var process in processes)
        {
            try
            {
                var processInfo = CreateProcessDisplayInfo(process);
                ProcessListBox.Items.Add(processInfo);
            }
            catch (Exception ex)
            {
                ProcessListBox.Items.Add($"Process {process.Id}: Error accessing info - {ex.Message}");
            }
        }
    }

    private string CreateProcessDisplayInfo(Process process)
    {
        var memoryMB = process.WorkingSet64 / (1024 * 1024);
        var startTime = process.StartTime.ToString("HH:mm:ss");
        var cpuTime = process.TotalProcessorTime.TotalSeconds;
        
        return $"PID: {process.Id:D5} | Memory: {memoryMB:D4} MB | Started: {startTime} | CPU: {cpuTime:F1}s";
    }

    private async Task UpdateSystemResourceUsage()
    {
        await Task.Run(() =>
        {
            var cpuUsage = GetCpuUsage();
            var memoryInfo = GetMemoryInformation();
            
            Dispatcher.Invoke(() =>
            {
                UpdateCpuDisplay(cpuUsage);
                UpdateMemoryDisplay(memoryInfo);
            });
            
            Dispatcher.Invoke(async () =>
            {
                await StartOrStopKlvExtractor(cpuUsage, memoryInfo.usagePercentage);
            });
        });
    }

    private double GetCpuUsage()
    {
        if (!_countersInitialized)
        {
            return 0;
        }

        try
        {
            var cpuValue = _cpuCounter.NextValue();
            return Math.Round(cpuValue, 1);
        }
        catch (Exception ex)
        {
            UpdateStatus($"CPU counter error: {ex.Message}");
            return 0;
        }
    }

    private (long availableMemoryMB, long totalMemoryMB, double usagePercentage) GetMemoryInformation()
    {
        if (!_countersInitialized)
        {
            return (0, 0, 0);
        }

        try
        {
            var availableMemory = (long)_memoryCounter.NextValue();
            var totalMemory = GetTotalPhysicalMemory();
            
            if (totalMemory == 0)
            {
                return (availableMemory, 0, 0);
            }
            
            var usedMemory = totalMemory - availableMemory;
            var usagePercentage = (double)usedMemory / totalMemory * 100;
            
            return (availableMemory, totalMemory, Math.Round(usagePercentage, 1));
        }
        catch (Exception ex)
        {
            UpdateStatus($"Memory counter error: {ex.Message}");
            return (0, 0, 0);
        }
    }

    private long GetTotalPhysicalMemory()
    {
        try
        {
            using var searcher = new ManagementObjectSearcher("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem");
            using var results = searcher.Get();
            
            foreach (ManagementObject result in results)
            {
                var totalBytes = Convert.ToInt64(result["TotalPhysicalMemory"]);
                return totalBytes / (1024 * 1024);
            }
        }
        catch (Exception ex)
        {
            UpdateStatus($"WMI memory query failed, using fallback method: {ex.Message}");
        }
        
        try
        {
            var memoryStatus = new MemoryStatus();
            if (GlobalMemoryStatusEx(memoryStatus))
            {
                return (long)(memoryStatus.TotalPhys / (1024 * 1024));
            }
        }
        catch
        {
            // Fallback failed too
        }
        
        return 0;
    }

    [StructLayout(LayoutKind.Sequential)]
    private class MemoryStatus
    {
        public uint Length = 64;
        public uint MemoryLoad;
        public ulong TotalPhys;
        public ulong AvailPhys;
        public ulong TotalPageFile;
        public ulong AvailPageFile;
        public ulong TotalVirtual;
        public ulong AvailVirtual;
        public ulong AvailExtendedVirtual;
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GlobalMemoryStatusEx(MemoryStatus lpBuffer);



    private void UpdateCpuDisplay(double cpuUsage)
    {
        CpuProgressBar.Value = cpuUsage;
        CpuPercentageLabel.Text = $"{cpuUsage}%";
    }

    private void UpdateMemoryDisplay((long available, long total, double percentage) memoryInfo)
    {
        MemoryProgressBar.Value = memoryInfo.percentage;
        var usedMemory = memoryInfo.total - memoryInfo.available;
        MemoryUsageLabel.Text = $"{usedMemory:N0} MB / {memoryInfo.total:N0} MB ({memoryInfo.percentage}%)";
    }

    private void UpdateGpuDisplay(double gpuUsage)
    {
        GpuProgressBar.Value = gpuUsage;
        GpuPercentageLabel.Text = $"{gpuUsage}%";
    }

    private void UpdateSystemInformation()
    {
        var systemInfo = GatherSystemInformation();
        SystemInfoLabel.Text = systemInfo;
    }

    private string GatherSystemInformation()
    {
        var info = new StringBuilder();
        
        try
        {
            info.AppendLine($"OS: {Environment.OSVersion}");
            info.AppendLine($"Machine: {Environment.MachineName}");
            info.AppendLine($"Processors: {Environment.ProcessorCount}");
            info.AppendLine($".NET Version: {Environment.Version}");
        }
        catch (Exception ex)
        {
            info.AppendLine($"Error gathering system info: {ex.Message}");
        }
        
        return info.ToString();
    }

    private void UpdateStatus(string message)
    {
        Dispatcher.Invoke(() =>
        {
            StatusLabel.Text = $"{DateTime.Now:HH:mm:ss} - {message}";
        });
    }

    private async void StartExtractorsButton_Click(object sender, RoutedEventArgs e)
    {
        await StartThreeExtractorInstances();
    }

    private async void KillAllExtractorsButton_Click(object sender, RoutedEventArgs e)
    {
        await TerminateAllExtractorProcesses();
    }

    private async Task TerminateAllExtractorProcesses()
    {
        UpdateStatus("Terminating all KLV Extractor processes...");
        
        var processes = GetKlvExtractorProcesses();
        var terminatedCount = 0;
        
        foreach (var process in processes)
        {
            if (await AttemptGracefulProcessTermination(process))
            {
                terminatedCount++;
            }
        }
        
        await RefreshProcessInformation();
        UpdateStatus($"Terminated {terminatedCount} extractor processes");
    }

    private async Task TerminateAllExtractorProcessesOnShutdown()
    {
        UpdateStatus("Application closing - terminating all KLV Extractor processes...");
        
        var processes = GetKlvExtractorProcesses();
        
        if (!processes.Any())
        {
            UpdateStatus("No extractor processes to terminate");
            return;
        }
        
        var terminatedCount = 0;
        
        foreach (var process in processes)
        {
            if (await AttemptGracefulProcessTermination(process))
            {
                terminatedCount++;
            }
        }
        
        UpdateStatus($"Application shutdown - terminated {terminatedCount} extractor processes");
    }

    private async Task<bool> AttemptGracefulProcessTermination(Process process)
    {
        try
        {
            if (process.HasExited)
            {
                return true;
            }

            UpdateStatus($"Attempting graceful shutdown of process {process.Id}");
            
            if (await AttemptGracefulShutdown(process))
            {
                UpdateStatus($"Process {process.Id} closed gracefully");
                return true;
            }
            
            UpdateStatus($"Graceful shutdown failed for process {process.Id}, forcing termination");
            return await ForceProcessTermination(process);
        }
        catch (Exception ex)
        {
            UpdateStatus($"Failed to terminate process {process.Id}: {ex.Message}");
            return false;
        }
    }

    private async Task<bool> AttemptGracefulShutdown(Process process)
    {
        try
        {
            if (!process.CloseMainWindow())
            {
                return false;
            }
            
            const int gracefulShutdownTimeoutMs = 5000;
            var timeoutTask = Task.Delay(gracefulShutdownTimeoutMs);
            var exitTask = process.WaitForExitAsync();
            
            var completedTask = await Task.WhenAny(exitTask, timeoutTask);
            
            return completedTask == exitTask;
        }
        catch
        {
            return false;
        }
    }

    private async Task<bool> ForceProcessTermination(Process process)
    {
        try
        {
            process.Kill();
            await process.WaitForExitAsync();
            return true;
        }
        catch
        {
            return false;
        }
    }

    private async void RefreshButton_Click(object sender, RoutedEventArgs e)
    {
        UpdateStatus("Refreshing statistics...");
        await RefreshProcessInformation();
        await UpdateSystemResourceUsage();
        UpdateSystemInformation();
        UpdateStatus("Statistics refreshed");
    }
}