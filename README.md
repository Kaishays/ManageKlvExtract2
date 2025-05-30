# KLV Extractor Manager

A WPF application for managing KLV Extractor processes and monitoring system resources.

## Features

### Process Management
- **Automatic Startup**: Launches 3 instances of KlvExtractor.exe on application startup
- **Process Monitoring**: Real-time tracking of running KLV Extractor processes
- **Process Control**: Start/stop extractor instances with dedicated buttons
- **Path Validation**: Verifies KlvExtractor.exe exists at the specified path

### System Monitoring
- **CPU Usage**: Real-time CPU utilization display with visual progress bar
- **Memory Usage**: Physical memory consumption with available/total information
- **GPU Usage**: Graphics processing unit utilization monitoring
- **System Information**: Basic system details (OS, machine name, processor count)

### User Interface
- **Modern Dark Theme**: Professional dark UI with color-coded progress bars
- **Real-time Updates**: Automatic refresh every 2 seconds
- **Process Details**: Detailed information including PID, memory usage, start time, and CPU time
- **Status Updates**: Real-time status messages with timestamps

## Configuration

### KLV Extractor Path
The application is configured to launch KlvExtractor.exe from:
```
S:\Projects\AltiCam Vision\Software (MSP & GUI)\KLV_Metadata_Extraction\Release\KlvExtractor.exe
```

To modify this path, update the `KLV_EXTRACTOR_PATH` constant in `MainWindow.xaml.cs`.

## System Requirements

- **Operating System**: Windows 10/11
- **.NET Framework**: .NET 8.0 or later
- **Permissions**: Administrator rights may be required for process management and system monitoring
- **Dependencies**: 
  - System.Management NuGet package
  - Windows Performance Counters

## Usage

### Startup Behavior
1. Application validates the KlvExtractor.exe path
2. Automatically launches 3 instances of the extractor
3. Begins real-time system monitoring
4. Updates display every 2 seconds

### Manual Controls
- **Start 3 Extractors**: Launch additional extractor instances
- **Kill All Extractors**: Terminate all running KLV Extractor processes
- **Refresh Stats**: Manually update process and system information

### Process Information Display
For each running KLV Extractor process:
- Process ID (PID)
- Memory usage in MB
- Start time
- Total CPU time consumed

### System Resources Display
- CPU usage percentage with green progress bar
- Memory usage with blue progress bar showing used/total
- GPU usage with orange progress bar
- System information panel

## Building and Running

### Build the Application
```bash
cd ManageKlvExtract
dotnet build
```

### Run the Application
```bash
dotnet run
```

Or run the compiled executable:
```
bin\Debug\net8.0-windows\ManageKlvExtract.exe
```

## Architecture

### Key Components
- **MainWindow.xaml**: UI layout and styling
- **MainWindow.xaml.cs**: Application logic and system monitoring
- **Performance Counters**: CPU and memory monitoring
- **WMI Integration**: GPU monitoring and system information
- **Process Management**: KLV Extractor instance control

### Design Principles
- **Single Responsibility**: Each method handles one specific task
- **Clean Code**: Descriptive naming and minimal comments
- **Error Handling**: Graceful handling of system access exceptions
- **Resource Management**: Proper disposal of performance counters and timers

## Error Handling

The application includes robust error handling for:
- Missing KlvExtractor.exe file
- Process launch failures
- System monitoring access issues
- Performance counter initialization problems

## Notes

- GPU monitoring may not be available on all systems
- Administrator privileges may be required for full functionality
- The application automatically validates the KLV Extractor path on startup
- System monitoring continues running until the application is closed 