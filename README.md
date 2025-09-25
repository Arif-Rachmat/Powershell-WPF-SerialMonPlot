# WPF Serial Monitor & Plotter

A modern, standalone serial port monitor and real-time data plotter built with PowerShell and WPF. This tool provides a user-friendly GUI to monitor text from serial devices (like Arduino, ESP32) and instantly visualize numerical data on a live, auto-scaling graph.

The main interface showing the Serial Monitor and Plotter tabs.

## Features
- **Responsive UI**: Built with WPF to be fast and non-blocking.
- **Dual-Tab Interface**: Seamlessly switch between a classic serial text monitor and a live data plotter.
- **Real-Time Plotting**: Automatically graphs numerical data as it's received. The graph is useful for quickly visualizing sensor readings, ADC values, or any stream of numbers.
- **Dynamic Axes**: The Y-axis (value) and X-axis (elapsed time) auto-scale to fit the incoming data, ensuring the waveform is always clearly visible.
- **Live Typing Mode**: An interactive mode that sends each character to the device the instant it is typed.
- **Configurable Sending**: Send messages with selectable line endings (NL, CR, NL & CR, or none).
- **Timestamping**: Prepend timestamps (absolute or relative) to each incoming line for easy debugging and logging.
- **Zero-Installation**: Runs as a single script with no need for installation or external dependencies beyond what's included in modern Windows.

## Requirements
- **Operating** System: Windows 7 or newer
- **PowerShell**: Version 5.1 or higher (included by default in Windows 10/11)
- **.NET Framework**: Version 4.5 or newer (included by default in modern Windows versions)

## Quick Start: Using WebRequest
No installation is needed. You can run this application directly from your PowerShell terminal.

Open PowerShell and run the following command:
```powershell
irm bit.ly/SerialMonPlots | iex
```
This will download the script and execute it in memory. The GUI window will appear after a few moments.

## How to Use
1. **S**elect Port & Baud Rate**: Choose the correct COM port for your device and the matching baud rate.

2. **Connect**: Click the "Connect" button. The application will start listening for data.

3. **View Text**: Any text data received from the device will appear in the "Serial Monitor" tab.

4. **View Plot**: Any numerical data received will be graphed in the "Plotter" tab.
    - The plotter expects one number per line. For example, an Arduino sending `Serial.println(analogRead(A0));` will work perfectly
5. **Send Data**:
    - **Standard Mode**: Type a message in the text box, select a line ending, and click "Send" or press <kbd>Enter</kbd>.
    - **Live Typing Mode**: Check the "Live Typing" checkbox. Each character you type will be sent instantly and then cleared from the input box in a short time.
## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.