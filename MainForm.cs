using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace BalStanokLogger
{
    public class MainForm : Form
    {
        private readonly ComboBox cbPort = new() { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly ComboBox cbBaud = new() { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly ComboBox cbDataBits = new() { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly ComboBox cbParity = new() { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly ComboBox cbStopBits = new() { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly ComboBox cbHandshake = new() { DropDownStyle = ComboBoxStyle.DropDownList };

        private readonly Button btnRefresh = new() { Text = "Refresh Ports" };
        private readonly Button btnTryCommon = new() { Text = "Try Common Settings" };
        private readonly Button btnConnect = new() { Text = "Connect" };
        private readonly Button btnDisconnect = new() { Text = "Disconnect", Enabled = false };
        private readonly Button btnClear = new() { Text = "Clear" };
        private readonly Button btnOpenLogFolder = new() { Text = "Open Logs Folder" };
        private readonly Button btnSupportPack = new() { Text = "Create Support Pack (.zip)" };

        private readonly TextBox tbHex = new() { Multiline = true, ScrollBars = ScrollBars.Both, WordWrap = false, ReadOnly = true };
        private readonly TextBox tbAscii = new() { Multiline = true, ScrollBars = ScrollBars.Both, WordWrap = false, ReadOnly = true };

        private readonly Label lblStatus = new() { AutoSize = true, Text = "Status: Disconnected" };
        private readonly Label lblBytes = new() { AutoSize = true, Text = "RX bytes: 0" };

        private SerialPort? _port;
        private long _rxBytes = 0;

        private readonly object _logLock = new();
        private string _sessionDir = "";
        private StreamWriter? _hexLog;
        private StreamWriter? _asciiLog;

        private readonly List<byte> _asciiBuf = new();

        public MainForm()
        {
            Text = "BalStanok Logger (COM Sniffer + Support Pack)";
            Width = 1200;
            Height = 760;

            var top = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 120,
                ColumnCount = 12,
                RowCount = 2,
                Padding = new Padding(8),
                AutoSize = false
            };

            for (int i = 0; i < top.ColumnCount; i++)
                top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / top.ColumnCount));
            top.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            top.RowStyles.Add(new RowStyle(SizeType.Percent, 50));

            cbBaud.Items.AddRange(new object[] { "9600", "19200", "38400", "57600", "115200" });
            cbBaud.SelectedItem = "9600";

            cbDataBits.Items.AddRange(new object[] { "7", "8" });
            cbDataBits.SelectedItem = "8";

            cbParity.Items.AddRange(Enum.GetNames(typeof(Parity)));
            cbParity.SelectedItem = Parity.None.ToString();

            cbStopBits.Items.AddRange(Enum.GetNames(typeof(StopBits)));
            cbStopBits.SelectedItem = StopBits.One.ToString();

            cbHandshake.Items.AddRange(Enum.GetNames(typeof(Handshake)));
            cbHandshake.SelectedItem = Handshake.None.ToString();

            AddLabeled(top, 0, 0, "Port", cbPort);
            AddLabeled(top, 2, 0, "Baud", cbBaud);
            AddLabeled(top, 4, 0, "Data Bits", cbDataBits);
            AddLabeled(top, 6, 0, "Parity", cbParity);
            AddLabeled(top, 8, 0, "Stop Bits", cbStopBits);
            AddLabeled(top, 10, 0, "Handshake", cbHandshake);

            top.Controls.Add(btnRefresh, 0, 1);
            top.SetColumnSpan(btnRefresh, 2);
            top.Controls.Add(btnTryCommon, 2, 1);
            top.SetColumnSpan(btnTryCommon, 2);

            top.Controls.Add(btnConnect, 4, 1);
            top.SetColumnSpan(btnConnect, 2);
            top.Controls.Add(btnDisconnect, 6, 1);
            top.SetColumnSpan(btnDisconnect, 2);

            top.Controls.Add(btnClear, 8, 1);
            top.SetColumnSpan(btnClear, 1);

            top.Controls.Add(btnOpenLogFolder, 9, 1);
            top.SetColumnSpan(btnOpenLogFolder, 2);

            top.Controls.Add(btnSupportPack, 11, 1);

            var statusPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 30,
                Padding = new Padding(8),
                FlowDirection = FlowDirection.LeftToRight
            };
            statusPanel.Controls.Add(lblStatus);
            statusPanel.Controls.Add(new Label { Text = "   " });
            statusPanel.Controls.Add(lblBytes);

            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 580
            };

            var leftGroup = new GroupBox { Text = "HEX (raw RX)", Dock = DockStyle.Fill };
            leftGroup.Controls.Add(tbHex);
            tbHex.Dock = DockStyle.Fill;

            var rightGroup = new GroupBox { Text = "ASCII (best-effort)", Dock = DockStyle.Fill };
            rightGroup.Controls.Add(tbAscii);
            tbAscii.Dock = DockStyle.Fill;

            split.Panel1.Controls.Add(leftGroup);
            split.Panel2.Controls.Add(rightGroup);

            Controls.Add(split);
            Controls.Add(statusPanel);
            Controls.Add(top);

            btnRefresh.Click += (_, __) => RefreshPorts();
            btnTryCommon.Click += (_, __) => TryCommonSettings();
            btnConnect.Click += (_, __) => Connect();
            btnDisconnect.Click += (_, __) => Disconnect();
            btnClear.Click += (_, __) => { tbHex.Clear(); tbAscii.Clear(); };
            btnOpenLogFolder.Click += (_, __) => OpenLogFolder();
            btnSupportPack.Click += (_, __) => CreateSupportPack();

            FormClosing += (_, __) => Disconnect();

            RefreshPorts();
            EnsureSessionDir();
        }

        private void AddLabeled(TableLayoutPanel p, int col, int row, string label, Control ctrl)
        {
            var lbl = new Label { Text = label, Dock = DockStyle.Fill, TextAlign = System.Drawing.ContentAlignment.MiddleLeft };
            p.Controls.Add(lbl, col, row);
            p.Controls.Add(ctrl, col + 1, row);
            ctrl.Dock = DockStyle.Fill;
        }

        private void RefreshPorts()
        {
            var ports = SerialPort.GetPortNames().OrderBy(x => x).ToArray();
            cbPort.Items.Clear();
            foreach (var port in ports) cbPort.Items.Add(port);

            if (ports.Length == 0)
            {
                cbPort.Text = "";
                return;
            }

            var prefer = ports.Contains("COM3") ? "COM3" : ports[0];
            cbPort.SelectedItem = prefer;
        }

        private void TryCommonSettings()
        {
            cbBaud.SelectedItem = "9600";
            cbDataBits.SelectedItem = "8";
            cbParity.SelectedItem = Parity.None.ToString();
            cbStopBits.SelectedItem = StopBits.One.ToString();
            cbHandshake.SelectedItem = Handshake.None.ToString();
        }

        private void EnsureSessionDir()
        {
            var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "BalStanokLoggerLogs");
            Directory.CreateDirectory(root);
            _sessionDir = Path.Combine(root, DateTime.Now.ToString("yyyyMMdd_HHmmss"));
            Directory.CreateDirectory(_sessionDir);
        }

        private void OpenLogFolder()
        {
            var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "BalStanokLoggerLogs");
            Directory.CreateDirectory(root);
            Process.Start(new ProcessStartInfo("explorer.exe", root) { UseShellExecute = true });
        }

        private void Connect()
        {
            if (_port != null) return;
            var portName = cbPort.SelectedItem?.ToString();
            if (string.IsNullOrWhiteSpace(portName))
            {
                MessageBox.Show("No COM port selected.");
                return;
            }

            EnsureSessionDir();

            var baud = int.Parse(cbBaud.SelectedItem?.ToString() ?? "9600");
            var dataBits = int.Parse(cbDataBits.SelectedItem?.ToString() ?? "8");
            var parity = (Parity)Enum.Parse(typeof(Parity), cbParity.SelectedItem?.ToString() ?? Parity.None.ToString());
            var stopBits = (StopBits)Enum.Parse(typeof(StopBits), cbStopBits.SelectedItem?.ToString() ?? StopBits.One.ToString());
            var handshake = (Handshake)Enum.Parse(typeof(Handshake), cbHandshake.SelectedItem?.ToString() ?? Handshake.None.ToString());

            try
            {
                _port = new SerialPort(portName, baud, parity, dataBits, stopBits)
                {
                    Handshake = handshake,
                    ReadTimeout = 500,
                   WriteTimeout = 500,
                    DtrEnable = true,
                    RtsEnable = true,
                    Encoding = Encoding.ASCII
                };

                _port.DataReceived += PortOnDataReceived;
                _port.Open();

                _rxBytes = 0;
                lblBytes.Text = "RX bytes: 0";
                lblStatus.Text = $"Status: Connected to {portName} ({baud} {dataBits}{parity.ToString()[0]}{(stopBits == StopBits.One ? "1" : stopBits.ToString())}, {handshake})";

                btnConnect.Enabled = false;
                btnDisconnect.Enabled = true;

                lock (_logLock)
                {
                    _hexLog = new StreamWriter(Path.Combine(_sessionDir, "serial_raw_hex.log"), append: true, Encoding.UTF8) { AutoFlush = true };
                    _asciiLog = new StreamWriter(Path.Combine(_sessionDir, "serial_ascii.log"), append: true, Encoding.UTF8) { AutoFlush = true };
                    _hexLog.WriteLine($"# START {DateTime.Now:O} Port={portName} Baud={baud} DataBits={dataBits} Parity={parity} StopBits={stopBits} Handshake={handshake}");
                    _asciiLog.WriteLine($"# START {DateTime.Now:O} Port={portName} Baud={baud} DataBits={dataBits} Parity={parity} StopBits={stopBits} Handshake={handshake}");
                }
            }
            catch (Exception ex)
            {
                Disconnect();
                MessageBox.Show($"Failed to open {portName}: {ex.Message}");
            }
        }

        private void Disconnect()
        {
            try
            {
                if (_port != null)
                {
                    _port.DataReceived -= PortOnDataReceived;
                    if (_port.IsOpen) _port.Close();
                    _port.Dispose();
                }
            }
            catch { }
            finally
            {
                _port = null;
                lblStatus.Text = "Status: Disconnected";
                btnConnect.Enabled = true;
                btnDisconnect.Enabled = false;

                lock (_logLock)
                {
                    _hexLog?.Dispose();
                    _asciiLog?.Dispose();
                    _hexLog = null;
                    _asciiLog = null;
                }
            }
        }

        private void PortOnDataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            if (_port == null) return;

            try
            {
                int bytesToRead = _port.BytesToRead;
                if (bytesToRead <= 0) return;

                byte[] buffer = new byte[bytesToRead];
                int read = _port.Read(buffer, 0, buffer.Length);
                if (read <= 0) return;

                Interlocked.Add(ref _rxBytes, read);
                var ts = DateTime.Now.ToString("O");

                var hex = BitConverter.ToString(buffer, 0, read).Replace("-", " ");
                lock (_logLock) { _hexLog?.WriteLine($"{ts}  {hex}"); }

                AppendAscii(buffer, read, ts);

                BeginInvoke(new Action(() =>
                {
                    lblBytes.Text = $"RX bytes: {_rxBytes}";
                    tbHex.AppendText($"{ts}  {hex}{Environment.NewLine}");
                    TruncateTextBox(tbHex, 20000);
                    TruncateTextBox(tbAscii, 20000);
                }));
            }
            catch (Exception ex)
            {
                BeginInvoke(new Action(() =>
                {
                    tbAscii.AppendText($"[{DateTime.Now:O}] ERROR reading serial: {ex.Message}{Environment.NewLine}");
                }));
            }
        }

        private void AppendAscii(byte[] buffer, int read, string ts)
        {
            for (int i = 0; i < read; i++) _asciiBuf.Add(buffer[i]);

            int lfIndex;
            while ((lfIndex = _asciiBuf.IndexOf((byte)'\n')) >= 0)
            {
                var lineBytes = _asciiBuf.Take(lfIndex + 1).ToArray();
                _asciiBuf.RemoveRange(0, lfIndex + 1);

                var line = ToPrintableAscii(lineBytes);

                lock (_logLock) { _asciiLog?.WriteLine($"{ts}  {line}"); }

                BeginInvoke(new Action(() =>
                {
                    tbAscii.AppendText($"{ts}  {line}{Environment.NewLine}");
                }));
            }

            if (_asciiBuf.Count > 512)
            {
                var chunk = _asciiBuf.Take(512).ToArray();
                _asciiBuf.RemoveRange(0, 512);

                var line = ToPrintableAscii(chunk);

                lock (_logLock) { _asciiLog?.WriteLine($"{ts}  {line}"); }

                BeginInvoke(new Action(() =>
                {
                    tbAscii.AppendText($"{ts}  {line}{Environment.NewLine}");
                }));
            }
        }

        private static string ToPrintableAscii(byte[] bytes)
        {
            var sb = new StringBuilder(bytes.Length);
            foreach (var b in bytes)
            {
                if (b == '\r' || b == '\n') continue;
                if (b >= 32 && b <= 126) sb.Append((char)b);
                else sb.Append('.');
            }
            return sb.ToString();
        }

        private static void TruncateTextBox(TextBox tb, int maxLines)
        {
            var lines = tb.Lines;
            if (lines.Length <= maxLines) return;
            tb.Lines = lines.Skip(lines.Length - maxLines).ToArray();
        }

        private void CreateSupportPack()
        {
            try
            {
                using var fbd = new FolderBrowserDialog
                {
                    Description = "Select BalStanok folder (e.g., Documents\\TatKardan-4\\TatKardanV18N1-01)",
                    UseDescriptionForTitle = true
                };
                if (fbd.ShowDialog() != DialogResult.OK) return;

                var balFolder = fbd.SelectedPath;

                var outRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "BalStanokLoggerLogs");
                Directory.CreateDirectory(outRoot);

                var zipPath = Path.Combine(outRoot, $"support_pack_{DateTime.Now:yyyyMMdd_HHmmss}.zip");

                var sysInfoPath = Path.Combine(_sessionDir, "system_com_info.txt");
                File.WriteAllText(sysInfoPath, GetComPortInfo(), Encoding.UTF8);

                if (File.Exists(zipPath)) File.Delete(zipPath);
                using (var zip = ZipFile.Open(zipPath, ZipArchiveMode.Create))
                {
                    AddFolderToZip(zip, balFolder, "BalStanokFolder");
                    AddFolderToZip(zip, _sessionDir, "LoggerSession");
                }

                MessageBox.Show($"Support pack created:\n{zipPath}");
                Process.Start(new ProcessStartInfo("explorer.exe", outRoot) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to create support pack: {ex.Message}");
            }
        }

        private static void AddFolderToZip(ZipArchive zip, string folderPath, string rootName)
        {
            foreach (var file in Directory.GetFiles(folderPath, "*", SearchOption.AllDirectories))
            {
                var rel = Path.GetRelativePath(folderPath, file);
                var entryName = Path.Combine(rootName, rel).Replace("\\", "/");
                zip.CreateEntryFromFile(file, entryName, CompressionLevel.Optimal);
            }
        }

        private static string GetComPortInfo()
        {
            var sb = new StringBuilder();
            sb.AppendLine($"Collected: {DateTime.Now:O}");
            sb.AppendLine();

            try
            {
                using var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%(COM%'");
                foreach (ManagementObject obj in searcher.Get())
                {
                    var name = (obj["Name"] ?? "").ToString();
                    var pnp = (obj["PNPDeviceID"] ?? "").ToString();
                    sb.AppendLine(name);
                    sb.AppendLine($"  PNPDeviceID: {pnp}");
                    sb.AppendLine();
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine("WMI failed: " + ex.Message);
            }

            return sb.ToString();
        }
    }
}
