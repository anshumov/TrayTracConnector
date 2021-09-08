using CheckPointTracConnector.Properties;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Timer = System.Windows.Forms.Timer;


namespace CheckPointTracConnector
{
    internal static class Program
    {
        private static string appGuid = "35e6bbe0-7ac3-4b5b-8f08-ae3a2b1ac3c0";

        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        private static void Main()
        {
            using (Mutex mutex = new Mutex(false, "Global\\" + appGuid))
            {
                if (!mutex.WaitOne(0, false))
                {
                    return;
                }

                GC.Collect();
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                var ctx = new TrayTracConnectorApplicationContext();
                Application.Run(ctx);

            }
        }
    }

    public class TrayTracConnectorApplicationContext : ApplicationContext
    {
        private const string AppName = "CheckPointTracConnector";
        private const string StartupRegistryKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        private Dictionary<string, ConnectionState> _connList;
        private readonly ContextMenuStrip _contextMenuStrip = new ContextMenuStrip();
        private readonly Process _tracProcess;
        private readonly NotifyIcon _trayIcon = new NotifyIcon();
        private readonly Timer _timer;
        public TrayTracConnectorApplicationContext()
        {
            CheckExecAbility();
            try
            {
                _tracProcess = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = Settings.Default.ExecFileName,
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        CreateNoWindow = true
                    }
                };
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            _trayIcon.DoubleClick += _trayIcon_DoubleClick;
            
            _timer = new Timer();
            _timer.Interval = 10000;
            _timer.Tick += T_Tick;
            _timer.Start();

            _connList = GetSitesNStates();
            FillContextMenu();
        }

        private void T_Tick(object sender, EventArgs e)
        {
            _connList = GetSitesNStates();
            FillContextMenu();
        }

        private Dictionary<string, ConnectionState> GetSitesNStates()
        {
            var connList = new Dictionary<string, ConnectionState>();
            try
            {
                _tracProcess.StartInfo.Arguments = "info";
                _tracProcess.Start();

                while (!_tracProcess.StandardOutput.EndOfStream)
                {
                    var match = Regex.Match(_tracProcess.StandardOutput.ReadLine() ?? string.Empty, @"Conn (.+)\:");
                    if (!match.Success) continue;
                    
                    var site = match.Groups[1].Value;
                    var state = ConnectionState.Idle;

                    while (!_tracProcess.StandardOutput.EndOfStream)
                    {
                        match = Regex.Match(_tracProcess.StandardOutput.ReadLine() ?? string.Empty, @"status\: (.+)");
                        if (!match.Success) continue;
                        
                        switch (match.Groups[1].Value)
                        {
                            case "Idle":
                                state = ConnectionState.Idle;
                                break;
                            case "Connected":
                                state = ConnectionState.Connected;
                                if (Settings.Default.LastConnectedSite != site)
                                {
                                    Settings.Default.PreviousConnectedSite = Settings.Default.LastConnectedSite;
                                    Settings.Default.LastConnectedSite = site;
                                    Settings.Default.Save();
                                }
                                break;
                            case "Connecting":
                                state = ConnectionState.Connecting;
                                break;
                            default:
                                state = ConnectionState.Unavailable;
                                break;
                        }
                        break;
                    }
                    connList.Add(site, state);
                }
                _tracProcess.WaitForExit();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            return connList;
        }

        private async void ConnectAsync(string site)
        {
            if (string.IsNullOrEmpty(site) || _connList[site] == ConnectionState.Connected)
            {
                return;
            }

            ConnectionMenuItemsEnable(_contextMenuStrip, false);

            if (_connList.ContainsValue(ConnectionState.Connected) || _connList.ContainsValue(ConnectionState.Connecting))
            {
                Disconnect();
            }

            try
            {
                _trayIcon.Text = @"Выполняется подключение к " + site;
                _trayIcon.ShowBalloonTip(1000, "Состояние подключения", $"Выполняется подключение к {site}",
                    ToolTipIcon.Info);
                _tracProcess.StartInfo.Arguments = $"connect -s \"{site}\"";

                await Task.Run(() => _tracProcess.Start());
                Thread.Sleep(500);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                ConnectionMenuItemsEnable(_contextMenuStrip, true);
            }
            _connList = GetSitesNStates();
            FillContextMenu();
        }

        private void Disconnect()
        {
            try
            {
                _tracProcess.StartInfo.Arguments = "disconnect";
                _tracProcess.Start();
                _tracProcess.WaitForExit();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        private void FillContextMenu()
        {
            _trayIcon.Text = @"Активные подключения отсутствуют";
            var currentStateIsConnecting = false;
            var connItems = new List<ToolStripItem>();
            var newItems = false;
            foreach (var item in _connList)
            {
                ToolStripMenuItem menuItem;
                
                var foundItems = _contextMenuStrip.Items.Find(item.Key, true);
                if (foundItems.Length == 0)
                {
                    menuItem = new ToolStripMenuItem(item.Key, null, Connect, item.Key);
                    newItems = true;
                }
                else
                    menuItem = foundItems[0] as ToolStripMenuItem;

                if (menuItem != null)
                {
                    menuItem.Checked = item.Value == ConnectionState.Connected ||
                                       item.Value == ConnectionState.Connecting;
                    menuItem.Font = Settings.Default.PreviousConnectedSite == item.Key &&
                                    item.Value != ConnectionState.Connected
                        ? new Font(menuItem.Font, FontStyle.Bold)
                        : new Font(menuItem.Font, FontStyle.Regular);
                    connItems.Add(menuItem);
                }

                switch (item.Value)
                {
                    case ConnectionState.Connected:
                        _trayIcon.Text = $@"Подключено к {item.Key}";
                        currentStateIsConnecting = false;
                        break;
                    case ConnectionState.Connecting:
                        _trayIcon.Text = $@"Выполняется подключение к {item.Key}";
                        currentStateIsConnecting = true;
                        break;
                }
            }

            if (_contextMenuStrip.Items.Count == 0 || newItems)
            {
                _contextMenuStrip.Items.Clear();
                _contextMenuStrip.Items.AddRange(connItems.ToArray());
                _contextMenuStrip.Items.Add(new ToolStripSeparator());
                _contextMenuStrip.Items.Add("Обновить список сайтов", null, RefreshSites);
                if (CheckForStartup())
                {
                    _contextMenuStrip.Items.Add("Удалить из автозагрузки", null, RemoveFromStartup);
                }
                else
                {
                    _contextMenuStrip.Items.Add("Добавить в автозагрузку", null, AddToStartup);
                }

                _contextMenuStrip.Items.Add(new ToolStripSeparator());
                _contextMenuStrip.Items.Add("Exit", null, Exit);

                _trayIcon.Icon = Icon.FromHandle(Resources.vpn_32_white.GetHicon());
                _trayIcon.ContextMenuStrip = _contextMenuStrip;
                _trayIcon.Visible = true;
            }

            ConnectionMenuItemsEnable(_contextMenuStrip, !currentStateIsConnecting);
        }
        private void CheckExecAbility()
        {
            var fileInfo = new FileInfo(Settings.Default.ExecFileName);

            if (!fileInfo.Exists)
            {
                var ofd = new OpenFileDialog
                {
                    CheckFileExists = true,
                    Multiselect = false,
                    FileName = fileInfo.Name,
                    InitialDirectory = fileInfo.DirectoryName,
                    Filter = @"trac.exe|trac.exe"
                };

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    Settings.Default.ExecFileName = ofd.FileName;
                    Settings.Default.Save();
                }
            }
        }
        private void ConnectionMenuItemsEnable(ContextMenuStrip contextMenuStrip, bool enabled)
        {
            foreach (ToolStripItem item in contextMenuStrip.Items)
            {
                if (_connList.Any(i => i.Key == item.Text))
                {
                    item.Enabled = enabled;
                }
            }
        }
        private bool CheckForStartup()
        {
            var rkApp = Registry.CurrentUser.OpenSubKey(StartupRegistryKey, false);
            return !string.IsNullOrEmpty(rkApp?.GetValue(AppName) as string);
        }

        private void AddToStartup(object sender, EventArgs eventArgs)
        {
            var rkApp = Registry.CurrentUser.OpenSubKey(StartupRegistryKey, true);
            var startPath = Environment.GetFolderPath(Environment.SpecialFolder.Programs) +
                            @"\Sibintek LLC\CheckPoint Connector.appref-ms";
            rkApp?.SetValue(AppName, startPath);
            FillContextMenu();
        }

        private void RemoveFromStartup(object sender, EventArgs eventArgs)
        {
            var rkApp = Registry.CurrentUser.OpenSubKey(StartupRegistryKey, true);
            rkApp?.DeleteValue(AppName);
            FillContextMenu();
        }

        private void Connect(object sender, EventArgs e)
        {
            if (sender is ToolStripItem s)
            {
                ConnectAsync(s.Text);
            }
        }

        private void RefreshSites(object sender, EventArgs e)
        {
            _connList = GetSitesNStates();
            FillContextMenu();
        }

        private void _trayIcon_DoubleClick(object sender, EventArgs e)
        {
            var previousSite = Settings.Default.PreviousConnectedSite;
            if (!string.IsNullOrEmpty(previousSite))
            {
                ConnectAsync(previousSite);
            }
        }

        private void Exit(object sender, EventArgs e)
        {
            _timer.Stop();
            _trayIcon.Visible = false;
            Application.Exit();
        }
    }

    [Flags]
    internal enum ConnectionState
    {
        Connected = 0,
        Idle = 1,
        Connecting = 2,
        Unavailable = 4
    }
}