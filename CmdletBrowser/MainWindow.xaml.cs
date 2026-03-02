using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;

namespace CmdletBrowser
{
    public class CommandItem
    {
        public string Name { get; set; }
        public string ModuleName { get; set; }
        public string CommandType { get; set; }
        public string Source { get; set; }
        public string Version { get; set; }
        public string Path { get; set; }
    }

    public class ParameterItem
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string Required { get; set; }
        public string Position { get; set; }
        public string Pipeline { get; set; }
        public string Aliases { get; set; }
    }

    internal class HelpResult
    {
        public string Synopsis { get; set; } = string.Empty;
        public string Syntax { get; set; } = string.Empty;
        public string Examples { get; set; } = string.Empty;
        public List<ParameterItem> Parameters { get; set; } = new List<ParameterItem>();
    }

    public partial class MainWindow : Window
    {
        private List<CommandItem> _allCommands = new List<CommandItem>();
        private List<CommandItem> _filtered = new List<CommandItem>();
        private bool _loading = false;
        private static readonly string WinPS = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), @"WindowsPowerShell\v1.0\powershell.exe");

        private static Runspace NewRunspace()
        {
            var rs = RunspaceFactory.CreateRunspace();
            rs.Open();
            return rs;
        }

        public MainWindow() => InitializeComponent();

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            using (var identity = WindowsIdentity.GetCurrent())
            {
                var principal = new WindowsPrincipal(identity);
                if (principal.IsInRole(WindowsBuiltInRole.Administrator))
                {
                    this.Title += " (Administrator)";
                }
            }
            LoadCommands();
        }

        #region Load and Filter Commands
        private async Task LoadCommands() // Changed to return a Task
        {
            if (_loading) return;
            _loading = true;
            SetStatus("Loading commands...");
            Cursor = Cursors.Wait;
            RefreshBtn.IsEnabled = false;
            DeleteModuleBtn.IsEnabled = false;

            _allCommands = await Task.Run(() => FetchCommands());

            BuildModuleTree();
            ApplyFilter();

            Cursor = Cursors.Arrow;
            RefreshBtn.IsEnabled = true;
            SetStatus($"Loaded {_allCommands.Count:N0} total commands.");
            _loading = false;
        }

        private static List<CommandItem> FetchCommands()
        {
            using (var rs = NewRunspace())
            using (var ps = PowerShell.Create())
            {
                ps.Runspace = rs;
                ps.AddCommand("Get-Command").AddParameter("ErrorAction", "SilentlyContinue");
                return ps.Invoke()
                    .Where(r => r?.BaseObject is CommandInfo)
                    .Select(r => (CommandInfo)r.BaseObject)
                    .OrderBy(c => c.Name)
                    .Select(c => new CommandItem
                    {
                        Name = c.Name,
                        ModuleName = c.ModuleName ?? string.Empty,
                        CommandType = c.CommandType.ToString(),
                        Source = c.Source ?? string.Empty,
                        Version = c.Module?.Version?.ToString() ?? "N/A",
                        Path = c.Module?.Path ?? string.Empty
                    })
                    .ToList();
            }
        }

        private void BuildModuleTree()
        {
            ModuleTree.Items.Clear();
            var root = new TreeViewItem { Header = $"All Modules ({_allCommands.Count(c => !string.IsNullOrEmpty(c.ModuleName))})", Tag = "*", IsExpanded = true };
            ModuleTree.Items.Add(root);

            foreach (var g in _allCommands.Where(c => !string.IsNullOrEmpty(c.ModuleName)).GroupBy(c => c.ModuleName).OrderBy(g => g.Key))
            {
                root.Items.Add(new TreeViewItem { Header = $"{g.Key} ({g.Count()})", Tag = g.Key });
            }

            ModuleTree.SelectedItemChanged -= ModuleTree_SelectedItemChanged;
            root.IsSelected = true;
            ModuleTree.SelectedItemChanged += ModuleTree_SelectedItemChanged;
        }

        private void ApplyFilter()
        {
            var search = SearchBox.Text.Trim();
            string module = (ModuleTree.SelectedItem as TreeViewItem)?.Tag as string;
            if (module == "*") module = null;

            IEnumerable<CommandItem> list = _allCommands;

            if (!string.IsNullOrEmpty(module))
                list = list.Where(c => c.ModuleName.Equals(module, StringComparison.OrdinalIgnoreCase));

            var typesToShow = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Cmdlet" };
            if (IncludeFnBox.IsChecked == true) typesToShow.Add("Function");
            if (IncludeAlBox.IsChecked == true) typesToShow.Add("Alias");
            list = list.Where(c => typesToShow.Contains(c.CommandType));

            if (!string.IsNullOrEmpty(search))
                list = list.Where(c => c.Name.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0);

            _filtered = list.ToList();
            CmdList.ItemsSource = _filtered;
            SetStatus($"Showing {_filtered.Count:N0} item(s).");
        }
        #endregion

        #region Help Methods
        private async void LoadHelp(string name) { SelectedCmdText.Text = name; SynopsisBox.Text = string.Empty; SyntaxBox.Text = string.Empty; ExamplesBox.Text = string.Empty; ParamsGrid.ItemsSource = null; if (string.IsNullOrWhiteSpace(name)) return; SetStatus($"Loading help for {name}..."); var result = await Task.Run(() => FetchHelp(name)); SynopsisBox.Text = result.Synopsis; SyntaxBox.Text = result.Syntax; ExamplesBox.Text = result.Examples; ParamsGrid.ItemsSource = result.Parameters; SetStatus($"Help loaded for {name}."); }
        private static HelpResult FetchHelp(string name) { var result = new HelpResult(); try { using (var rs = NewRunspace()) { var cmdInfo = GetCommandInfo(name, rs); var helpObj = RunGetHelpFull(name, rs); result.Synopsis = ExtractSynopsis(helpObj); result.Syntax = BuildSyntax(cmdInfo, name, rs); result.Examples = BuildExamples(helpObj); result.Parameters = BuildParametersFromCommandInfo(cmdInfo, helpObj); } } catch (Exception ex) { result.Synopsis = $"Error loading help: {ex.Message}"; } if (string.IsNullOrWhiteSpace(result.Synopsis)) result.Synopsis = "No local synopsis available."; if (string.IsNullOrWhiteSpace(result.Syntax)) result.Syntax = "No syntax available."; if (string.IsNullOrWhiteSpace(result.Examples)) result.Examples = "No examples available."; return result; }
        private static CommandInfo GetCommandInfo(string name, Runspace rs) { using (var ps = PowerShell.Create()) { ps.Runspace = rs; ps.AddCommand("Get-Command").AddParameter("Name", name).AddParameter("ErrorAction", "SilentlyContinue"); return ps.Invoke().FirstOrDefault()?.BaseObject as CommandInfo; } }
        private static PSObject RunGetHelpFull(string name, Runspace rs) { try { using (var ps = PowerShell.Create()) { ps.Runspace = rs; ps.AddScript($"Get-Help -Name '{name.Replace("'", "''")}' -Full -ErrorAction SilentlyContinue"); return ps.Invoke().FirstOrDefault(); } } catch { return null; } }
        private static string ExtractSynopsis(PSObject help) { if (help == null) return string.Empty; var s = help.Properties["Synopsis"]?.Value?.ToString()?.Trim(); if (!string.IsNullOrEmpty(s) && !s.StartsWith("Get-Help ")) return s; var details = help.Properties["details"]?.Value as PSObject; var text = PsFirstText(details?.Properties["description"]?.Value); return !string.IsNullOrEmpty(text) ? text : string.Empty; }
        private static string BuildSyntax(CommandInfo cmdInfo, string name, Runspace rs) { if (cmdInfo is not (CmdletInfo or FunctionInfo)) return string.Empty; try { if (cmdInfo.ParameterSets != null && cmdInfo.ParameterSets.Count > 0) { return string.Join(Environment.NewLine + Environment.NewLine, cmdInfo.ParameterSets.Select(set => { var line = new StringBuilder(name); foreach (var p in set.Parameters.OrderBy(p => p.Position < 0 ? 999 : p.Position)) { if (IsCommonParameter(p.Name)) continue; string token = p.ParameterType == typeof(SwitchParameter) ? $"[-{p.Name}]" : $"[-{p.Name} <{p.ParameterType.Name}>]"; if (p.IsMandatory) token = token.Trim('[', ']'); line.Append($" {token}"); } return line.ToString(); })); } } catch { } return string.Empty; }
        private static readonly HashSet<string> _commonParams = new(StringComparer.OrdinalIgnoreCase) { "Verbose", "Debug", "ErrorAction", "WarningAction", "InformationAction", "ErrorVariable", "WarningVariable", "InformationVariable", "OutVariable", "OutBuffer", "PipelineVariable", "WhatIf", "Confirm", "ProgressAction" };
        private static bool IsCommonParameter(string name) => _commonParams.Contains(name);
        private static List<ParameterItem> BuildParametersFromCommandInfo(CommandInfo cmdInfo, PSObject help) { if (cmdInfo?.Parameters == null) return new List<ParameterItem>(); return cmdInfo.Parameters.Values.Where(p => !IsCommonParameter(p.Name)).OrderBy(p => p.Name).Select(p => { var attrs = p.Attributes.OfType<ParameterAttribute>().ToList(); bool pipeline = attrs.Any(a => a.ValueFromPipeline || a.ValueFromPipelineByPropertyName); return new ParameterItem { Name = p.Name, Type = FriendlyTypeName(p.ParameterType), Required = attrs.Any(a => a.Mandatory).ToString(), Position = attrs.Where(a => a.Position >= 0).Select(a => (int?)a.Position).FirstOrDefault()?.ToString() ?? "Named", Pipeline = pipeline ? (attrs.Any(a => a.ValueFromPipeline) ? "true (ByValue)" : "true (ByPropertyName)") : "false", Aliases = string.Join(", ", p.Aliases) }; }).ToList(); }
        private static string FriendlyTypeName(Type t) { if (t == null) return string.Empty; if (t == typeof(SwitchParameter)) return "SwitchParameter"; if (t.IsArray) return FriendlyTypeName(t.GetElementType()) + "[]"; var u = Nullable.GetUnderlyingType(t); return u != null ? FriendlyTypeName(u) + "?" : t.Name; }
        private static string BuildExamples(PSObject help) { if (help?.Properties["examples"]?.Value is not PSObject examplesNode) return string.Empty; if (examplesNode.Properties["example"]?.Value is not object exampleProp) return string.Empty; var sb = new StringBuilder(); foreach (var raw in ToObjectList(exampleProp)) { if (AsPSObject(raw) is not PSObject ex) continue; var title = Regex.Replace(ex.Properties["title"]?.Value?.ToString() ?? string.Empty, @"^EXAMPLE\s*\d+\s*[-–]?\s*", string.Empty, RegexOptions.IgnoreCase).Trim(); var code = ex.Properties["code"]?.Value?.ToString()?.Trim() ?? string.Empty; var remarks = ExtractRemarksText(ex.Properties["remarks"]?.Value); if (!string.IsNullOrEmpty(title)) sb.AppendLine($"# {title}"); if (!string.IsNullOrEmpty(code)) sb.AppendLine(code); if (!string.IsNullOrEmpty(remarks)) sb.AppendLine(remarks); sb.AppendLine(); } return sb.ToString().Trim(); }
        private static string ExtractRemarksText(object remarksRaw) { if (remarksRaw == null) return string.Empty; var parts = ToObjectList(remarksRaw).Select(item => AsPSObject(item)?.Properties["Text"]?.Value?.ToString()?.Trim() ?? item.ToString().Trim()); return string.Join(Environment.NewLine, parts.Where(s => !string.IsNullOrEmpty(s))); }
        #endregion

        #region UI and Event Handlers
        private void TryShowHelpWindow(string name) { if (!File.Exists(WinPS)) { MessageBox.Show("Windows PowerShell 5.1 not found.", "Show Help Window", MessageBoxButton.OK, MessageBoxImage.Information); OpenOnlineHelp(name); return; } try { Process.Start(new ProcessStartInfo { FileName = WinPS, Arguments = $"-NoProfile -NonInteractive -WindowStyle Hidden -Command \"Get-Help -Name '{name.Replace("'", "''")}' -ShowWindow; Start-Sleep 5\"", UseShellExecute = false, CreateNoWindow = true }); SetStatus($"Help window opened for {name}."); } catch (Exception ex) { MessageBox.Show($"Could not open help window: {ex.Message}", "Show Help Window", MessageBoxButton.OK, MessageBoxImage.Warning); OpenOnlineHelp(name); } }
        private static void OpenOnlineHelp(string name) { try { Process.Start(new ProcessStartInfo($"https://learn.microsoft.com/powershell/module/?term={Uri.EscapeDataString(name)}") { UseShellExecute = true }); } catch { } }
        private static List<object> ToObjectList(object value) { if (value == null) return new List<object>(); if (value is System.Collections.IEnumerable e && value is not string) return e.Cast<object>().Where(o => o != null).ToList(); return new List<object> { value }; }
        private static PSObject AsPSObject(object obj) => obj as PSObject ?? PSObject.AsPSObject(obj);
        private static string PsFirstText(object value) => ToObjectList(value).Select(item => AsPSObject(item)?.Properties["Text"]?.Value?.ToString()?.Trim() ?? item.ToString().Trim()).FirstOrDefault(s => !string.IsNullOrEmpty(s)) ?? string.Empty;
        private void SetStatus(string text) => Dispatcher.InvokeAsync(() => StatusText.Text = text);

        private async void RefreshBtn_Click(object sender, RoutedEventArgs e) => await LoadCommands();
        private void FilterOption_Changed(object sender, RoutedEventArgs e) { if (IsLoaded) ApplyFilter(); }
        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e) => ApplyFilter();

        private void ModuleTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (ModuleTree.SelectedItem is TreeViewItem item && item.Tag is string moduleName && moduleName != "*")
            {
                var cmd = _allCommands.FirstOrDefault(c => c.ModuleName.Equals(moduleName, StringComparison.OrdinalIgnoreCase));
                DeleteModuleBtn.IsEnabled = cmd != null &&
                                            !string.IsNullOrEmpty(cmd.Path) &&
                                            !cmd.Path.StartsWith(Environment.GetFolderPath(Environment.SpecialFolder.System), StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                DeleteModuleBtn.IsEnabled = false;
            }
            ApplyFilter();
        }

        private void CmdList_SelectionChanged(object sender, SelectionChangedEventArgs e) { if (CmdList.SelectedItem is CommandItem item) LoadHelp(item.Name); }
        private void CmdList_MouseDoubleClick(object sender, MouseButtonEventArgs e) { if (CmdList.SelectedItem is CommandItem item) TryShowHelpWindow(item.Name); }
        private void CopyNameBtn_Click(object sender, RoutedEventArgs e) { if (CmdList.SelectedItem is CommandItem item) { Clipboard.SetText(item.Name); SetStatus($"Copied name: {item.Name}"); } }
        private void CopySyntaxBtn_Click(object sender, RoutedEventArgs e) { if (!string.IsNullOrWhiteSpace(SyntaxBox.Text)) { Clipboard.SetText(SyntaxBox.Text); SetStatus("Copied syntax."); } }
        private void HelpWindowBtn_Click(object sender, RoutedEventArgs e) { if (CmdList.SelectedItem is CommandItem item) TryShowHelpWindow(item.Name); }
        private void HelpOnlineBtn_Click(object sender, RoutedEventArgs e) { if (CmdList.SelectedItem is CommandItem item) OpenOnlineHelp(item.Name); }
        private void ExportBtn_Click(object sender, RoutedEventArgs e) { var dlg = new SaveFileDialog { Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*", FileName = "Commands.csv" }; if (dlg.ShowDialog() != true) return; try { var sb = new StringBuilder(); sb.AppendLine("Name,ModuleName,CommandType,Source,Version,Path"); foreach (var c in _filtered) sb.AppendLine($"{Csv(c.Name)},{Csv(c.ModuleName)},{Csv(c.CommandType)},{Csv(c.Source)},{Csv(c.Version)},{Csv(c.Path)}"); File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8); SetStatus($"Exported {_filtered.Count:N0} item(s) to '{dlg.FileName}'."); } catch (Exception ex) { MessageBox.Show($"Export failed: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error); } }
        private static string Csv(string v) => (v != null && v.Contains(',')) ? $"\"{v.Replace("\"", "\"\"")}\"" : v;
        #endregion

        #region Delete Module
        private async void DeleteModuleBtn_Click(object sender, RoutedEventArgs e)
        {
            if (ModuleTree.SelectedItem is not TreeViewItem item || item.Tag is not string moduleName || moduleName == "*") return;

            var commandInModule = _allCommands.FirstOrDefault(c => c.ModuleName.Equals(moduleName, StringComparison.OrdinalIgnoreCase));
            if (commandInModule == null || string.IsNullOrEmpty(commandInModule.Path))
            {
                MessageBox.Show("Could not determine the path for this module.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string modulePath = Path.GetDirectoryName(commandInModule.Path);
            string message = $"You are about to force-delete the module '{moduleName}'.\n\nThis will permanently remove the folder and all its contents.\n\nPath:\n{modulePath}\n\nDo you want to proceed?";

            var result = MessageBox.Show(message, "Confirm Force Deletion", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result != MessageBoxResult.Yes) { SetStatus("Deletion cancelled."); return; }

            // --- THIS IS THE IMPROVED WORKFLOW ---

            // 1. Immediately start the UI refresh to give instant feedback.
            Task refreshTask = LoadCommands();

            // 2. Run the deletion in the background.
            bool success = await Task.Run(() => ForceDeleteWithPowerShell(modulePath));

            // 3. Wait for the UI refresh to complete.
            await refreshTask;

            // 4. Now show the final result message.
            if (success)
            {
                MessageBox.Show($"Module '{moduleName}' was successfully deleted and the list has been refreshed.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show($"Failed to delete the module.\n\nThis can happen if the application does not have Administrator privileges or if a file is in use by a process other than this one.", "Deletion Failed", MessageBoxButton.OK, MessageBoxImage.Error);
                SetStatus("Deletion failed.");
            }
        }

        private static bool ForceDeleteWithPowerShell(string path)
        {
            if (!Directory.Exists(path))
            {
                return true;
            }

            string script = $"Remove-Item -Path '{path}' -Recurse -Force -ErrorAction Stop";

            var startInfo = new ProcessStartInfo()
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -WindowStyle Hidden -Command \"{script}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
            };

            try
            {
                using (var process = Process.Start(startInfo))
                {
                    process.WaitForExit();
                    return process.ExitCode == 0;
                }
            }
            catch
            {
                return false;
            }
        }
        #endregion
    }
}
