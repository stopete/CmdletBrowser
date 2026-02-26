using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
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

        // Path to Windows PowerShell 5.1 — always at this location on Windows
        private static readonly string WinPS =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),
                         @"WindowsPowerShell\v1.0\powershell.exe");

        private static Runspace NewRunspace()
        {
            var rs = RunspaceFactory.CreateRunspace();
            rs.Open();
            return rs;
        }

        public MainWindow() => InitializeComponent();
        private void Window_Loaded(object sender, RoutedEventArgs e) => LoadCommands();

        // ---------------------------------------------------------------
        // Load commands
        // ---------------------------------------------------------------
        private async void LoadCommands()
        {
            if (_loading) return;
            _loading = true;
            SetStatus("Loading commands...");
            Cursor = Cursors.Wait;
            RefreshBtn.IsEnabled = false;

            bool incFn = IncludeFnBox.IsChecked == true;
            bool incAl = IncludeAlBox.IsChecked == true;

            var commands = await Task.Run(() => FetchCommands(incFn, incAl));
            _allCommands = commands;
            BuildModuleTree();
            ApplyFilter();

            Cursor = Cursors.Arrow;
            RefreshBtn.IsEnabled = true;
            SetStatus($"Loaded {_allCommands.Count:N0} command(s).");
            _loading = false;
        }

        private static List<CommandItem> FetchCommands(bool incFn, bool incAl)
        {
            var types = CommandTypes.Cmdlet;
            if (incFn) types |= CommandTypes.Function;
            if (incAl) types |= CommandTypes.Alias;

            using (var rs = NewRunspace())
            using (var ps = PowerShell.Create())
            {
                ps.Runspace = rs;
                ps.AddCommand("Get-Command")
                  .AddParameter("CommandType", types)
                  .AddParameter("ErrorAction", "SilentlyContinue");

                return ps.Invoke()
                    .Where(r => r?.BaseObject is CommandInfo)
                    .Select(r => (CommandInfo)r.BaseObject)
                    .OrderBy(c => c.Name)
                    .Select(c => new CommandItem
                    {
                        Name = c.Name,
                        ModuleName = c.ModuleName ?? string.Empty,
                        CommandType = c.CommandType.ToString(),
                        Source = c.Source ?? string.Empty
                    })
                    .ToList();
            }
        }

        // ---------------------------------------------------------------
        // Module tree
        // ---------------------------------------------------------------
        private void BuildModuleTree()
        {
            ModuleTree.Items.Clear();
            var root = new TreeViewItem
            {
                Header = $"All Modules ({_allCommands.Count})",
                Tag = "*",
                IsExpanded = true
            };
            ModuleTree.Items.Add(root);

            foreach (var g in _allCommands
                .Where(c => !string.IsNullOrEmpty(c.ModuleName))
                .GroupBy(c => c.ModuleName)
                .OrderBy(g => g.Key))
            {
                root.Items.Add(new TreeViewItem
                {
                    Header = $"{g.Key} ({g.Count()})",
                    Tag = g.Key
                });
            }

            ModuleTree.SelectedItemChanged -= ModuleTree_SelectedItemChanged;
            root.IsSelected = true;
            ModuleTree.SelectedItemChanged += ModuleTree_SelectedItemChanged;
        }

        // ---------------------------------------------------------------
        // Filter
        // ---------------------------------------------------------------
        private void ApplyFilter()
        {
            var search = (SearchBox.Text ?? string.Empty).Trim();
            string module = (ModuleTree.SelectedItem as TreeViewItem)?.Tag as string;
            if (module == "*") module = null;

            IEnumerable<CommandItem> list = _allCommands;
            if (!string.IsNullOrEmpty(module))
                list = list.Where(c => c.ModuleName == module);
            if (!string.IsNullOrEmpty(search))
                list = list.Where(c => c.Name.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0);

            _filtered = list.ToList();
            CmdList.ItemsSource = _filtered;
            SetStatus($"Showing {_filtered.Count:N0} item(s).");
        }

        // ---------------------------------------------------------------
        // Load help
        // ---------------------------------------------------------------
        private async void LoadHelp(string name)
        {
            SelectedCmdText.Text = name;
            SynopsisBox.Text = string.Empty;
            SyntaxBox.Text = string.Empty;
            ExamplesBox.Text = string.Empty;
            ParamsGrid.ItemsSource = null;
            if (string.IsNullOrWhiteSpace(name)) return;
            SetStatus($"Loading help for {name}...");

            var result = await Task.Run(() => FetchHelp(name));

            SynopsisBox.Text = result.Synopsis;
            SyntaxBox.Text = result.Syntax;
            ExamplesBox.Text = result.Examples;
            ParamsGrid.ItemsSource = result.Parameters;
            SetStatus($"Help loaded for {name}.");
        }

        private static HelpResult FetchHelp(string name)
        {
            var result = new HelpResult();
            try
            {
                using (var rs = NewRunspace())
                {
                    var cmdInfo = GetCommandInfo(name, rs);
                    var helpObj = RunGetHelpFull(name, rs);

                    result.Synopsis = ExtractSynopsis(helpObj);
                    result.Syntax = BuildSyntax(cmdInfo, name, rs);
                    result.Examples = BuildExamples(helpObj);
                    result.Parameters = BuildParametersFromCommandInfo(cmdInfo, helpObj);
                }
            }
            catch (Exception ex)
            {
                result.Synopsis = $"Error loading help: {ex.Message}";
            }

            if (string.IsNullOrWhiteSpace(result.Synopsis))
                result.Synopsis = "No local synopsis available. Try: Update-Help -ErrorAction SilentlyContinue";
            if (string.IsNullOrWhiteSpace(result.Syntax))
                result.Syntax = "No syntax available.";
            if (string.IsNullOrWhiteSpace(result.Examples))
                result.Examples = "No examples available.";

            return result;
        }

        // ---------------------------------------------------------------
        // Get CommandInfo
        // ---------------------------------------------------------------
        private static CommandInfo GetCommandInfo(string name, Runspace rs)
        {
            using (var ps = PowerShell.Create())
            {
                ps.Runspace = rs;
                ps.AddCommand("Get-Command")
                  .AddParameter("Name", name)
                  .AddParameter("ErrorAction", "SilentlyContinue");
                return ps.Invoke()
                         .FirstOrDefault(r => r?.BaseObject is CommandInfo)
                         ?.BaseObject as CommandInfo;
            }
        }

        // ---------------------------------------------------------------
        // Get-Help -Full
        // ---------------------------------------------------------------
        private static PSObject RunGetHelpFull(string name, Runspace rs)
        {
            try
            {
                using (var ps = PowerShell.Create())
                {
                    ps.Runspace = rs;
                    ps.AddScript($"Get-Help -Name '{name.Replace("'", "''")}' -Full -ErrorAction SilentlyContinue");
                    return ps.Invoke().FirstOrDefault();
                }
            }
            catch { return null; }
        }

        // ---------------------------------------------------------------
        // Synopsis
        // ---------------------------------------------------------------
        private static string ExtractSynopsis(PSObject help)
        {
            if (help == null) return string.Empty;
            var s = help.Properties["Synopsis"]?.Value?.ToString()?.Trim();
            if (!string.IsNullOrEmpty(s) && !s.StartsWith("Get-Help ")) return s;

            var details = help.Properties["details"]?.Value as PSObject;
            if (details != null)
            {
                var text = PsFirstText(details.Properties["description"]?.Value);
                if (!string.IsNullOrEmpty(text)) return text;
            }
            return string.Empty;
        }

        // ---------------------------------------------------------------
        // Syntax
        // ---------------------------------------------------------------
        private static string BuildSyntax(CommandInfo cmdInfo, string name, Runspace rs)
        {
            if (cmdInfo is CmdletInfo || cmdInfo is FunctionInfo)
            {
                try
                {
                    var paramSets = cmdInfo.ParameterSets;
                    if (paramSets != null && paramSets.Count > 0)
                    {
                        var sb = new StringBuilder();
                        int i = 0;
                        foreach (var set in paramSets)
                        {
                            i++;
                            sb.AppendLine($"PARAMETER SET {i}{(set.IsDefault ? " (default)" : string.Empty)}: {set.Name}");
                            sb.AppendLine("----------------------------------------");

                            var line = new StringBuilder(name);
                            foreach (var p in set.Parameters.OrderBy(p => p.Position < 0 ? 999 : p.Position))
                            {
                                if (IsCommonParameter(p.Name)) continue;
                                var typeName = p.ParameterType.Name;
                                bool isSwitch = p.ParameterType == typeof(bool) ||
                                                p.ParameterType == typeof(SwitchParameter);
                                string token = isSwitch
                                    ? (p.IsMandatory ? $"-{p.Name}" : $"[-{p.Name}]")
                                    : (p.IsMandatory ? $"-{p.Name} <{typeName}>" : $"[-{p.Name} <{typeName}>]");
                                line.Append(" " + token);
                            }
                            sb.AppendLine(line.ToString());
                            sb.AppendLine();
                        }
                        return sb.ToString().Trim();
                    }
                }
                catch { /* fall through */ }
            }

            // Fallback
            try
            {
                using (var ps = PowerShell.Create())
                {
                    ps.Runspace = rs;
                    ps.AddScript($"Get-Command -Name '{name.Replace("'", "''")}' -Syntax -ErrorAction SilentlyContinue");
                    var raw = string.Join(Environment.NewLine,
                        ps.Invoke().Select(r => r?.ToString() ?? string.Empty)).Trim();
                    if (string.IsNullOrEmpty(raw)) return string.Empty;

                    var sets = raw.Split(new[] { "\r\n\r\n", "\n\n" }, StringSplitOptions.RemoveEmptyEntries)
                                  .Select(s => s.Trim()).Where(s => s.Length > 0).ToArray();
                    var sb = new StringBuilder();
                    for (int i = 0; i < sets.Length; i++)
                    {
                        sb.AppendLine($"PARAMETER SET {i + 1}");
                        sb.AppendLine("----------------------------------------");
                        sb.AppendLine(sets[i]);
                        sb.AppendLine();
                    }
                    return sb.ToString().Trim();
                }
            }
            catch { return string.Empty; }
        }

        private static readonly HashSet<string> _commonParams = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Verbose","Debug","ErrorAction","WarningAction","InformationAction",
            "ErrorVariable","WarningVariable","InformationVariable","OutVariable",
            "OutBuffer","PipelineVariable","WhatIf","Confirm","ProgressAction"
        };
        private static bool IsCommonParameter(string name) => _commonParams.Contains(name);

        // ---------------------------------------------------------------
        // Parameters
        // ---------------------------------------------------------------
        private static List<ParameterItem> BuildParametersFromCommandInfo(CommandInfo cmdInfo, PSObject help)
        {
            var list = new List<ParameterItem>();
            if (cmdInfo?.Parameters == null) return list;

            foreach (var kv in cmdInfo.Parameters.OrderBy(k => k.Key))
            {
                var p = kv.Value;
                if (p == null || IsCommonParameter(p.Name)) continue;

                var paramAttrs = p.Attributes.OfType<ParameterAttribute>().ToList();
                bool mandatory = paramAttrs.Any(a => a.Mandatory);
                int? position = paramAttrs.Where(a => a.Position >= 0)
                                            .Select(a => (int?)a.Position)
                                            .FirstOrDefault();
                bool pipeline = paramAttrs.Any(a => a.ValueFromPipeline || a.ValueFromPipelineByPropertyName);
                string pipeStr = pipeline
                    ? (paramAttrs.Any(a => a.ValueFromPipeline) ? "true (ByValue)" : "true (ByPropertyName)")
                    : "false";

                list.Add(new ParameterItem
                {
                    Name = p.Name,
                    Type = FriendlyTypeName(p.ParameterType),
                    Required = mandatory ? "true" : "false",
                    Position = position.HasValue ? position.Value.ToString() : "Named",
                    Pipeline = pipeStr,
                    Aliases = p.Aliases.Count > 0 ? string.Join(", ", p.Aliases) : string.Empty
                });
            }
            return list;
        }

        private static string FriendlyTypeName(Type t)
        {
            if (t == null) return string.Empty;
            if (t == typeof(SwitchParameter)) return "SwitchParameter";
            if (t.IsArray) return FriendlyTypeName(t.GetElementType()) + "[]";
            var u = Nullable.GetUnderlyingType(t);
            return u != null ? FriendlyTypeName(u) + "?" : t.Name;
        }

        // ---------------------------------------------------------------
        // Examples
        // ---------------------------------------------------------------
        private static string BuildExamples(PSObject help)
        {
            if (help == null) return string.Empty;
            try
            {
                var examplesNode = help.Properties["examples"]?.Value as PSObject;
                var exampleProp = examplesNode?.Properties["example"]?.Value;
                if (exampleProp == null) return string.Empty;

                var exampleList = ToObjectList(exampleProp);
                if (exampleList.Count == 0) return string.Empty;

                var sb = new StringBuilder();
                foreach (var raw in exampleList)
                {
                    var ex = AsPSObject(raw);
                    if (ex == null) continue;

                    var title = ex.Properties["title"]?.Value?.ToString() ?? string.Empty;
                    title = Regex.Replace(title, @"^[-\s]+|[-\s]+$", string.Empty).Trim();
                    title = Regex.Replace(title, @"^EXAMPLE\s*\d+\s*[-–]?\s*", string.Empty, RegexOptions.IgnoreCase).Trim();
                    if (!string.IsNullOrEmpty(title)) sb.AppendLine($"# {title}");

                    var code = ex.Properties["code"]?.Value?.ToString()?.Trim() ?? string.Empty;
                    if (!string.IsNullOrEmpty(code)) sb.AppendLine(code);

                    var remarksText = ExtractRemarksText(ex.Properties["remarks"]?.Value);
                    if (!string.IsNullOrEmpty(remarksText)) sb.AppendLine(remarksText);

                    sb.AppendLine();
                }
                return sb.ToString().Trim();
            }
            catch { return string.Empty; }
        }

        private static string ExtractRemarksText(object remarksRaw)
        {
            if (remarksRaw == null) return string.Empty;
            var parts = new List<string>();
            foreach (var item in ToObjectList(remarksRaw))
            {
                var pso = AsPSObject(item);
                var text = pso?.Properties["Text"]?.Value?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(text)) parts.Add(text);
                else { var s = item?.ToString()?.Trim(); if (!string.IsNullOrEmpty(s)) parts.Add(s); }
            }
            return string.Join(Environment.NewLine, parts).Trim();
        }

        // ---------------------------------------------------------------
        // Show Help Window — must use an external WinPS 5.1 process
        // because Get-Help -ShowWindow requires the WinPS GUI host and
        // will always fail inside an embedded PS7 runspace.
        // ---------------------------------------------------------------
        private void TryShowHelpWindow(string name)
        {
            // Sanitise the name so it's safe inside a quoted PS string
            var safeName = name.Replace("'", "''");

            // Check whether Windows PowerShell 5.1 is available
            if (!File.Exists(WinPS))
            {
                MessageBox.Show(
                    $"Windows PowerShell 5.1 was not found at:\n{WinPS}\n\nOpening online help instead.",
                    "Show Help Window",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                OpenOnlineHelp(name);
                return;
            }

            try
            {
                // Launch powershell.exe (WinPS 5.1) as a hidden process.
                // The -NonInteractive flag stops it waiting for input.
                // The script calls Get-Help -ShowWindow which opens the
                // native PS help viewer window, then sleeps long enough
                // for the viewer to fully open before the host exits.
                var script = $"Get-Help -Name '{safeName}' -ShowWindow; Start-Sleep -Seconds 5";

                var psi = new ProcessStartInfo
                {
                    FileName = WinPS,
                    Arguments = $"-NoProfile -NonInteractive -WindowStyle Hidden -Command \"{script}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                Process.Start(psi);
                SetStatus($"Help window opened for {name}.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Could not open help window: {ex.Message}\n\nOpening online help instead.",
                    "Show Help Window",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                OpenOnlineHelp(name);
            }
        }

        private static void OpenOnlineHelp(string name)
        {
            try
            {
                Process.Start(new ProcessStartInfo(
                    $"https://learn.microsoft.com/powershell/module/?term={Uri.EscapeDataString(name)}")
                { UseShellExecute = true });
            }
            catch { }
        }

        // ---------------------------------------------------------------
        // Shared helpers
        // ---------------------------------------------------------------
        private static List<object> ToObjectList(object value)
        {
            if (value == null) return new List<object>();
            if (value is string) return new List<object> { value };
            if (value is System.Collections.IEnumerable e)
                return e.Cast<object>().Where(o => o != null).ToList();
            return new List<object> { value };
        }

        private static PSObject AsPSObject(object obj)
        {
            if (obj == null) return null;
            return obj as PSObject ?? PSObject.AsPSObject(obj);
        }

        private static string PsFirstText(object value)
        {
            if (value == null) return string.Empty;
            foreach (var item in ToObjectList(value))
            {
                var pso = AsPSObject(item);
                var text = pso?.Properties["Text"]?.Value?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(text)) return text;
                var str = item?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(str)) return str;
            }
            return string.Empty;
        }

        private void SetStatus(string text) =>
            Dispatcher.InvokeAsync(() => StatusText.Text = text);

        // ---------------------------------------------------------------
        // UI events
        // ---------------------------------------------------------------
        private void RefreshBtn_Click(object sender, RoutedEventArgs e) => LoadCommands();
        private void FilterOption_Changed(object sender, RoutedEventArgs e) { if (IsLoaded) LoadCommands(); }
        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e) => ApplyFilter();
        private void ModuleTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e) => ApplyFilter();

        private void CmdList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CmdList.SelectedItem is CommandItem item) LoadHelp(item.Name);
        }

        private void CmdList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (CmdList.SelectedItem is CommandItem item) TryShowHelpWindow(item.Name);
        }

        private void CopyNameBtn_Click(object sender, RoutedEventArgs e)
        {
            if (CmdList.SelectedItem is CommandItem item)
            {
                Clipboard.SetText(item.Name);
                SetStatus($"Copied name: {item.Name}");
            }
        }

        private void CopySyntaxBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(SyntaxBox.Text))
            {
                Clipboard.SetText(SyntaxBox.Text);
                SetStatus("Copied syntax.");
            }
        }

        private void HelpWindowBtn_Click(object sender, RoutedEventArgs e)
        {
            if (CmdList.SelectedItem is CommandItem item) TryShowHelpWindow(item.Name);
        }

        private void HelpOnlineBtn_Click(object sender, RoutedEventArgs e)
        {
            if (CmdList.SelectedItem is CommandItem item) OpenOnlineHelp(item.Name);
        }

        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                FileName = "Commands.csv"
            };
            if (dlg.ShowDialog() == true)
            {
                try
                {
                    var sb = new StringBuilder();
                    sb.AppendLine("Name,ModuleName,CommandType,Source");
                    foreach (var c in _filtered)
                        sb.AppendLine($"{Csv(c.Name)},{Csv(c.ModuleName)},{Csv(c.CommandType)},{Csv(c.Source)}");
                    File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);
                    SetStatus($"Exported {_filtered.Count:N0} item(s) to '{dlg.FileName}'.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Export failed: {ex.Message}", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private static string Csv(string v)
        {
            if (string.IsNullOrEmpty(v)) return string.Empty;
            return (v.Contains(',') || v.Contains('"') || v.Contains('\n'))
                ? $"\"{v.Replace("\"", "\"\"")}\""
                : v;
        }
    }
}
