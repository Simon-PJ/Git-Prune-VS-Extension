using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.Windows;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace GitPrune
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class PruneCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0180;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("9efc1b39-15e7-4465-a175-1dd034081d15");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="PruneCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private PruneCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.PruneCurrentGitDirectory, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static PruneCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new PruneCommand(package);
        }

        private void PruneCurrentGitDirectory(object sender, EventArgs e)
        {
            var dte = (DTE)Package.GetGlobalService(typeof(DTE));

            if (string.IsNullOrEmpty(dte.Solution.FullName)) return;

            var directory = System.IO.Path.GetDirectoryName(dte.Solution.FullName);

            if (!System.IO.Directory.Exists($"{directory}/.git")) return;

            var displayString = RunPruneCommand(directory);

            MessageBox.Show(displayString);
        }

        private string RunPruneCommand(string solutionDirectory)
        {
            var pruneCmd = $"/c cd {solutionDirectory}&git remote prune origin";

            var process = new System.Diagnostics.Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "CMD",
                    WindowStyle = ProcessWindowStyle.Hidden,
                    Arguments = pruneCmd,
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            process.Start();

            var output = process.StandardOutput.ReadToEnd();
            var displayString = string.IsNullOrEmpty(output) ? "Nothing to prune" : output;

            return displayString;
        }
    }
}
