using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Linq;
using System.Runtime.InteropServices;

namespace AttachTo
{
    //// This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    //// This attribute is used to register the informations needed to show the this package in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    //// This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidAttachToPkgString)]
    [ProvideOptionPage(typeof(GeneralOptionsPage), "AttachTo", "General", 110, 120, false)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)]
    public sealed class AttachToPackage : Package
    {
        protected override void Initialize()
        {
            base.Initialize();
            var mcs = (OleMenuCommandService)GetService(typeof(IMenuCommandService));
            AddAttachToCommand(mcs, ProgramsList.IIS, gop => gop.ShowAttachToIIS, "W3WP.exe");
            AddAttachToCommand(mcs, ProgramsList.IISExpress, gop => gop.ShowAttachToIISExpress, "IISEXPRESS.exe");
            AddAttachToCommand(mcs, ProgramsList.NUnit, gop => gop.ShowAttachToNUnit, "nunit-agent.exe", "nunit.exe", "nunit-console.exe", 
                "nunit-agent-x86.exe", "nunit-x86.exe", "nunit-console-x86.exe");
            AddAttachToCommand(mcs, ProgramsList.ArcSOC, gop => gop.ShowAttachToArcSOC, "ArcSOC.exe");
            AddAttachToCommand(mcs, ProgramsList.ArcMap, gop => gop.ShowAttachToArcMap, "ArcMap.exe");
            AddAttachToCommand(mcs, ProgramsList.ArcCatalog, gop => gop.ShowAttachToArcCatalog, "ArcCatalog.exe");
            AddAttachToCommand(mcs, ProgramsList.DBTool, gop => gop.ShowAttachToDBTool, "DBTool.exe");
        }

        private void AddAttachToCommand(OleMenuCommandService mcs, 
            uint commandId, 
            Func<GeneralOptionsPage, bool> isVisible, 
            params string[] programsToAttach)
        {
            var menuItemCommand = new OleMenuCommand((sender, e) =>
            {
                var statusBar = (IVsStatusbar)GetService(typeof(SVsStatusbar));
                bool processFound = false;
                var dte = (DTE)GetService(typeof(DTE));
                      
                foreach (var process in dte.Debugger.LocalProcesses.Cast<Process>().Where(p => IsToAttach(p, programsToAttach)))
                {
                    processFound = true;
                    NotifyInStatusBar(statusBar, string.Format("Attaching to process {0}", process.Name));
                    NotifyAttaching(statusBar, process);
                    process.Attach();
                }

                if(!processFound)
                    NotifyInStatusBar(
                        statusBar, 
                        string.Format("AttachTo: process {0} is not running!",
                        string.Join(" ", programsToAttach)));
            }, 
            new CommandID(GuidList.guidAttachToCmdSet, (int)commandId));

            menuItemCommand.BeforeQueryStatus += (s, e) =>  menuItemCommand.Visible = isVisible((GeneralOptionsPage)GetDialogPage(typeof(GeneralOptionsPage)));
            mcs.AddCommand(menuItemCommand);
        }

        private static void NotifyAttaching(IVsStatusbar statusBar, Process process)
        {
            // Use the standard Visual Studio icon for finding files
            object icon = (short)Microsoft.VisualStudio.Shell.Interop.Constants.SBAI_Find;

            // Display the icon in the Animation region.
            statusBar.Animation(1, ref icon);

            // Do long action
            process.Attach();

            // Stop the animation. 
            statusBar.Animation(0, ref icon);
        }

        private static void NotifyInStatusBar(IVsStatusbar statusBar, string message)
        {
            // Make sure the status bar is not frozen
            int frozen;
            statusBar.IsFrozen(out frozen);
            if (frozen != 0)
            {
                statusBar.FreezeOutput(0);
            }

            // Set the status bar text and make its display static.
            statusBar.SetText(message);

            // Freeze the status bar.
            statusBar.FreezeOutput(1);

            // Clear the status bar text.
            statusBar.FreezeOutput(0);
            statusBar.Clear();
        }

        private static bool IsToAttach(Process process, params string[] programsList)
        {
            return programsList.Any(p => process.Name.EndsWith(p, StringComparison.OrdinalIgnoreCase));
        }
    }
}
