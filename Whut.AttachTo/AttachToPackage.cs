﻿using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;

namespace AttachTo
{
    using EnvDTE;

    using Constants = Microsoft.VisualStudio.Shell.Interop.Constants;

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
            AddAttachToCommand(mcs, ProgramsList.DotNet, gop => gop.ShowAttachToDotNet, "dotnet.exe");
        }

        private void AddAttachToCommand(OleMenuCommandService mcs, 
            uint commandId, 
            Func<GeneralOptionsPage, bool> isVisible, 
            params string[] programsToAttach)
        {
            var menuItemCommand = new OleMenuCommand((sender, e) =>
            {
                var statusBar = (IVsStatusbar)GetService(typeof(SVsStatusbar));
                foreach (string program in programsToAttach)
                {
                    NotifyInStatusBar(statusBar, 
                        string.Format("AttachTo: Attaching to process {0}", 
                        program));
                    if (!Attach(statusBar, program))
                    {
                        NotifyInStatusBar(statusBar,
                            string.Format("AttachTo: process {0} is not running!",
                            string.Join(" ", programsToAttach)));
                    }
                }
            }, 
            new CommandID(GuidList.guidAttachToCmdSet, (int)commandId));

            menuItemCommand.BeforeQueryStatus += (s, e) =>  menuItemCommand.Visible = isVisible((GeneralOptionsPage)GetDialogPage(typeof(GeneralOptionsPage)));
            mcs.AddCommand(menuItemCommand);
        }

        /// <summary>
        /// Attaches to running process
        /// </summary>
        /// <param name="statusBar"></param>
        /// <param name="processName"></param>
        /// <returns>True if attaching succeeds</returns>
        private bool Attach(IVsStatusbar statusBar, string processName)
        {
            object icon = (short)Constants.SBAI_Find;
            statusBar.Animation(1, ref icon);
            try
            {
                var dte = (DTE)GetService(typeof(DTE));
                foreach (Process process in dte.Debugger.LocalProcesses)
                {
                    if (process.Name.EndsWith(processName))
                    {
                        process.Attach();
                    }
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                // Stop the animation. 
                statusBar.Animation(0, ref icon);
            }
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
    }
}
