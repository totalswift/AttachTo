﻿using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.Linq;
using System.Runtime.InteropServices;

namespace Whut.AttachTo
{
    //// This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    //// This attribute is used to register the informations needed to show the this package in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    //// This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidAttachToPkgString)]
    [ProvideOptionPage(typeof(GeneralOptionsPage), "Whut.AttachTo", "General", 110, 120, false)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string)]
    public sealed class AttachToPackage : Package
    {
        protected override void Initialize()
        {
            base.Initialize();

            OleMenuCommandService mcs = (OleMenuCommandService)GetService(typeof(IMenuCommandService));

            this.AddAttachToCommand(mcs, PkgCmdIDList.cmdidWhutAttachToIIS, gop => gop.ShowAttachToIIS, "w3wp.exe");
            this.AddAttachToCommand(mcs, PkgCmdIDList.cmdidWhutAttachToIISExpress, gop => gop.ShowAttachToIISExpress, "iisexpress.exe");

            this.AddAttachToCommand(mcs, PkgCmdIDList.cmdidWhutAttachToNUnit, gop => gop.ShowAttachToNUnit, "nunit-agent.exe", "nunit.exe", "nunit-console.exe", 
                "nunit-agent-x86.exe", "nunit-x86.exe", "nunit-console-x86.exe");

            this.AddAttachToCommand(mcs, PkgCmdIDList.cmdidWhutAttachToArcSOC, gop => gop.ShowAttachToArcSOC, "arcsoc.exe");
            this.AddAttachToCommand(mcs, PkgCmdIDList.cmdidWhutAttachToArcMap, gop => gop.ShowAttachToArcMap, "arcmap.exe");
        }

        private void AddAttachToCommand(OleMenuCommandService mcs, uint commandId, Func<GeneralOptionsPage, bool> isVisible, params string[] programsToAttach)
        {
            OleMenuCommand menuItemCommand = new OleMenuCommand((sender, e) =>
            {
                DTE dte = (DTE)this.GetService(typeof(DTE));

                foreach (var process in dte.Debugger.LocalProcesses.Cast<Process>())
                {
                    if (programsToAttach.Any(programToAttach => process.Name.EndsWith(programToAttach, StringComparison.OrdinalIgnoreCase)))
                    {
                        process.Attach();
                    }
                }
            }, 
                new CommandID(GuidList.guidAttachToCmdSet, (int)commandId));

            menuItemCommand.BeforeQueryStatus += (s, e) => 
                menuItemCommand.Visible = isVisible((GeneralOptionsPage)this.GetDialogPage(typeof(GeneralOptionsPage)));

            mcs.AddCommand(menuItemCommand);
        }
    }
}
