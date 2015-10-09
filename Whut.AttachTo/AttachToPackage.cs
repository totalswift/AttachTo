using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
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
            AddAttachToCommand(mcs, ProgramsList
                .IIS, gop => gop.ShowAttachToIIS, "w3wp.exe");
            AddAttachToCommand(mcs, ProgramsList.IISExpress, gop => gop.ShowAttachToIISExpress, "iisexpress.exe");
            AddAttachToCommand(mcs, ProgramsList.NUnit, gop => gop.ShowAttachToNUnit, "nunit-agent.exe", "nunit.exe", "nunit-console.exe", 
                "nunit-agent-x86.exe", "nunit-x86.exe", "nunit-console-x86.exe");
            AddAttachToCommand(mcs, ProgramsList.ArcSOC, gop => gop.ShowAttachToArcSOC, "arcsoc.exe");
            AddAttachToCommand(mcs, ProgramsList.ArcMap, gop => gop.ShowAttachToArcMap, "arcmap.exe");
            AddAttachToCommand(mcs, ProgramsList.ArcCatalog, gop => gop.ShowAttachToArcCatalog, "arccatalog.exe");
            AddAttachToCommand(mcs, ProgramsList.DBTool, gop => gop.ShowAttachToDBTool, "dbtool.exe");
        }

        private void AddAttachToCommand(OleMenuCommandService mcs, uint commandId, Func<GeneralOptionsPage, bool> isVisible, params string[] programsToAttach)
        {
            var menuItemCommand = new OleMenuCommand((sender, e) =>
            {
                var dte = (DTE)GetService(typeof(DTE));
                foreach (var process in dte.Debugger.LocalProcesses.Cast<Process>())
                {
                    if (programsToAttach.Any(
                        p => process.Name.EndsWith(p, StringComparison.OrdinalIgnoreCase)))
                    {
                        process.Attach();
                    }
                }
            }, 
                new CommandID(GuidList.guidAttachToCmdSet, (int)commandId));

            menuItemCommand.BeforeQueryStatus += (s, e) => 
                menuItemCommand.Visible = isVisible((GeneralOptionsPage)GetDialogPage(typeof(GeneralOptionsPage)));

            mcs.AddCommand(menuItemCommand);
        }
    }
}
