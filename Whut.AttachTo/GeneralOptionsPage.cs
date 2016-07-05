using System.ComponentModel;
using Microsoft.VisualStudio.Shell;

namespace AttachTo
{
    public class GeneralOptionsPage : DialogPage
    {
        public GeneralOptionsPage()
        {
            ShowAttachToIIS = false;
            ShowAttachToIISExpress = false;
            ShowAttachToNUnit = false;
            ShowAttachToDotNet = true;
        }

        [Category("General")]
        [DisplayName("Show 'Attach to IIS' command")]
        [Description("Show 'Attach to IIS' command in Tools menu.")]
        public bool ShowAttachToIIS { get; set; }

        [Category("General")]
        [DisplayName("Show 'Attach to IIS Express command")]
        [Description("Show 'Attach to IIS Express command in Tools menu.")]
        public bool ShowAttachToIISExpress { get; set; }

        [Category("General")]
        [DisplayName("Show 'Attach to NUnit' command")]
        [Description("Show 'Attach to NUnit' command in Tools menu.")]
        public bool ShowAttachToNUnit { get; set; }

        [Category("General")]
        [DisplayName("Show 'Attach to DotNet.exe' command")]
        [Description("Show 'Attach to DotNet.exe' command in Tools menu.")]
        public bool ShowAttachToDotNet { get; set; }
    }
}
