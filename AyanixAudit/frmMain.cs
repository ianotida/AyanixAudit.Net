using System;
using System.Management;
using System.Windows.Forms;
using System.IO;
using System.Threading;

namespace AyanixAudit
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            this.DoubleBuffered = true;
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            txtStatus.Text = " " + Environment.NewLine +
                             " Note : This App may freeze for a moment while collecting data." + Environment.NewLine  + Environment.NewLine  +
                             " Starting scan..." + Environment.NewLine  + Environment.NewLine  +
                             " Please wait....       " + Environment.NewLine  + Environment.NewLine;
            Application.DoEvents();

            Threading_Task();
        }

        private void Threading_Task()
        {
            Thread T = new Thread(() =>
            {
                string sResult = "";

                ManagementScope MScope;
                ConnectionOptions ConnecOpt = new ConnectionOptions();

                ConnecOpt.Impersonation = ImpersonationLevel.Impersonate;

                MScope = new ManagementScope("\\\\.\\root\\cimv2", ConnecOpt);
                MScope.Connect();

                Delegate_Msg(" * Getting Board Information...");
                sResult += Globals.Get_WMI.Get_Board(MScope);

                Delegate_Msg(" * Getting Users Information...");
                sResult += Globals.Get_WMI.Get_Accounts(MScope);

                Delegate_Msg(" * Getting Display Information...");
                sResult += Globals.Get_WMI.Get_Graphics(MScope);

                Delegate_Msg(" * Getting Drive Information...");
                sResult += Globals.Get_WMI.Get_Drives(MScope);

                Delegate_Msg(" * Getting Input Information...");
                sResult += Globals.Get_WMI.Get_HID(MScope);

                Delegate_Msg(" * Getting Printer Information....");
                sResult += Globals.Get_WMI.Get_Printers(MScope);

                Delegate_Msg(" * Getting Network Information...");
                sResult += Globals.Get_WMI.Get_NetAdapters(MScope);

                Delegate_Msg(" * Getting Software Installed...");
                sResult += Globals.Get_WMI.Get_Softwares();

                Delegate_Msg(" * Saving Data...");
                Thread.Sleep(200);

                Delegate_Msg(sResult,true);

                try
                {
                    File.WriteAllText(Application.StartupPath + "\\" + Environment.MachineName + "_" + Environment.UserName + ".txt" , sResult);
                }
                catch(Exception x)
                {
                    MessageBox.Show(x.Message, "Error Saving..", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            });

            T.Start();
        }

        private void Delegate_Msg(string sMsg, bool bOverwrite = false)
        {
            this.Invoke((MethodInvoker)delegate ()
            {
                if (!bOverwrite)
                {
                    txtStatus.Text += sMsg + Environment.NewLine;
                }
                else
                {
                    txtStatus.Text = sMsg + Environment.NewLine;
                }

                Application.DoEvents();
            });
        }

    }
}
