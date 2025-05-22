using System;
using System.Management;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.Data;
using System.Data.SqlClient;
using AyanixAudit.Models;
using System.Linq;
using System.Collections.Generic;
using System.Net.NetworkInformation;

namespace AyanixAudit
{
    public partial class frmMain : Form
    {
		public static PC_Info _pcinfo;

        public frmMain()
        {
            this.DoubleBuffered = true;
            InitializeComponent();

			_pcinfo = new PC_Info();
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
				string sFileName = Environment.MachineName + "_" + Environment.UserName;

                ConnectionOptions ConOpt = new ConnectionOptions { Impersonation = ImpersonationLevel.Impersonate};
				ManagementScope MScope = new ManagementScope("\\\\.\\root\\cimv2", ConOpt);
                MScope.Connect();

				ManagementClass wmi;

                Delegate_Msg(" * Getting Board Information...");
                sResult += Globals.Get_WMI.Get_Board();

                Delegate_Msg(" * Getting RAM Information...");
                sResult += Globals.Get_WMI.Get_RAM();

                Delegate_Msg(" * Getting Display Information...");
                sResult += Globals.Get_WMI.Get_Graphics();

                Delegate_Msg(" * Getting Drive Information...");

				try
				{
					List<PC_Drive> lst = Globals.Get_WMI.Get_DrivesV2();

					sResult += Globals.Get_WMI.Get_Disk(lst.Where(d => d.Type == "Disk").ToList());

					sResult += Globals.Get_WMI.Get_Drives(lst.Where(d => d.Type == "Drive").ToList());
				}
				catch (Exception ex)
				{
					Delegate_Msg(" Error : " + ex.Message);
					File.WriteAllText(Application.StartupPath + "\\" + sFileName + "_errors.txt", txtStatus.Text);
				}

                Delegate_Msg(" * Getting Input Information...");

				try{
					sResult += Globals.Get_WMI.Get_Inputs();
				}
				catch (Exception ex)
				{
					Delegate_Msg(" Error : " + ex.Message);
					File.WriteAllText(Application.StartupPath + "\\" + sFileName + "_errors.txt", txtStatus.Text);
				}

				Delegate_Msg(" * Getting Printers...");

				try{
					sResult += Globals.Get_WMI.Get_Printers();
				}
				catch (Exception ex)
				{
					Delegate_Msg(" Error : " + ex.Message);
					File.WriteAllText(Application.StartupPath + "\\" + sFileName + "_errors.txt", txtStatus.Text);
				}

				Delegate_Msg(" * Getting Network Adapter...");

				try{
					sResult += Globals.Get_WMI.Get_NetworkAdapter();
				}
				catch (Exception ex)
				{
					Delegate_Msg(" Error : " + ex.Message);
					File.WriteAllText(Application.StartupPath + "\\" + sFileName + "_errors.txt", txtStatus.Text);
				}

                Delegate_Msg(" * Getting Network Connections...");
                sResult += Globals.Get_WMI.Get_Network(MScope);

                Delegate_Msg(" * Getting Software Installed...");
                sResult += Globals.Get_WMI.Get_Softwares();

				_pcinfo = Globals.Get_WMI._pc;

				Delegate_Msg(" * Saving Data...");

				try
				{
                    File.WriteAllText(Application.StartupPath + "\\" + sFileName + ".txt" , sResult);
                }
                catch(Exception){}


				Delegate_Msg(" * Checking Database connection for 192.168.121.210...");

				if (PingHost("192.168.121.210"))
				{
					if (SQL.Check_DB(SQL._DB1))
					{
						Delegate_Msg(" * Database connection is OK.");

						bool bUpload = UploadToSQL(_pcinfo, sResult, SQL._DB1);

						if (bUpload)
						{
							Delegate_Msg(" * Update Completed.");
						}
					}
					else
					{
						Delegate_Msg(" * Database connection failed for 192.168.121.210 .");
					}
				}
				else if (PingHost("10.0.1.29"))
				{
					if (SQL.Check_DB(SQL._DB2))
					{
						Delegate_Msg(" * Database connection is OK.");

						bool bUpload = UploadToSQL(_pcinfo, sResult, SQL._DB2);

						if (bUpload)
						{
							Delegate_Msg(" * Update Completed.");
						}
					}
					else
					{
						Delegate_Msg(" * Database connection failed for 10.0.1.29.");
					}
				}
				else
				{
					Delegate_Msg(" * Database connection failed.");
					Thread.Sleep(1000);
				}

				Thread.Sleep(1000);

				Delegate_Msg(sResult,true);
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




        private bool UploadToSQL(PC_Info _pc,string sRes, string sSQL)
        {
			bool bUpdate = false;

			using (SqlConnection SC = new SqlConnection(sSQL))
			{
				SC.Open();

				bool bExist = false;

				string sID = "";

				DataTable DT_tmp = new DataTable();

				SqlDataAdapter SDA = new SqlDataAdapter("SELECT * FROM Stocks_Devices ", SC);
				SDA.Fill(DT_tmp);

				foreach (DataRow dr in DT_tmp.Rows)
				{
					if (_pcinfo.OS_Hostname == dr["Dev_AssetNo"].ToString() ||
						_pcinfo.OS_Hostname == dr["Dev_OS_Hostname"].ToString())
					{
						bExist = true;

						sID = dr["id"].ToString();
					}
				}

				if (bExist)
				{
					SDA = new SqlDataAdapter("SELECT TOP 1 * FROM Stocks_Devices WHERE ID = " + sID, SC);

					DataTable DT_DevInfo = new DataTable();
					SDA.Fill(DT_DevInfo);

					if (DT_DevInfo.Rows.Count > 0)
					{
						//DT_DevInfo.Rows[0]["Dev_Type"] = "DESKTOP";
						DT_DevInfo.Rows[0]["Dev_ModelNo"] = _pc.System_Model;
						
						DT_DevInfo.Rows[0]["Dev_Maker"] = _pc.Board_Maker;
						DT_DevInfo.Rows[0]["Dev_SerialNo"] = _pc.Board_Serial;

						DT_DevInfo.Rows[0]["Dev_CPU"] = _pc.Processor_Name;
						DT_DevInfo.Rows[0]["Dev_CPU_Cores"] = _pc.Processor_Cores;
						DT_DevInfo.Rows[0]["Dev_RAM"] = _pc.System_RAM;

						DT_DevInfo.Rows[0]["Dev_AssetNo"] = _pc.OS_Hostname;
						DT_DevInfo.Rows[0]["Dev_OS_Hostname"] = _pc.OS_Hostname;

						DT_DevInfo.Rows[0]["Dev_OS_Name"] = _pc.OS_Name + " " + _pc.OS_Build;
						DT_DevInfo.Rows[0]["Dev_OS_Date"] = _pc.OS_InstallDate;
						DT_DevInfo.Rows[0]["Dev_OS_Domain"] = _pc.OS_Domain;

						DT_DevInfo.Rows[0]["Dev_NetProfile"] = "STD";
						DT_DevInfo.Rows[0]["Dev_Status"] = "OK-INUSE";

						if(_pc.OS_DiskType != "")
						{
							DT_DevInfo.Rows[0]["Dev_Storage_Type"] = _pc.OS_DiskType + "-SATA";
						}
						
						if(_pc.OS_DiskSize != "")
						{
							DT_DevInfo.Rows[0]["Dev_Storage_Size"] = _pc.OS_DiskSize;
						}

						DT_DevInfo.Rows[0]["Dev_User"] = _pc.OS_User;

						DT_DevInfo.Rows[0]["Updated_By"] = "0000";
						DT_DevInfo.Rows[0]["Updated_Date"] = DateTime.Now;

						using (new SqlCommandBuilder(SDA))
						{
							SDA.Update(DT_DevInfo);  // UPDATE DATA TO SQL 
							bUpdate = true;
						}
					}
				}
				else
				{
					DataTable DT_DevInfo = new DataTable();

					SDA = new SqlDataAdapter("SELECT TOP 1 * FROM Stocks_Devices ", SC);
					SDA.Fill(DT_DevInfo);

					DataRow dr_new = DT_DevInfo.NewRow();

					dr_new["Dev_Type"] = "DESKTOP";

					dr_new["Dev_AssetNo"] = _pc.OS_Hostname;
					dr_new["Dev_Maker"] = _pc.Board_Maker;
					dr_new["Dev_ModelNo"] = _pc.System_Model;
					dr_new["Dev_SerialNo"] = _pc.Board_Serial;

					dr_new["Dev_CPU"] = _pc.Processor_Name;
					dr_new["Dev_CPU_Cores"] = _pc.Processor_Cores;
					dr_new["Dev_RAM"] = _pc.System_RAM;

					dr_new["Dev_OS_Hostname"] = _pc.OS_Hostname;
					dr_new["Dev_OS_Name"] = _pc.OS_Name + " " + _pc.OS_Build;
					dr_new["Dev_OS_Date"] = _pc.OS_InstallDate;
					dr_new["Dev_OS_Domain"] = _pc.OS_Domain;

					dr_new["Dev_NetProfile"] = "STD";

					dr_new["Dev_Storage_Type"] = _pc.OS_DiskType != "" ? _pc.OS_DiskType + "-SATA" : "HDD-SATA";
					dr_new["Dev_Storage_Size"] = _pc.OS_DiskSize != "" ? _pc.OS_DiskSize :  "500 GB";

					dr_new["Dev_Status"] = "OK-INUSE";
					dr_new["Dev_Location"] = "";
					dr_new["Dev_Note"] = "";
					dr_new["Dev_User"] = _pc.OS_User;

					dr_new["Added_By"] = "0000";
					dr_new["Added_Date"] = DateTime.Now;

					DT_DevInfo.Rows.Add(dr_new);

					using (new SqlCommandBuilder(SDA))
					{
						SDA.Update(DT_DevInfo);  // UPDATE DATA TO SQL 
						bUpdate = true;
					}
				}


				// ------------------------------------------------------------------------------------------------------
				DataTable DT_Info = new DataTable();

				SDA = new SqlDataAdapter("SELECT * FROM Stocks_Devices_Info WHERE Dev_AssetNo = '" + _pc.OS_Hostname + "' ", SC);
				SDA.Fill(DT_Info);

				if (DT_Info.Rows.Count > 0)
				{
					DT_Info.Rows[0]["Dev_Info"] = sRes;
					DT_Info.Rows[0]["Date_Added"] = DateTime.Now;
					using (new SqlCommandBuilder(SDA))
					{
						SDA.Update(DT_Info);  // UPDATE DATA TO SQL 
					}
				}
				else
				{
					DataRow dr_new = DT_Info.NewRow();
					dr_new["Dev_AssetNo"] = _pc.OS_Hostname;
					dr_new["Dev_Info"] = sRes;
					dr_new["Date_Added"] = DateTime.Now;
					DT_Info.Rows.Add(dr_new);
					using (new SqlCommandBuilder(SDA))
					{
						SDA.Update(DT_Info);  // UPDATE DATA TO SQL 
					}
				}

			}



			return bUpdate;
        }

		public static bool PingHost(string nameOrAddress)
		{
			try
			{
				using (Ping pinger = new Ping())
				{
					PingReply reply = pinger.Send(nameOrAddress);
					return reply.Status == IPStatus.Success;
				}
			}
			catch (PingException)
			{
				return false;
			}
		}

    }
}
