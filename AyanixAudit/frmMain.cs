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

                ConnectionOptions ConOpt = new ConnectionOptions { Impersonation = ImpersonationLevel.Impersonate};
				ManagementScope MScope = new ManagementScope("\\\\.\\root\\cimv2", ConOpt);
                MScope.Connect();

                Delegate_Msg(" * Getting Board Information...");
                sResult += Globals.Get_WMI.Get_Board();

                //Delegate_Msg(" * Getting Users Information...");
                //sResult += Globals.Get_WMI.Get_Accounts(MScope);

                Delegate_Msg(" * Getting RAM Information...");
                sResult += Globals.Get_WMI.Get_RAM();

                Delegate_Msg(" * Getting Display Information...");
                sResult += Globals.Get_WMI.Get_Graphics();


                Delegate_Msg(" * Getting Drive Information...");

				try
				{
					//List<PC_Drive> lst_drives = new List<PC_Drive>();

					//ManagementClass wmi;

					//wmi = new ManagementClass("Win32_DiskDrive");
					//foreach (var m in wmi.GetInstances())
					//{
					//	Delegate_Msg(" * Drives: Adding Disk " + m["Caption"].ToString());

					//	if (m["Size"] != null)
					//	{
					//		lst_drives.Add(new PC_Drive
					//		{
					//			Type = "Disk",
					//			Index = Convert.ToInt32(m["Index"].ToString()),
					//			Name = m["Caption"].ToString(),
					//			Size_U64 = (ulong)m["Size"],
					//			Size = Helper.ToSize((ulong)m["Size"]),
					//			Partition = m["Partitions"]!= null ? m["Partitions"].ToString() : ""
					//		});
					//	}
					//}

					//wmi = new ManagementClass("Win32_LogicalDisk");
					//foreach (var m in wmi.GetInstances())
					//{
					//	int iDType = Convert.ToInt32(m["DriveType"]);

					//	Delegate_Msg(" * Drives: Adding Drives " + m["Name"].ToString());


					//	if (iDType <= 3 || iDType == 5)
					//	{
					//		PC_Drive _drv = new PC_Drive();

					//		string sName = "";

					//		try
					//		{
					//			sName = (string)m["VolumeName"] ?? "";

					//			if (sName == "")
					//				sName = m["Description"].ToString();

					//			if (m["Name"] != null)
					//				sName += " (" + m["Name"].ToString() + ")";
					//		}
					//		catch { }

					//		//if (m["Caption"].ToString() != "")
					//		//    sName += " (" + m["Caption"].ToString() + ")";

					//		_drv.Type = "Drive";
					//		_drv.Name = sName;
					//		_drv.Letter = m["Name"].ToString().Length > 1 ? m["Name"].ToString().Substring(0, 1) : m["Name"].ToString();

					//		if (m["Size"] != null)
					//		{
					//			_drv.FileSystem = m["FileSystem"].ToString();

					//			_drv.Size_U64 = (ulong)m["Size"];
					//			_drv.Size = Helper.ToSize((ulong)m["Size"]);
					//			_drv.Free = Helper.ToSize((ulong)m["FreeSpace"]);
					//			_drv.Used = Helper.ToSize(((ulong)m["Size"] - (ulong)m["FreeSpace"]));

					//			lst_drives.Add(_drv);
					//		}
					//	}

					//	if (iDType == 4 || iDType > 5)
					//	{
					//		lst_drives.Add(new PC_Drive
					//		{
					//			Type = "Drive",
					//			Name = m["ProviderName"].ToString() + " (" + m["Caption"].ToString() + ") ",
					//			Letter = m["Caption"].ToString().Length > 1 ? m["Caption"].ToString().Substring(0, 1) : m["Caption"].ToString(),
					//			FileSystem = "Network",
					//			DevType = "NET",
					//			Size = "",
					//			Free = "",
					//			Used = ""
					//		});
					//	}
					//}

					//ManagementScope MScope_MSFT = new ManagementScope(@"\\.\root\microsoft\windows\storage");
					//MScope_MSFT.Connect();
					//ManagementObjectSearcher ObjSch;

					//Delegate_Msg(" * Drives: Setting Disk to Drive ...");

					//ObjSch = new ManagementObjectSearcher(MScope_MSFT, new ObjectQuery("SELECT * FROM MSFT_Partition"));
					//foreach (ManagementObject m in ObjSch.Get())
					//{
					//	foreach (PC_Drive d in lst_drives)
					//	{
					//		if (d.Type == "Drive" && m["DriveLetter"].ToString() == d.Letter)
					//		{
					//			d.Index = Convert.ToInt32(m["DiskNumber"].ToString());
					//		}
					//	}
					//}

					//Delegate_Msg(" * Drives: Setting media type ...");

					//ObjSch = new ManagementObjectSearcher(MScope_MSFT, new ObjectQuery("SELECT * FROM MSFT_PhysicalDisk"));
					//foreach (ManagementObject m in ObjSch.Get()) //MSFT_PhysicalDisk - Determine HDD or SSD
					//{
					//	foreach (PC_Drive d in lst_drives)
					//	{
					//		if (m["DeviceID"].ToString() == d.Index.ToString() && d.FileSystem != "Network")
					//		{
					//			switch (m["MediaType"].ToString())
					//			{
					//				case "0":
					//				case "3": d.DevType = "HDD"; break;
					//				case "4": d.DevType = "SSD"; break;
					//				case "5": d.DevType = "SCM"; break;
					//					//default: d.DevType = ""; break;
					//			}
					//		}
					//	}
					//}


					List<PC_Drive> lst_drives = Globals.Get_WMI.Get_Volumes();
					sResult += Globals.Get_WMI.Get_Drives(lst_drives);



					//Delegate_Msg(" * Drives: Analyzing Disk...");

					//var vLst0 = lst_drives.Where(t => t.Type == "Disk").OrderBy(x => x.Index).ToList();
					//var vLst1 = lst_drives.Where(t => t.Type == "Drive").OrderBy(x => x.Letter).ToList();

					//var vOS_Drv = lst_drives.Where(t => t.Letter == "C").FirstOrDefault();
					//if (vOS_Drv != null)
					//{
					//	_pcinfo.OS_DiskIndex = vOS_Drv.Index;
					//}

					//sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;
					//sResult += Helper.PadString("   DISK ", 45) +
					//			Helper.PadString("Size ", 17) +
					//			Helper.PadString("Partitions ", 15) +
					//			Helper.PadString("Type ", 15) + Environment.NewLine;
					//sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;

					//foreach (PC_Drive drv in vLst0)
					//{
					//	Delegate_Msg(" * Drives: Disk " + drv.Index);

					//	sResult += Helper.PadString("     [" + drv.Index + "] " + drv.Name, 45);
					//	//sResult += Helper.PadString("     " + drv.Name, 45);
					//	sResult += Helper.PadString(drv.Size ?? "", 17);
					//	sResult += Helper.PadString(drv.Partition ?? "", 15);
					//	sResult += Helper.PadString(drv.DevType ?? "", 15) + Environment.NewLine;

					//	if (_pcinfo.OS_DiskIndex == drv.Index)
					//	{
					//		_pcinfo.OS_DiskType = drv.DevType ?? "";
					//		_pcinfo.OS_DiskSize = Helper.FormatSize(Convert.ToInt64(drv.Size_U64));
					//	}
					//}

					//sResult += Environment.NewLine;


					//Delegate_Msg(" * Drives: Analyzing Drives...");

					//sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;
					//sResult += Helper.PadString("   DRIVES ", 45) +
					//			Helper.PadString("Used ", 17) +
					//			Helper.PadString("Capacity ", 15) +
					//			Helper.PadString("File System ", 15) + Environment.NewLine;
					//sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;

					//foreach (PC_Drive drv in vLst1)
					//{
					//	Delegate_Msg(" * Drives: Drive " + drv.Letter);

					//	if (drv.FileSystem == "Network")
					//	{
					//		sResult += Helper.PadString("     " + drv.Name, 45);
					//	}
					//	else
					//	{
					//		sResult += Helper.PadString("     [" + drv.Index + "] " + drv.Name ?? "", 45);
					//		sResult += Helper.PadString(drv.Used ?? "", 17);
					//		sResult += Helper.PadString(drv.Size ?? "", 15);
					//	}

					//	sResult += Helper.PadString(drv.FileSystem, 15) + Environment.NewLine;
					//}

					//sResult += Environment.NewLine;

				}
				catch (Exception ex)
				{
					Delegate_Msg(" Error : " + ex.Message);

					File.WriteAllText(Application.StartupPath + "\\" + Environment.MachineName + "_errors.txt", txtStatus.Text);
				}

                Delegate_Msg(" * Getting Input Information...");

				try
				{
					//List<PC_Devices> lst_devices = new List<PC_Devices>();
					ManagementClass wmi;

					sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;
					sResult += Helper.PadString("   Input Device", 62) +
								   Helper.PadString("ID", 30) + Environment.NewLine;
					sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;

					wmi = new ManagementClass("Win32_KeyBoard");
					foreach (var kyb in wmi.GetInstances())
					{
						sResult += Helper.PadString("     " + kyb["Description"].ToString(), 62);
						sResult += Helper.PadString(kyb["PNPDeviceId"].ToString(), 30);
						sResult += Environment.NewLine;
					}

					wmi = new ManagementClass("Win32_PointingDevice");
					foreach (var mse in wmi.GetInstances())
					{
						sResult += Helper.PadString("     " + mse["Manufacturer"].ToString() + " " + mse["Caption"].ToString(), 62);
						sResult += Helper.PadString(mse["PNPDeviceId"].ToString(), 30);
						sResult += Environment.NewLine;
					}
				}
				catch (Exception ex)
				{
					Delegate_Msg(" Error : " + ex.Message);

					File.WriteAllText(Application.StartupPath + "\\" + Environment.MachineName + "_errors.txt", txtStatus.Text);
				}


				Delegate_Msg(" * Getting Printers...");

				try
				{
					ManagementClass wmi;

					sResult += Environment.NewLine;
					sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;
					sResult += Helper.PadString("   Printer Name", 62) +
							   Helper.PadString("Port", 30) + Environment.NewLine;
					sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;

					wmi = new ManagementClass("Win32_Printer");
					foreach (var prn in wmi.GetInstances())
					{
						sResult += Helper.PadString("     " + prn["DriverName"].ToString(), 62);
						sResult += Helper.PadString(prn["PortName"].ToString(), 30);
						sResult += Environment.NewLine;
					}
				}
				catch (Exception ex)
				{
					Delegate_Msg(" Error : " + ex.Message);
					File.WriteAllText(Application.StartupPath + "\\" + Environment.MachineName + "_errors.txt", txtStatus.Text);
				}



				Delegate_Msg(" * Getting Network Adapter...");

				try
				{
					List<PC_Devices> lst_devices = Globals.Get_WMI.Get_NetAdapters();

					sResult += Environment.NewLine;
					sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;
					sResult += Helper.PadString("   Network Adapter ", 62) +
								Helper.PadString("MAC Address ", 30) + Environment.NewLine;
					sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;

					foreach (PC_Devices dev in lst_devices)
					{
						sResult += Helper.PadString("     " + dev.Name, 62);
						sResult += Helper.PadString(dev.DevID, 30);
						sResult += Environment.NewLine;
					}

					sResult += Environment.NewLine;
				}
				catch (Exception ex)
				{
					Delegate_Msg(" Error : " + ex.Message);

					File.WriteAllText(Application.StartupPath + "\\" + Environment.MachineName + "_errors.txt", txtStatus.Text);
				}

                Delegate_Msg(" * Getting Network Connections...");
                sResult += Globals.Get_WMI.Get_Network(MScope);

                Delegate_Msg(" * Getting Software Installed...");
                sResult += Globals.Get_WMI.Get_Softwares();

				_pcinfo = Globals.Get_WMI._pc;

				Delegate_Msg(" * Saving Data...");
				try{
                    File.WriteAllText(Application.StartupPath + "\\" + Environment.MachineName + ".txt" , sResult);
                }
                catch(Exception){}


				Delegate_Msg(" * Checking Database connection...");

				if (PingHost("192.168.121.210"))
				{
					if (SQL.Check_DB())
					{
						Delegate_Msg(" * Database connection is OK.");

						bool bUpload = UploadToSQL(_pcinfo, sResult);

						if (bUpload)
						{
							Delegate_Msg(" * Update Completed.");
						}
					}
					else
					{
						Delegate_Msg(" * Database connection failed.");
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




        private bool UploadToSQL(PC_Info _pc,string sRes)
        {
			bool bUpdate = false;

			using (SqlConnection SC = new SqlConnection(SQL._NPCIT))
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
