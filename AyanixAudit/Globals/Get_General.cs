using AyanixAudit.Models;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Security.Cryptography;

namespace AyanixAudit.Globals
{
    public static class Get_WMI
    {
        private const string sDateTimeFormat = "yyyy-MM-dd HH:mm:ss";
        private const string sDateFormat = "yyyy-MM-dd";
        private const string sLinesss = " -------------------------------------------------------------------------------------------------------------------";


        public static PC_Info _pc = new PC_Info();

        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetPhysicallyInstalledSystemMemory(out long TotalMemoryInKilobytes);
  
        private static long Get_MemoryTotal()
        {
            long memKb;
            GetPhysicallyInstalledSystemMemory(out memKb);
            return (memKb / 1024); // MB
        }

        public static PC_Info Get_BoardInfo()
        {
            _pc = new PC_Info();

            try
            {
                ManagementClass wmi;

                wmi = new ManagementClass("Win32_BaseBoard");
                foreach (var m in wmi.GetInstances())
                {
                    _pc.Board_Maker = m["Manufacturer"].ToString();
                    _pc.Board_Model = m["Product"].ToString();
                    _pc.Board_Serial = m["SerialNumber"].ToString();
                }

                wmi = new ManagementClass("Win32_BIOS");
                foreach (var m in wmi.GetInstances())
                {
                    _pc.BIOS_Maker = m["Manufacturer"].ToString();
                    _pc.BIOS_Serial = m["SerialNumber"].ToString();
                    _pc.BIOS_Date = ManagementDateTimeConverter.ToDateTime(m["ReleaseDate"].ToString()).ToString(sDateTimeFormat);
                }

                wmi = new ManagementClass("Win32_Processor");
                foreach (var m in wmi.GetInstances())
                {
                    _pc.Processor_Name = m["Name"].ToString();
                    _pc.Processor_Cores = m["NumberOfCores"].ToString();
                    _pc.Processor_Logical = m["NumberOfLogicalProcessors"].ToString();
                }

                wmi = new ManagementClass("Win32_ComputerSystem");
                foreach (var m in wmi.GetInstances())
                {
                    _pc.System_Model = m["Model"].ToString();
                    _pc.System_RAM = (Get_MemoryTotal() / 1024).ToString();

                    _pc.OS_Hostname = Environment.MachineName;
                    _pc.OS_User = Environment.UserName;
                    _pc.OS_Domain = (string)m["Domain"];
                }

                wmi = new ManagementClass("Win32_OperatingSystem");
                foreach (var m in wmi.GetInstances())
                {
                    _pc.OS_Name = m["Caption"].ToString() + " (" + m["OSArchitecture"].ToString() + ")";
                    _pc.OS_Build = m["Version"].ToString();
                    _pc.OS_InstallDate = ManagementDateTimeConverter.ToDateTime(m["InstallDate"].ToString()).ToString(sDateTimeFormat);
                }

            }
            catch { }

            return _pc;
        }

        public static List<PC_Drive> Get_Volumes()
        {
            List<PC_Drive> _lst = new List<PC_Drive>();

            ManagementClass wmi;

            wmi = new ManagementClass("Win32_DiskDrive");
            foreach (var m in wmi.GetInstances())
            {
                if (m["Size"] != null)
                {
                    string sDevType = "";

                    if(m["Caption"].ToString().Length > 0)
                    {
                        if (m["Caption"].ToString().Contains("SSD")) 
                            sDevType = "SSD";
                    }

                    _lst.Add(new PC_Drive
                    {
                        Type = "Disk",
                        Index = Convert.ToInt32(m["Index"].ToString()),
                        Name = m["Caption"].ToString(),
                        Size_U64 = (ulong)m["Size"],
                        Size = Helper.ToSize((ulong)m["Size"]),
                        Partition = m["Partitions"] != null ? m["Partitions"].ToString() : "",
                        DevType = sDevType
                    });
                }
            }

            wmi = new ManagementClass("Win32_LogicalDisk");
            foreach (var m in wmi.GetInstances())
            {
                int iDType = Convert.ToInt32(m["DriveType"]);

                if (iDType <= 3 || iDType == 5)
                {
                    PC_Drive _drv = new PC_Drive();

                    string sVolume = "";

                    try
                    {
                        sVolume = m["VolumeName"] !=null ? m["VolumeName"].ToString() : "";
          
                        if (sVolume == "")
                            sVolume = m["Description"].ToString();

                        if (m["Caption"] != null)
                            sVolume += " (" + m["Caption"].ToString() + ")";
                    }
                    catch { }

                    _drv.Type = "Drive";
                    _drv.Name = sVolume;
                    _drv.Letter = m["Caption"].ToString().Length > 1 ? m["Caption"].ToString().Substring(0, 1) : m["Caption"].ToString();

                    if (m["Size"] != null)
                    {
                        _drv.FileSystem = m["FileSystem"].ToString();

                        _drv.Size_U64 = (ulong)m["Size"];
                        _drv.Size = Helper.ToSize((ulong)m["Size"]);
                        _drv.Free = Helper.ToSize((ulong)m["FreeSpace"]);
                        _drv.Used = Helper.ToSize(((ulong)m["Size"] - (ulong)m["FreeSpace"]));

                        _lst.Add(_drv);
                    }
                }

                if (iDType == 4 || iDType > 5)
                {
                    _lst.Add(new PC_Drive
                    {
                        Type = "Drive",
                        Name = m["ProviderName"].ToString() + " (" + m["Caption"].ToString() + ") ",
                        Letter = m["Caption"].ToString().Length > 1 ? m["Caption"].ToString().Substring(0, 1) : m["Caption"].ToString(),
                        FileSystem = "Network",
                        DevType = "NET",
                        Size = "",
                        Free = "",
                        Used = ""
                    });
                }
            }

            ManagementScope MScope = new ManagementScope(@"\\.\root\microsoft\windows\storage");
            MScope.Connect();
            ManagementObjectSearcher ObjSch;

            ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM MSFT_Partition"));
            foreach (ManagementObject m in ObjSch.Get())
            {
                foreach (PC_Drive d in _lst)
                {
                    if (d.Type == "Drive" && m["DriveLetter"].ToString() == d.Letter)
                    {
                        d.Index = Convert.ToInt32(m["DiskNumber"].ToString());
                    }
                }
            }

            ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM MSFT_PhysicalDisk"));
            foreach (ManagementObject m in ObjSch.Get()) //MSFT_PhysicalDisk - Determine HDD or SSD
            {
                foreach (PC_Drive d in _lst)
                {
                    if (m["DeviceID"].ToString() == d.Index.ToString() && d.FileSystem != "Network")
                    {
                        switch (m["MediaType"].ToString())
                        {
                            case "0":
                            case "3": d.DevType = "HDD"; break;
                            case "4": d.DevType = "SSD"; break;
                            case "5": d.DevType = "SCM"; break;
                                //default: d.DevType = ""; break;
                        }
                    }
                }
            }

            return _lst;
        }

        public static List<PC_Software> Get_Softwares2()
        {
            List<PC_Software> _lst = new List<PC_Software>();

            RegistryKey key;

            string sName = "", sVer = "";

            try
            {
                // Search in: LocalMachine_64
                key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall");
                foreach (String keyName in key.GetSubKeyNames())
                {
                    RegistryKey Rkey = key.OpenSubKey(keyName);

                    sName = Rkey.GetValue("DisplayName") != null ? Rkey.GetValue("DisplayName").ToString().Trim() : "";
                    sVer = Rkey.GetValue("DisplayVersion")!= null ? Rkey.GetValue("DisplayVersion").ToString().Trim() : "";

                    if (sName != "")
                    {
                        if (!sName.Contains("Security Update") && !sName.Contains("Update for Microsoft"))
                        {
                            PC_Software _sf = new PC_Software {
                                Name = sName,
                                Version = sVer
                            };

                            if(!_lst.Any(x=>x.Name == sName)) _lst.Add(_sf);
                        }
                    }
                }
            }
            catch { }

            try
            {
                // Search in: CurrentUser
                key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
                foreach (String keyName in key.GetSubKeyNames())
                {
                    RegistryKey Rkey = key.OpenSubKey(keyName);

                    sName = Rkey.GetValue("DisplayName") != null ? Rkey.GetValue("DisplayName").ToString().Trim() : "";
                    sVer = Rkey.GetValue("DisplayVersion")!= null ? Rkey.GetValue("DisplayVersion").ToString().Trim() : "";

                    if (sName != "")
                    {
                        if (!sName.Contains("Security Update") && !sName.Contains("Update for Microsoft"))
                        {
                            PC_Software _sf = new PC_Software {
                                Name = sName,
                                Version = sVer
                            };

                            if(!_lst.Any(x=>x.Name == sName)) _lst.Add(_sf);
                        }
                    }
                }
            }
            catch { }

            return _lst;
        }

        // ================================================================================================================================

        public static string Get_Board()
        {
            PC_Info _info = Get_BoardInfo();
            
            string sResult = Helper.Title("GENERAL INFORMATION");

            sResult += Helper.PadString("   System Model", 42) + " : " + _info.System_Model + Environment.NewLine;
            sResult += Helper.PadString("   Board Model", 42) + " : " + _info.Board_Model + Environment.NewLine;
            sResult += Helper.PadString("   Board Serial No", 42) + " : " + _info.Board_Serial + Environment.NewLine;
            sResult += Helper.PadString("   BIOS", 42) + " : " + _info.BIOS_Maker + " " + _info.BIOS_Serial + Environment.NewLine;
            sResult += Helper.PadString("   BIOS Date", 42) + " : " + _info.BIOS_Date + Environment.NewLine;
            sResult += Helper.PadString("   CPU", 42) + " : " + _info.Processor_Name + Environment.NewLine;
            sResult += Helper.PadString("   CPU Cores", 42) + " : " + _info.Processor_Cores + Environment.NewLine;
            sResult += Helper.PadString("   Logical Cores", 42) + " : " + _info.Processor_Logical + Environment.NewLine;
            sResult += Helper.PadString("   Memory", 42) + " : " + _info.System_RAM + " GB"+ Environment.NewLine;
            sResult += sLinesss + Environment.NewLine;
            sResult += Helper.PadString("   Hostname", 42) + " : " + _info.OS_Hostname + Environment.NewLine;
            sResult += Helper.PadString("   User Login", 42) + " : " + _info.OS_User + Environment.NewLine;
            sResult += Helper.PadString("   OS Domain", 42) + " : " + _info.OS_Domain + Environment.NewLine;
            sResult += Helper.PadString("   OS Installed", 42) + " : " + _info.OS_Name + " " + _info.OS_Build + Environment.NewLine;
            sResult += Helper.PadString("   OS Installed Date", 42) + " : " + _info.OS_InstallDate + Environment.NewLine;
            sResult += Environment.NewLine;

            frmMain._pcinfo = _info;

            return sResult;
        }

        public static string Get_RAM()
        {
            string sResult = "";

            try
            {
                ManagementClass wmi;

                sResult += sLinesss + Environment.NewLine;
                sResult += Helper.PadString("   DIMM SLOT ", 45) +
                           Helper.PadString("Maker", 32) +
                           Helper.PadString("Size ", 17) +
                           Helper.PadString("Speed (MHz) ", 15)  + Environment.NewLine;
                sResult += sLinesss + Environment.NewLine;

                wmi = new ManagementClass("Win32_PhysicalMemory");
                foreach  (var m in wmi.GetInstances())
                {
                    sResult += Helper.PadString("     " + m["DeviceLocator"].ToString().Trim() + " - " + m["BankLabel"].ToString(), 45);
                    sResult += Helper.PadString(m["Manufacturer"].ToString().Trim(), 32);
                    sResult += Helper.PadString(Helper.ToSize((ulong)m["Capacity"]), 17);
                    sResult += Helper.PadString(m["Speed"].ToString(), 15) + Environment.NewLine ;        
                }

                sResult +=Environment.NewLine;
            }
            catch { }

            return sResult;
        }

        public static string Get_Printers()
        {
            string sResult = "";

            ManagementClass wmi;

			sResult += Environment.NewLine;
			sResult += sLinesss + Environment.NewLine;
			sResult += Helper.PadString("   Printer Name", 62) +
						Helper.PadString("Port", 30) + Environment.NewLine;
			sResult += sLinesss + Environment.NewLine;

			wmi = new ManagementClass("Win32_Printer");
			foreach (var prn in wmi.GetInstances())
			{
				sResult += Helper.PadString("     " + prn["DriverName"].ToString(), 62);
				sResult += Helper.PadString(prn["PortName"].ToString(), 30);
				sResult += Environment.NewLine;
			}

            return sResult;
        }

        public static string Get_Graphics()
        {
            string sResult = "";

            try
            {
                ManagementClass wmi;

                sResult +=  sLinesss + Environment.NewLine;
                sResult +=  Helper.PadString("   DISPLAY GPU ", 62) +
                            Helper.PadString("Resolution ", 32) +
                            Helper.PadString("Version ", 20)  + Environment.NewLine;    
                sResult +=  sLinesss + Environment.NewLine;

                wmi = new ManagementClass("Win32_VideoController");
                foreach  (var m in wmi.GetInstances())
                {
                    sResult += Helper.PadString("     " + m["Caption"].ToString(), 62);

                    if (m["CurrentHorizontalResolution"] != null)
                    {
                        sResult += Helper.PadString(m["CurrentHorizontalResolution"].ToString() + " x " + 
                                                    m["CurrentVerticalResolution"].ToString()  + " @ " + 
                                                    m["CurrentRefreshRate"].ToString() + " Hz" ,32);

                        sResult += Helper.PadString(m["DriverVersion"].ToString(), 20);
                        sResult += Environment.NewLine ;        
                    }
                }

                sResult += Environment.NewLine;
            }
            catch { }

            return sResult;
        }

        public static string Get_Inputs()
        {
            string sResult = "";
            int iCount = 0;
            ManagementClass wmi;

            sResult += sLinesss + Environment.NewLine;
			sResult += Helper.PadString("   Input Device", 62) +
							Helper.PadString("ID", 30) + Environment.NewLine;
			sResult += sLinesss + Environment.NewLine;

			wmi = new ManagementClass("Win32_KeyBoard");
			foreach (var kyb in wmi.GetInstances())
			{
				sResult += Helper.PadString("     " + kyb["Description"].ToString(), 62);
				sResult += Helper.PadString(kyb["PNPDeviceId"].ToString(), 30);
				sResult += Environment.NewLine;
			}

            iCount += wmi.GetInstances().Count;

			wmi = new ManagementClass("Win32_PointingDevice");
			foreach (var mse in wmi.GetInstances())
			{
				sResult += Helper.PadString("     " + mse["Manufacturer"].ToString() + " " + mse["Caption"].ToString(), 62);
				sResult += Helper.PadString(mse["PNPDeviceId"].ToString(), 30);
				sResult += Environment.NewLine;
			}

             iCount += wmi.GetInstances().Count;


            return iCount == 0 ? "" : sResult;
        }


        public static string Get_NetworkAdapter()
        {
            string sResult = "";

			sResult += Environment.NewLine;
			sResult +=  sLinesss + Environment.NewLine;
			sResult += Helper.PadString("   Network Adapter ", 94) +
						Helper.PadString("MAC Address ", 30) + Environment.NewLine;
			sResult +=  sLinesss + Environment.NewLine;

            ManagementClass wmi;

            wmi = new ManagementClass("Win32_NetworkAdapter");
            foreach (var m in wmi.GetInstances())
            {
                if (m["Manufacturer"] != null)
                {
                    if (!m["Manufacturer"].ToString().Contains("Microsoft") &&
                        !m["Manufacturer"].ToString().Contains("VMware") &&
                        !m["ProductName"].ToString().Contains("Virtual"))
                    {
                        sResult += Helper.PadString("     " + m["ProductName"].ToString(), 94);
				        sResult += Helper.PadString(m["MACAddress"] != null? m["MACAddress"].ToString() : "-NA-", 30);
				        sResult += Environment.NewLine;
                    }
                }
            }

			sResult += Environment.NewLine;

            return sResult;
        }

        public static string Get_Network(ManagementScope MScope)
        {
            string sResult = "";

            ManagementObjectSearcher ObjSch;

            try
            {
                sResult += sLinesss + Environment.NewLine;
                sResult += Helper.PadString("   IP Address ", 45) +
                              Helper.PadString("Netmask ", 17) +
                              Helper.PadString("Gateway ", 32) +
                              Helper.PadString("MAC Address ", 18) + Environment.NewLine;
                sResult += sLinesss + Environment.NewLine;
              
                ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT IPAddress,IPSubnet,DefaultIPGateway,MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'TRUE'"));
                foreach (ManagementObject m in ObjSch.Get())
                {
                    string sIPGateway = "";

                    if(m["DefaultIPGateway"] != null)
                        sIPGateway = string.Join(", ", (string[])m["DefaultIPGateway"]);

                    string[] arr_IP = (string[])m["IPAddress"];
                    string[] arr_Sub = (string[])m["IPSubnet"];

                    for(int i = 0; i < arr_IP.Length;i++)
                    {
                        sResult += Helper.PadString("     " + arr_IP[i], 45) +
                                   Helper.PadString(arr_Sub[i], 17) +
                                   Helper.PadString(sIPGateway, 32) +
                                   Helper.PadString(m["MACAddress"].ToString(), 18) + Environment.NewLine;
                    }
                }

                sResult += Environment.NewLine;
            }
            catch { }

            return sResult;
        }


        public static string Get_Disk(List<PC_Drive> lst)
        {
            string sResult = "";
            var vLst = lst.OrderBy(x => x.Index).ToList();

            sResult += sLinesss + Environment.NewLine;
            sResult += Helper.PadString("   DISK ", 45) +
                        Helper.PadString("Size ", 17) +
                        Helper.PadString("Partitions ", 15) +
                        Helper.PadString("Type ", 15) + Environment.NewLine;
            sResult += sLinesss + Environment.NewLine;

            foreach (PC_Drive drv in vLst)
            {
                sResult += Helper.PadString("     [" + drv.Index + "] " + drv.Name, 45);
                sResult += Helper.PadString(drv.Size, 17);
                sResult += Helper.PadString(drv.Partition, 15);
                sResult += Helper.PadString(drv.DevType, 15) + Environment.NewLine;

                if (drv.Boot == 1)
                {
                    _pc.OS_DiskIndex = drv.Index;
                    _pc.OS_DiskSize = Helper.FormatSize(Convert.ToInt64(drv.Size_U64));
                    _pc.OS_DiskType = drv.DevType;
                }
            }

            return sResult;
        }

        public static string Get_Drives(List<PC_Drive> lst)
        {
            string sResult = "";
            var vLst1 = lst.OrderBy(x => x.Letter).ToList();

            sResult += Environment.NewLine;
            sResult += sLinesss + Environment.NewLine;
            sResult += Helper.PadString("   DRIVES ", 45) +
                        Helper.PadString("Used ", 17) +
                        Helper.PadString("Capacity ", 15) +
                        Helper.PadString("File System ", 15) + Environment.NewLine;
            sResult += sLinesss + Environment.NewLine;

            foreach (PC_Drive drv in vLst1)
            {
                if(drv.FileSystem == "Network"){
                    sResult += Helper.PadString("     " + drv.Name, 45);
                }
                else
                {
                    sResult += Helper.PadString("     [" + drv.Index + "] " + drv.Name, 45);
                }

                sResult += Helper.PadString(drv.Used, 17);
                sResult += Helper.PadString(drv.Size, 15);
                sResult += Helper.PadString(drv.FileSystem, 15) + Environment.NewLine;
            }
            
            sResult += Environment.NewLine;

            return sResult;
        }

        public static string Get_Softwares()
        {
            string sResult = "";

            List<PC_Software> _lst = Get_Softwares2().OrderBy(x => x.Name).ToList();

            sResult += sLinesss + Environment.NewLine;
            sResult += Helper.PadString("   INSTALLED SOFTWARE", 94) + "VERSION" + Environment.NewLine;
            sResult += sLinesss + Environment.NewLine;

            foreach (PC_Software sf in _lst )
            {
                sResult += Helper.PadString("     " + sf.Name , 94) + sf.Version + Environment.NewLine;
				//Helper.PadString( sf.Version , iVerLen + 3)
            }
            
            return sResult;
        }

        public static List<PC_Drive> Get_DrivesV2()
        {
            string sBookDisk = "";

			ManagementClass wmi = new ManagementClass("Win32_DiskPartition");
			foreach (var m in wmi.GetInstances())
			{
				if ((bool)m["Bootable"]) sBookDisk = m["DiskIndex"].ToString();
			}

            List<PC_Drive> _lst = new List<PC_Drive>();

            ConnectionOptions ConOpt = new ConnectionOptions { Impersonation = ImpersonationLevel.Impersonate};
			ManagementScope MScope = new ManagementScope("\\\\.\\root\\cimv2", ConOpt);
            MScope.Connect();

            ManagementObjectSearcher ObjSch = new ManagementObjectSearcher(MScope,new ObjectQuery("select * from Win32_DiskDrive"));
            foreach (ManagementObject d in ObjSch.Get())
            {
                int iDiskIndex = Convert.ToInt32(d["Index"].ToString());

                _lst.Add(new PC_Drive
                {
                    Type = "Disk",
                    Index = iDiskIndex,
                    Name = d["Caption"].ToString(),
                    Size_U64 = (ulong)d["Size"],
                    Size = Helper.ToSize((ulong)d["Size"]),
                    Partition = d["Partitions"].ToString(),

                    Boot = iDiskIndex.ToString() == sBookDisk ? 1 : 0,
                    DevType = d["Caption"].ToString().Contains("SSD") ? "SSD" : "HDD"
                });

                var PatQueryText = string.Format("associators of {{{0}}} where AssocClass = Win32_DiskDriveToDiskPartition", d.Path.RelativePath);
                var PatQuery = new ManagementObjectSearcher(PatQueryText);
                foreach (ManagementObject p in PatQuery.Get())
                {
                    var LDQueryText = string.Format("associators of {{{0}}} where AssocClass = Win32_LogicalDiskToPartition", p.Path.RelativePath);
                    var LDDriveQuery = new ManagementObjectSearcher(LDQueryText);

                    foreach (ManagementObject ld in LDDriveQuery.Get())
                    {
                        int iDType = Convert.ToInt32(ld["DriveType"]);

                        if (iDType <= 3 || iDType == 5)
                        {
                            PC_Drive _drv = new PC_Drive();
                            _drv.Index = iDiskIndex;
                            _drv.Type = "Drive";
                            _drv.Letter = ld["Name"].ToString() != ""? ld["Name"].ToString().Substring(0, 1) : "";

                            string sVolume = ld["VolumeName"].ToString() != "" ? ld["VolumeName"].ToString() : "";

							if (sVolume == "")
								sVolume = ld["Description"].ToString();

								sVolume += " (" + ld["Name"].ToString() + ")";
                            
                            _drv.Name = sVolume;
                            
                            if (ld["Size"] != null)
                            {
                                _drv.FileSystem = ld["FileSystem"].ToString();

                                _drv.Size_U64 = (ulong)ld["Size"];
                                _drv.Size = Helper.ToSize((ulong)ld["Size"]);
                                _drv.Free = Helper.ToSize((ulong)ld["FreeSpace"]);
                                _drv.Used = Helper.ToSize(((ulong)ld["Size"] - (ulong)ld["FreeSpace"]));
                                _lst.Add(_drv);
                            }
                        }

                        if (iDType == 4 || iDType > 5)
                        {
                            _lst.Add(new PC_Drive
                            {
                                Type = "Drive",
                                Name = ld["ProviderName"].ToString() + " (" + ld["Name"].ToString() + ") ",
                                Letter = ld["Name"].ToString() != ""? ld["Name"].ToString().Substring(0, 1) : "",
                                FileSystem = "Network",
                                DevType = "NET",
                                Size = "",
                                Free = "",
                                Used = ""
                            });
                        }

                    }
                }
            }

            return _lst;
        }

    }
}
