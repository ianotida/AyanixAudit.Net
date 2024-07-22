using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Management;
using System.Runtime.InteropServices;
using System.Data;
using System.Linq;
using System.Net;
using AyanixAudit.Models;
using System.Runtime.Remoting.Contexts;

namespace AyanixAudit.Globals
{
    public static class Get_WMI
    {
        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetPhysicallyInstalledSystemMemory(out long TotalMemoryInKilobytes);
  
        public static PCInfo Get_BoardInfo(ManagementScope MScope)
        {
            PCInfo _info = new PCInfo();

            ManagementObjectSearcher ObjSch;

            try
            {
                ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM Win32_BaseBoard"));
                foreach (ManagementObject m in ObjSch.Get())
                {
                    _info.Board_Model = m["Manufacturer"].ToString() + " " + m["Product"].ToString();
                    _info.Board_Serial = m["SerialNumber"].ToString();
                }

                ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM Win32_BIOS"));
                foreach (ManagementObject m in ObjSch.Get())
                {
                    _info.BIOS_Maker = m["Manufacturer"].ToString();
                    _info.BIOS_Serial = m["SerialNumber"].ToString();
                    _info.BIOS_Date = ManagementDateTimeConverter.ToDateTime(m["ReleaseDate"].ToString()).ToString("ddd, MMM dd yyyy hh:mm:ss tt");
                }

                ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM Win32_Processor"));
                foreach (ManagementObject m in ObjSch.Get())
                {
                    _info.Processor_Name = m["Name"].ToString();
                    _info.Processor_Cores = m["NumberOfCores"].ToString();
                    _info.Processor_Logical = m["NumberOfLogicalProcessors"].ToString();
                }

                ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM Win32_ComputerSystem"));
                foreach (ManagementObject m in ObjSch.Get())
                {
                    _info.System_Model = m["Model"].ToString();
                    _info.System_RAM = (Get_MemoryTotal() / 1024) + " GB";

                    _info.OS_Hostname = Environment.MachineName;
                    _info.OS_User = Environment.UserName;
                    _info.OS_Domain = m["Domain"].ToString();
                }

                ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM Win32_OperatingSystem"));
                foreach (ManagementObject m in ObjSch.Get())
                {
                    _info.OS_Name = m["Caption"].ToString() + " (" + m["OSArchitecture"].ToString() + ")";
                    _info.OS_Build = m["Version"].ToString();
                    _info.OS_InstallDate = ManagementDateTimeConverter.ToDateTime(m["InstallDate"].ToString()).ToString("MMM dd yyyy hh:mm:ss tt");
                }

            }
            catch { }

            return _info;
        }

        public static List<PC_Devices> Get_Devices()
        {
            List<PC_Devices> _lst = new List<PC_Devices>();

            ManagementClass wmi;

            try
            {
                wmi = new ManagementClass("Win32_KeyBoard");
                foreach (var keyboard in wmi.GetInstances())
                {
                    _lst.Add(new PC_Devices{
                        Type = "Keyboard",
                        Name = (string)keyboard["Description"],
                        DeviceID = (string)keyboard["PNPDeviceId"]
                    });
                }

                wmi = new ManagementClass("Win32_PointingDevice");
                foreach (var mouse in wmi.GetInstances())
                {
                     _lst.Add(new PC_Devices{
                        Type = "Mouse",
                        Name = (string)mouse["Manufacturer"] + " " + (string)mouse["Caption"],
                        DeviceID = (string)mouse["PNPDeviceId"]
                    });
                }

                wmi = new ManagementClass("Win32_Printer");
                foreach  (var m in wmi.GetInstances())
                {
                    if( !m["DriverName"].ToString().Contains("Microsoft"))
                    {
                        _lst.Add(new PC_Devices{
                            Type = "Printer",
                            Name = (string)m["DriverName"],
                            DeviceID = (string)m["PortName"]
                        });
                    }             
                }

                wmi = new ManagementClass("Win32_NetworkAdapter");
                foreach  (var m in wmi.GetInstances())
                {
                    if(m["Manufacturer"] != null)
                    {
                        if (!m["Manufacturer"].ToString().Contains("Microsoft") && 
                            !m["Manufacturer"].ToString().Contains("VMware"))
                        {
                            _lst.Add(new PC_Devices{
                                Type = "Network Adapter",
                                Name = m["ProductName"].ToString(),
                                DeviceID = (m["MACAddress"] != null ? m["MACAddress"].ToString() : "-NA-")
                            });
                        }
                    }
                }
            }
            catch { }

            return _lst;
        }

        public static List<PC_Software> Get_Softwares2()
        {
            List<PC_Software> _lst = new List<PC_Software>();

            //List<string> LST = new List<string>();
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
                            PC_Software _sf = new PC_Software
                            {
                                Name = sName,
                                Version = sVer,
                                Profile = "Local 64"
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
                            PC_Software _sf = new PC_Software
                            {
                                Name = sName,
                                Version = sVer,
                                Profile = "Current User"
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

        public static string Get_Board(ManagementScope MScope)
        {
            PCInfo _info = Get_BoardInfo(MScope);
            
            string sResult = Helper.Title("GENERAL INFORMATION");

            sResult += Helper.PadString("   System Model", 42) + " : " + _info.System_Model + Environment.NewLine;
            sResult += Helper.PadString("   Board Model", 42) + " : " + _info.Board_Model + Environment.NewLine;
            sResult += Helper.PadString("   Board Serial No", 42) + " : " + _info.Board_Serial + Environment.NewLine;
            sResult += Helper.PadString("   BIOS", 42) + " : " + _info.BIOS_Maker + " " + _info.BIOS_Serial + Environment.NewLine;
            sResult += Helper.PadString("   BIOS Date", 42) + " : " + _info.BIOS_Date + Environment.NewLine;
            sResult += Helper.PadString("   Processor", 42) + " : " + _info.Processor_Name + Environment.NewLine;
            sResult += Helper.PadString("   Processor Cores", 42) + " : " + _info.Processor_Cores + Environment.NewLine;
            sResult += Helper.PadString("   Logical Cores", 42) + " : " + _info.Processor_Logical + Environment.NewLine;
            sResult += Helper.PadString("   Memory", 42) + " : " + _info.System_RAM + Environment.NewLine;
            sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;
            sResult += Helper.PadString("   Hostname", 42) + " : " + _info.OS_Hostname + Environment.NewLine;
            sResult += Helper.PadString("   User Login", 42) + " : " + _info.OS_User + Environment.NewLine;
            sResult += Helper.PadString("   Domain", 42) + " : " + _info.OS_Domain + Environment.NewLine;
            sResult += Helper.PadString("   OS Installed", 42) + " : " + _info.OS_Name + ";build " + _info.OS_Build + Environment.NewLine;
            sResult += Helper.PadString("   OS Installed Date", 42) + " : " + _info.OS_InstallDate + Environment.NewLine;
            sResult += Environment.NewLine;

            return sResult;
        }

        public static string Get_RAM(ManagementScope MScope)
        {
            string sResult = "";

            ManagementObjectSearcher ObjSch;

            try
            {
                sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;
                sResult += Helper.PadString("   DIMM SLOT ", 45) +
                            Helper.PadString("Maker", 17) +
                            Helper.PadString("Size ", 15) +
                            Helper.PadString("Speed (MHz) ", 15) +
                            Helper.PadString("Serial No", 15) + Environment.NewLine;
                sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;

                ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM Win32_PhysicalMemory"));
                foreach (ManagementObject m in  ObjSch.Get())
                {
                    sResult += Helper.PadString("     " + m["DeviceLocator"].ToString() + " - " + m["BankLabel"].ToString(),45);
                    sResult += Helper.PadString(m["Manufacturer"].ToString(), 17);
                    sResult += Helper.PadString(Helper.FormatSize(Convert.ToInt64(m["Capacity"].ToString())), 15);
                    sResult += Helper.PadString(m["Speed"].ToString(), 15);
                    sResult += Helper.PadString(m["SerialNumber"].ToString(), 15) + Environment.NewLine ;        
                }

                sResult +=Environment.NewLine;
            }
            catch { }


            return sResult;
        }

        public static string Get_HID(ManagementScope MScope)
        {
            ManagementClass wmi;

            string sResult = "";

            try
            {
                sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;
                sResult += Helper.PadString("   Input Device ", 62) +
                           Helper.PadString("ID", 30) + Environment.NewLine;
                sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;

                wmi = new ManagementClass("Win32_KeyBoard");
                foreach (var keyboard in wmi.GetInstances())
                {
                    sResult += Helper.PadString("     " +  (string)keyboard["Description"],62);
                    sResult += Helper.PadString((string)keyboard["PNPDeviceId"],30);
                    sResult += Environment.NewLine;
                }

                wmi = new ManagementClass("Win32_PointingDevice");
                foreach (var mouse in wmi.GetInstances())
                {
                    sResult += Helper.PadString("     " + (string)mouse["Manufacturer"].ToString() + " " + (string)mouse["Caption"].ToString() ,62);
                    //sResult += Helper.PadString("   Description",30) + " : " + (string)mouse["Caption"].ToString() + Environment.NewLine;
                    sResult += Helper.PadString( (string)mouse["PNPDeviceId"].ToString(), 30);
                    sResult += Environment.NewLine;
                }

                sResult += Environment.NewLine;

                sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;
                sResult += Helper.PadString("   Printer Name ", 62) +
                           Helper.PadString("Port", 30) + Environment.NewLine;
                sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;

                wmi = new ManagementClass("Win32_Printer");
                foreach (var prn in wmi.GetInstances())
                {
                    sResult += Helper.PadString("     " + (string)prn["DriverName"].ToString()  ,62);
                    sResult += Helper.PadString( (string)prn["PortName"].ToString(), 30);
                    sResult += Environment.NewLine;
                }


                sResult += Environment.NewLine;
            }
            catch { }

            return sResult;
        }

        public static string Get_Accounts(ManagementScope MScope)
        {
            ManagementObjectSearcher ObjSch;

            string sResult = "";

            try
            {
                sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;
                sResult += Helper.PadString("   Local Users ", 62) + Environment.NewLine;
                sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;

                ObjSch = new ManagementObjectSearcher(MScope,  new ObjectQuery("SELECT * FROM Win32_UserAccount WHERE Domain = '" + Environment.MachineName + "' "));
                foreach (ManagementObject m in  ObjSch.Get())
                {
                    sResult += Helper.PadString("     " + m["Caption"].ToString(), 62)  + Environment.NewLine;
                }

                sResult +=Environment.NewLine;
            }
            catch { }

            return sResult;
        }

        public static string Get_Graphics(ManagementScope MScope)
        {
            string sResult = "";
            int cnt = 1;
            ManagementObjectSearcher ObjSch;

            try
            {
                sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;
                sResult +=  Helper.PadString("   DISPLAY GPU ", 45) +
                            Helper.PadString("Resolution ", 32) +
                            Helper.PadString("Version ", 20) +
                            Helper.PadString("Date ", 20) + Environment.NewLine;    
                sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;

                ObjSch = new ManagementObjectSearcher(MScope,  new ObjectQuery("SELECT * FROM Win32_VideoController"));
                foreach (ManagementObject m in  ObjSch.Get())
                {
                    sResult += Helper.PadString("     " + m["Caption"].ToString(), 45);

                    if (m["CurrentHorizontalResolution"] != null)
                    {
                        sResult += Helper.PadString(m["CurrentHorizontalResolution"].ToString() + " x " + 
                                                    m["CurrentVerticalResolution"].ToString()  + " @ " + 
                                                    m["CurrentRefreshRate"].ToString() + " Hz" ,32);

                        sResult += Helper.PadString(m["DriverVersion"].ToString(), 20);
                        sResult += Helper.PadString(ManagementDateTimeConverter.ToDateTime(m["DriverDate"].ToString()).ToString("ddd, MMM dd yyyy") , 20) + Environment.NewLine ;        
                    }
                }

                sResult += Environment.NewLine;
            }
            catch { }

            return sResult;
        }

        public static string Get_Printers(ManagementScope MScope)
        {
            string sResult = "";
            int cnt = 1;

            ManagementObjectSearcher ObjSch;

            sResult += Helper.Title("PRINTERS");

            try
            {
                ObjSch = new ManagementObjectSearcher(MScope,  new ObjectQuery("SELECT * FROM Win32_Printer"));
                foreach (ManagementObject m in  ObjSch.Get())
                {
                    if( !m["DriverName"].ToString().Contains("Microsoft"))
                    {
                        sResult += Helper.PadString("   Printer " + cnt,30) + " : " + m["DriverName"].ToString() + Environment.NewLine;
                        sResult += Helper.PadString("     Port",30) + " :   " + m["PortName"].ToString() + Environment.NewLine;
                    }

                    cnt++;
                }

                sResult += Environment.NewLine;
            }
            catch {  return sResult; }


            return sResult;
        }

        public static string Get_NetAdapters(ManagementScope MScope)
        {
            string sResult = "";

            ManagementObjectSearcher ObjSch;

            try
            {
                sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;
                sResult += Helper.PadString("   Network Adapter ", 62) +
                           Helper.PadString("MAC Address ", 30) + Environment.NewLine;
                sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;

                ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM Win32_NetworkAdapter "));
                foreach (ManagementObject m in ObjSch.Get())
                {
                    if(m["Manufacturer"] != null)
                    {
                        if (!m["Manufacturer"].ToString().Contains("Microsoft") && 
                            !m["Manufacturer"].ToString().Contains("VMware"))
                        {
                            sResult += Helper.PadString("     " + m["ProductName"].ToString(), 62);                            
                            sResult += Helper.PadString(m["MACAddress"] != null ? m["MACAddress"].ToString() : "-NA-" , 30) ;
                            sResult += Environment.NewLine ;
                        }
                    }
                }

                sResult += Environment.NewLine;
            }
            catch { }


            try
            {
                sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;
                sResult += Helper.PadString("   IP Address ", 45) +
                              Helper.PadString("Netmask ", 17) +
                              Helper.PadString("Gateway ", 18) +
                              Helper.PadString("MAC Address ", 18) + Environment.NewLine;
                sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;
              
                ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT Description,IPAddress,IPSubnet,DefaultIPGateway,MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'TRUE'"));
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
                                   Helper.PadString(sIPGateway, 18) +
                                   Helper.PadString(m["MACAddress"].ToString(), 18) + Environment.NewLine;
                    }
                }

                sResult += Environment.NewLine;
            }
            catch { }

            return sResult;
        }

        public static string Get_Drives(ManagementScope MScope)
        {
            string sResult = "", sLDrv = "", sNDrv = "", sName = "";
  
            ManagementObjectSearcher ObjSch;
            try
            {
                sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;
                sResult +=  Helper.PadString("   DISK ", 45) +
                            Helper.PadString("Size ", 17) +
                            Helper.PadString("Partitions ", 15) + Environment.NewLine;
                sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;

                ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM Win32_DiskDrive"));
                foreach (ManagementObject m in ObjSch.Get())
                {
                    sResult += Helper.PadString("     " + m["Caption"].ToString(), 45);
                    sResult += Helper.PadString(Helper.FormatSize(Convert.ToInt64(m["Size"].ToString())), 17);
                    sResult += Helper.PadString(m["Partitions"].ToString(), 15) + Environment.NewLine ;               
                }

                sResult +=Environment.NewLine;

                sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;
                sResult +=  Helper.PadString("   DRIVES ", 45) +
                            Helper.PadString("File System ", 17) +
                            Helper.PadString("Total Used ", 15) +
                            Helper.PadString("Total Capacity ", 15)  + Environment.NewLine;
                sResult += " -------------------------------------------------------------------------------------------------------------------" + Environment.NewLine;

                ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM Win32_LogicalDisk"));
                foreach (ManagementObject m in ObjSch.Get())
                {
                    int iDrvType = Convert.ToInt32(m["DriveType"].ToString());

                    if(iDrvType <= 3 || iDrvType == 5)
                    {
                        sName = m["VolumeName"] != null ? m["VolumeName"].ToString() : m["Description"].ToString();

                        if (sName == "") sName = m["Description"].ToString();

                        sLDrv += Helper.PadString("     " + sName + " (" + m["Caption"].ToString() + ") ", 45);

                        if (m["Size"] != null )
                        {
                            sLDrv += Helper.PadString(m["FileSystem"].ToString(), 17);
                            sLDrv += Helper.PadString(Helper.FormatSize(Convert.ToInt64(m["Size"].ToString()) - Convert.ToInt64(m["FreeSpace"].ToString())), 15);
                            sLDrv += Helper.PadString(Helper.FormatSize(Convert.ToInt64(m["Size"].ToString())), 15);
                        }

                        sLDrv += Environment.NewLine;
                    }

                    if (iDrvType == 4 || iDrvType > 5)
                    {
                        sNDrv += Helper.PadString("     " + m["ProviderName"].ToString() + " (" + m["Caption"].ToString() + ") ", 45);

                        sNDrv += Environment.NewLine;
                    }
                }

                sResult += sLDrv + sNDrv;
                sResult +=Environment.NewLine;
            }
            catch { return sResult; }

            return sResult;
        }

        public static string Get_Softwares()
        {
            string sResult = "";

            int iNameLen = 0;
            int iVerLen = 0;

            List<PC_Software> _lst = Get_Softwares2().OrderBy(x => x.Name).ToList();

            iNameLen = _lst.Max(x => x.Name.Length);
            iVerLen = _lst.Max(x=>x.Version.Length);

            sResult += Helper.Title(Helper.PadString("INSTALLED SOFTWARE", iNameLen + 4) + " " + Helper.PadString("VERSION", iVerLen + 3) );

            foreach (PC_Software sf in _lst )
            {
                sResult += Helper.PadString("     " + sf.Name , iNameLen + 6) + " " + 
                           Helper.PadString( sf.Version , iVerLen + 3) + Environment.NewLine;
            }
            
            return sResult;
        }


        private static long Get_MemoryTotal()
        {
            long memKb;
            GetPhysicallyInstalledSystemMemory(out memKb);
            return (memKb / 1024); // MB
        }
    }


  
}
