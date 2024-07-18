using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Management;
using System.Runtime.InteropServices;
using System.Data;
using System.Linq;
using System.Net;
using AyanixAudit.Models;

namespace AyanixAudit.Globals
{
    public static class Get_WMI
    {
        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetPhysicallyInstalledSystemMemory(out long TotalMemoryInKilobytes);

        public static PCInfo Get_Board2(ManagementScope MScope)
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

                            if (!_lst.Contains(_sf)) _lst.Add(_sf);
                        }
                    }
                }
            }
            catch { }

            try
            { 
                // Search in: LocalMachine_32
                key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
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
                                Profile = "Local 32"
                            };

                            if (!_lst.Contains(_sf)) _lst.Add(_sf);
                        }
                    }
                }
            }
            catch { }

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

                            if (!_lst.Contains(_sf)) _lst.Add(_sf);
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
            string sResult = "";
            
            ManagementObjectSearcher ObjSch;

            sResult += Helper.Title("GENERAL INFORMATION");

            try
            {
                ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM Win32_BaseBoard"));
                foreach (ManagementObject m in ObjSch.Get())
                {
                    sResult += Helper.PadString("   Board Model", 30) + " : " + m["Manufacturer"].ToString() + " " + m["Product"].ToString() + Environment.NewLine;
                    sResult += Helper.PadString("   Board Serial No", 30) + " : " + m["SerialNumber"].ToString() + Environment.NewLine;
                }

                ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM Win32_BIOS"));
                foreach (ManagementObject m in ObjSch.Get())
                {
                    sResult += Helper.PadString("   BIOS", 30) + " : " + m["Manufacturer"].ToString() + " " + m["SerialNumber"].ToString()  + Environment.NewLine;
                    sResult += Helper.PadString("   BIOS Release Date", 30) + " : " + ManagementDateTimeConverter.ToDateTime(m["ReleaseDate"].ToString()).ToString("ddd, MMM dd yyyy hh:mm:ss tt") + Environment.NewLine;
                }

                ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM Win32_Processor"));
                foreach (ManagementObject m in ObjSch.Get())
                {
                    sResult += Helper.PadString("   Processor", 30) + " : " + m["Name"].ToString() + Environment.NewLine;
                    sResult += Helper.PadString("   Processor Cores", 30) + " : " + m["NumberOfCores"].ToString() + Environment.NewLine;
                    sResult += Helper.PadString("   Logical Cores", 30) + " : " + m["NumberOfLogicalProcessors"].ToString() + Environment.NewLine;
                }

                sResult += " --------------------------------------------------------------------------------------------------------------" + Environment.NewLine;

                ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM Win32_ComputerSystem"));
                foreach (ManagementObject m in ObjSch.Get())
                {
                    sResult += Helper.PadString("   System Name", 30) + " : " + Environment.MachineName + Environment.NewLine;
                    sResult += Helper.PadString("   System Model", 30) + " : " + m["Model"].ToString() + Environment.NewLine;
                    sResult += Helper.PadString("   System Memory", 30) + " : " + (Get_MemoryTotal() /1024) + " GB"  + Environment.NewLine;
                    sResult += Helper.PadString("   Logon User", 30) + " : " + m["UserName"].ToString() + Environment.NewLine;
                    sResult += Helper.PadString("   Domain", 30) + " : " + m["Domain"].ToString() + Environment.NewLine;
                }

                ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM Win32_OperatingSystem"));
                foreach (ManagementObject m in ObjSch.Get())
                {
                    sResult += Helper.PadString("   OS Name", 30) + " : " + m["Caption"].ToString() + " (" +  m["OSArchitecture"].ToString() + ")" + Environment.NewLine;
                    sResult += Helper.PadString("   OS Build", 30) + " : " + m["Version"].ToString() + Environment.NewLine;
                    sResult += Helper.PadString("   Date Installed", 30) + " : " + ManagementDateTimeConverter.ToDateTime(m["InstallDate"].ToString()).ToString("ddd, MMM dd yyyy hh:mm:ss tt") + Environment.NewLine;                    
                }

                //sResult += Environment.NewLine;
            }
            catch { }

            return sResult;
        }

        public static string Get_HID(ManagementScope MScope)
        {
            string sResult = "";

            ManagementClass wmi;

            sResult += Helper.Title("INPUT DEVICES");

            try
            {
                wmi = new ManagementClass("Win32_KeyBoard");
                foreach (var keyboard in wmi.GetInstances())
                {
                    sResult += Helper.PadString("   Description",30) + " : " + (string)keyboard["Description"] + Environment.NewLine;
                    sResult += Helper.PadString("     Device ID",30) + " :   " + (string)keyboard["PNPDeviceId"] + Environment.NewLine;
                    //sResult += Environment.NewLine;
                }

                wmi = new ManagementClass("Win32_PointingDevice");
                foreach (var mouse in wmi.GetInstances())
                {
                    sResult += Helper.PadString("   Description",30) + " : " + (string)mouse["Manufacturer"].ToString() + " " + (string)mouse["Caption"].ToString() + Environment.NewLine;
                    //sResult += Helper.PadString("   Description",30) + " : " + (string)mouse["Caption"].ToString() + Environment.NewLine;
                    sResult += Helper.PadString("     Device ID",30) + " :   " + (string)mouse["PNPDeviceId"].ToString() + Environment.NewLine;

                    //sResult += Environment.NewLine;
                }

                //sResult += Environment.NewLine;
            }
            catch { }

            return sResult;
        }

        public static string Get_Accounts(ManagementScope MScope)
        {
            string sResult = "";
            int cnt = 1;

            ManagementObjectSearcher ObjSch;

            sResult += Helper.Title("LOCAL ACCOUNTS");

            try
            {
                ObjSch = new ManagementObjectSearcher(MScope,  new ObjectQuery("SELECT * FROM Win32_UserAccount WHERE Domain = '" + Environment.MachineName + "' "));
                foreach (ManagementObject m in  ObjSch.Get())
                {
                    sResult += Helper.PadString("   User " + cnt, 30) + " : " + m["Caption"].ToString() + Environment.NewLine;

                    cnt++;
                }

                //sResult += Environment.NewLine;
            }
            catch { }

            return sResult;
        }

        public static string Get_Graphics(ManagementScope MScope)
        {
            string sResult = "";
            int cnt = 1;

            ManagementObjectSearcher ObjSch;

            sResult += Helper.Title("DISPLAY ADAPTER");

            try
            {
                ObjSch = new ManagementObjectSearcher(MScope,  new ObjectQuery("SELECT * FROM Win32_VideoController"));
                foreach (ManagementObject m in  ObjSch.Get())
                {
                    sResult += Helper.PadString("   Display " + cnt,30) + " : " + m["Caption"].ToString() + Environment.NewLine;

                    if (m["CurrentHorizontalResolution"] != null)
                    {
                        sResult += Helper.PadString("     Resolution",30) + " : " + m["CurrentHorizontalResolution"].ToString() + " x " + 
                                                                                    m["CurrentVerticalResolution"].ToString()  + " @ " + 
                                                                                    m["CurrentRefreshRate"].ToString() + " Hz" + Environment.NewLine;

                        sResult += Helper.PadString("     Driver Version",30) + " : " + m["DriverVersion"].ToString() + Environment.NewLine;
                        sResult += Helper.PadString("     Driver Date",30) + " : " + ManagementDateTimeConverter.ToDateTime(m["DriverDate"].ToString()).ToString("ddd, MMM dd yyyy hh:mm:ss tt") + Environment.NewLine;
                    }
                   
                    cnt++;
                    //sResult += Environment.NewLine;
                }

                //sResult += Environment.NewLine;
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

                //sResult += Environment.NewLine;
            }
            catch {  return sResult; }


            return sResult;
        }

        public static string Get_NetAdapters(ManagementScope MScope)
        {
            string sResult = "";
            int cnt = 1;

            ManagementObjectSearcher ObjSch;

            sResult += Helper.Title("NETWORK ADAPTERS");

            try
            {
                ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM Win32_NetworkAdapter "));
                foreach (ManagementObject m in ObjSch.Get())
                {
                    if(m["Manufacturer"] != null)
                    {
                        if (!m["Manufacturer"].ToString().Contains("Microsoft") && 
                            !m["Manufacturer"].ToString().Contains("VMware"))
                        {
                            sResult += Helper.PadString("   Network Adapter " + cnt, 30) + " : " + m["ProductName"].ToString() + Environment.NewLine;                            
                            sResult += Helper.PadString("     MAC Address", 30) + " :   " + (m["MACAddress"] != null ? m["MACAddress"].ToString() : "-NA-");

                            sResult += Environment.NewLine ;
                            cnt++;
                        }
                    }
                }

                sResult += " --------------------------------------------------------------------------------------------------------------" + Environment.NewLine;
                sResult += Helper.PadString("   IP Address ", 33) +
                                Helper.PadString("Netmask ", 18) +
                                Helper.PadString("Gateway ", 18) +
                                Helper.PadString("Network Adapter ", 18) + Environment.NewLine;
                sResult += " --------------------------------------------------------------------------------------------------------------" + Environment.NewLine;
              
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
                        sResult += Helper.PadString("   " + arr_IP[i], 33) +
                                   Helper.PadString(arr_Sub[i], 18) +
                                   Helper.PadString(sIPGateway, 18) +
                                   Helper.PadString(m["Description"].ToString(), 18) + Environment.NewLine;
                    }

                    //sIPAddress = string.Join(", ", (string[])m["IPAddress"]);
                    //sIPMask = string.Join(", ", (string[])m["IPSubnet"]);



                    //if( sIPAddress != "" )
                    //{
                    //    //foreach(string s in (string[])m["IPSubnet"]){           sIPMask = sIPMask == "" ? s : "," + s;          }
                    //    //foreach(string g in (string[])m["DefaultIPGateway"]){   sIPGateway = sIPGateway == "" ? g : "," + g;    }

                    //    sResult += Helper.PadString("   " + sIPAddress, 25) +
                    //               Helper.PadString(sIPMask, 25) +
                    //               Helper.PadString(sIPGateway, 25) +
                    //               Helper.PadString(m["Description"].ToString(), 25) + Environment.NewLine;
                    //}
                }


                //ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM Win32_IP4RouteTable"));

                //foreach (ManagementObject m in ObjSch.Get())
                //{
                //    sResult += Helper.PadString("   " + m["Destination"].ToString(), 25) +
                //                Helper.PadString(m["Mask"].ToString(), 25) +
                //                Helper.PadString(m["NextHop"].ToString(), 25) +
                //                Helper.PadString(m["Metric1"].ToString(), 25) + Environment.NewLine;
                //}

                //sResult += Environment.NewLine;
            }
            catch { }

            return sResult;
        }

        public static string Get_Drives(ManagementScope MScope)
        {
            string sResult = "", sLDrv = "", sNDrv = "", sName = "";
            int cnt = 1;

            ManagementObjectSearcher ObjSch;

            sResult += Helper.Title("DISK AND DRIVES");

            try
            {
                ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM Win32_DiskDrive"));
                foreach (ManagementObject m in ObjSch.Get())
                {
                    sResult += Helper.PadString("   Physical Disk " + cnt, 30) + " : " + m["Caption"].ToString() + Environment.NewLine;
                    sResult += Helper.PadString("     Size", 30) + " : " + Helper.FormatSize(Convert.ToInt64(m["Size"].ToString())) + Environment.NewLine;
                    sResult += Helper.PadString("     Partitions", 30) + " : " + m["Partitions"].ToString() + Environment.NewLine ;

                    cnt++;
                }

                sResult += " --------------------------------------------------------------------------------------------------------------" + Environment.NewLine;
                sResult +=  Helper.PadString("   Drives ", 33) +
                            Helper.PadString("File System ", 17) +
                            Helper.PadString("Total Used ", 15) +
                            Helper.PadString("Total Capacity ", 15)  + Environment.NewLine;
                sResult += " --------------------------------------------------------------------------------------------------------------" + Environment.NewLine;

                ObjSch = new ManagementObjectSearcher(MScope, new ObjectQuery("SELECT * FROM Win32_LogicalDisk"));
                foreach (ManagementObject m in ObjSch.Get())
                {
                    int iDrvType = Convert.ToInt32(m["DriveType"].ToString());

                    if(iDrvType <= 3 || iDrvType == 5)
                    {
                        sName = m["VolumeName"] != null ? m["VolumeName"].ToString() : m["Description"].ToString();

                        if (sName == "") sName = m["Description"].ToString();

                        sLDrv += Helper.PadString("   " + sName + " (" + m["Caption"].ToString() + ") ", 33);

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
                        sNDrv += Helper.PadString("   " + m["ProviderName"].ToString() + " (" + m["Caption"].ToString() + ") ", 55);
     
                        sNDrv += Environment.NewLine;
                    }
                }

                sResult += sLDrv + sNDrv;
                //sResult +=Environment.NewLine;
            }
            catch { return sResult; }

            return sResult;
        }

        public static string Get_Softwares()
        {
            string sResult = "";

            DataTable DTProg = new DataTable();
            DTProg.Columns.Add("Name", typeof(string));
            DTProg.Columns.Add("Version", typeof(string));
            DTProg.Columns.Add("Installed", typeof(string));

            //List<string> LST = new List<string>();
            RegistryKey key;

            string sName = "", sVer = "";

            try
            {
                // Search in: CurrentUser
                key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
                foreach (String keyName in key.GetSubKeyNames())
                {
                    RegistryKey Rkey = key.OpenSubKey(keyName);

                    sName = Rkey.GetValue("DisplayName") != null ? Rkey.GetValue("DisplayName").ToString().Trim() : "";
                    sVer = Rkey.GetValue("DisplayVersion")!= null ? Rkey.GetValue("DisplayVersion").ToString().Trim() : "";

                    if (sName != "" && DTProg.Select("Name = '" + sName + "' AND Version = '" + sVer + "'").Length == 0)
                    {
                        if(!sName.Contains("Security Update") && 
                           !sName.Contains("Update for Microsoft"))
                        {
                            DTProg.Rows.Add(sName, sVer,"Current User");
                        }
                        
                    }

                    //if (!LST.Contains(Rkey.GetValue("DisplayName") + " " + Rkey.GetValue("DisplayVersion")))
                    //{
                    //    //LST.Add(Rkey.GetValue("DisplayName") + " " + Rkey.GetValue("DisplayVersion"));
                    //}
                }
            }
            catch { }

            try
            { 
                // Search in: LocalMachine_32
                key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");
                foreach (String keyName in key.GetSubKeyNames())
                {
                    RegistryKey Rkey = key.OpenSubKey(keyName);

                    sName = Rkey.GetValue("DisplayName") != null ? Rkey.GetValue("DisplayName").ToString().Trim() : "";
                    sVer = Rkey.GetValue("DisplayVersion")!= null ? Rkey.GetValue("DisplayVersion").ToString().Trim() : "";

                    if (sName != "" && DTProg.Select("Name = '" + sName + "' AND Version = '" + sVer + "'").Length == 0)
                    {

                        if(!sName.Contains("Security Update") && 
                           !sName.Contains("Update for Microsoft"))
                        {
                            DTProg.Rows.Add(sName, sVer,"Local 32");
                        }
                        
                    }

                    //if (!LST.Contains(Rkey.GetValue("DisplayName") + " " + Rkey.GetValue("DisplayVersion")))
                    //{
                    //    //LST.Add(Rkey.GetValue("DisplayName") + " " + Rkey.GetValue("DisplayVersion"));
                    //    DTProg.Rows.Add(Rkey.GetValue("DisplayName"), Rkey.GetValue("DisplayVersion"));
                    //}
                }
            }
            catch { }

            try
            {
                // Search in: LocalMachine_64
                key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall");
                foreach (String keyName in key.GetSubKeyNames())
                {
                    RegistryKey Rkey = key.OpenSubKey(keyName);

                    sName = Rkey.GetValue("DisplayName") != null ? Rkey.GetValue("DisplayName").ToString().Trim() : "";
                    sVer = Rkey.GetValue("DisplayVersion")!= null ? Rkey.GetValue("DisplayVersion").ToString().Trim() : "";

                    if (sName != "" && DTProg.Select("Name = '" + sName + "' AND Version = '" + sVer + "'").Length == 0)
                    {
                        if(!sName.Contains("Security Update") && 
                           !sName.Contains("Update for Microsoft"))
                        {
                            DTProg.Rows.Add(sName, sVer,"Local 64");
                        }
                        
                    }

                    //if (!LST.Contains(Rkey.GetValue("DisplayName") + " " + Rkey.GetValue("DisplayVersion")))
                    //{
                    //    //LST.Add(Rkey.GetValue("DisplayName") + " " + Rkey.GetValue("DisplayVersion"));
                    //    DTProg.Rows.Add(Rkey.GetValue("DisplayName"), Rkey.GetValue("DisplayVersion"));
                    //}
                }
            }
            catch { }

            DTProg.DefaultView.Sort = "Name ASC, Version ASC";


            int iNameLen = 0;
            int iVerLen = 0;

            foreach (DataRow Dr in DTProg.DefaultView.ToTable().Rows)
            {
                if (Dr["Name"].ToString().Length > iNameLen) iNameLen = Dr["Name"].ToString().Length;
                if (Dr["Version"].ToString().Length > iVerLen) iVerLen = Dr["Version"].ToString().Length;
            }


            sResult += Helper.Title(Helper.PadString("INSTALLED SOFTWARE", iNameLen) + " : " + Helper.PadString("VERSION", iVerLen + 2) + "INSTALLED");

            foreach (DataRow Dr in DTProg.DefaultView.ToTable().Rows)
            {
                sResult += Helper.PadString("  " + Dr["Name"].ToString() , iNameLen + 2) + " : " + 
                           Helper.PadString(Dr["Version"].ToString() , iVerLen + 2) + 
                           Dr["Installed"].ToString() + Environment.NewLine;
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
