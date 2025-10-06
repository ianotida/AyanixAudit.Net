using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPCAudit.Models
{
    public class PC_Info
    {
        public string Board_Maker { get; set; }
        public string Board_Model { get; set; }
        public string Board_Serial { get; set; }
        public string BIOS_Maker { get; set; }
        public string BIOS_Serial { get; set; }
        public string BIOS_Date { get; set; }
        public string Processor_Name { get; set; }
        public string Processor_Cores { get; set; }
        public string Processor_Logical { get; set; }
        public string System_Model { get; set; }
        public string System_RAM { get; set; }
        public string OS_Name { get; set; }
        public string OS_Build { get; set; }
        public string OS_Hostname { get; set; }
        public string OS_InstallDate { get; set; }
        public string OS_Domain{ get; set; }
        public string OS_User { get; set; }
        public string OS_DiskType { get; set; }
        public string OS_DiskSize { get; set; }
        public int OS_DiskIndex { get; set; }
    }

    public class PC_Devices
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string DevID { get; set;}
    }


    public class PC_Drive
    {
        public int Index { get; set; }
        public int Boot { get; set; }
        public string Letter { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public string FileSystem { get; set; }
        public string Partition { get; set; }
        public ulong Size_U64 { get; set; }
        public string Size { get; set; }
        public string Free { get; set; }
        public string Used { get; set; }
        public string DevType { get; set; }
    }

    public class PC_Software
    {
        public string Name { get; set; }
        public string Version { get; set; }
        public string Profile { get; set; }
    }

    public class PC_Net
    {
        public string Name { get; set; }
        public string IP_Address { get; set; }
        public string IP_Mask { get; set;}
        public string Gateway { get; set; }
        public string MAC_Address { get; set; }
    }
}
