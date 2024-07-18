using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AyanixAudit.Models
{
    public class PCInfo
    {
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

        public List<PC_Disk> List_Disk { get; set; }
        public List<PC_Drive> List_Drives { get; set; }
        public List<PC_Devices> List_Devices { get; set; }
        public List<PC_Software> List_Softwares { get; set; }
        public List<PC_Network> List_Network { get; set; }
    }

    public class PC_Devices
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string DeviceID { get; set;}
    }

    public class PC_Disk
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string Size { get; set; }
        public string Partition { get; set; }
    }

    public class PC_Drive
    {
        public string Name { get; set; }
        public string File_System { get; set; }
        public string Total_Free { get; set; }
        public string Total_Used { get; set; }
        public string Total_Capacity { get; set; }
    }

    public class PC_Software
    {
        public string Name { get; set; }
        public string Version { get; set; }
        public string Profile { get; set; }
    }

    public class PC_Network
    {
        public string MAC_Address { get; set; }
        public string IP_Address { get; set; }
        public string Subnet { get; set;}
        public string Gateway { get; set; }
        public string Adapter { get; set; }
    }
}
