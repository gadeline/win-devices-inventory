from datetime import date
from datetime import time
from datetime import datetime
import math
import wmi
import os





# This function allows to convert the Byte size value to KB/MB/GB/TB used in this script 
def convert_size(size_bytes):
   if size_bytes == 0:
       return "0B"
   size_name = ("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
   i = int(math.floor(math.log(size_bytes, 1024)))
   p = math.pow(1024, i)
   s = round(size_bytes / p, 2)
   return "%s %s" % (s, size_name[i])


# Delete output text file if existing
if os.path.exists("c:/temp/export.txt"):
   os.remove("c:/temp/export.txt")

# Creation and opening of output file on mode append
file = open("c:/temp/export.txt","a+")


#file.write("Hostname" + "," + "OS" + "," + "Installation Date" + "," + "Last reboot" + "," + "Architecture" + "," + "Domaine/Workgroup" + "," + "Manufacturer" + "," + "Model" + "," + "TotalPhysicalMemory" + "," + "Serial" + "," + "Processor" + "," + "MaxClockspeed (Ghz)" + "," + "NumberOfCores" + "," + "NumberOfLogicalProcessors" + "," + "DiskName" + "," + "volsize" + "," + "volfreespace" + "\n")
file.write("Date/Time;Hostname;OS;Installation Date;Last reboot;Architecture;Domaine/Workgroup;Manufacturer;Model;Serial;TotalPhysicalMemory;ProcessorType;MaxClockspeed (Mhz);NumberOfCores;NumberOfLogicalProcessors;DiskName;TotalVolSize;TotalVolfreespace;InstalledSoftwares" + "\n")


# Insert Inventory date in output file
now = datetime.now()
file.write(now.strftime("%x") + " " + now.strftime("%X") + ";")


# Creation of WMI object
wmiObject = wmi.WMI()

# Retrieve some system informations on servers via WMI (eg.: Computername, OS version, Install date...)
for OS in wmiObject.win32_OperatingSystem():

   file.write(OS.CSName + ";" + OS.Caption + ";" + OS.InstallDate + ";" + OS.LastBootUpTime + ";" + OS.OSArchitecture + ";")
   #file.close()

# Retrieve some system informations on servers via WMI (eg.: Domain/Workgroup, Manufacturer, Model...)
for CS in wmiObject.win32_ComputerSystem():

   file.write(CS.Domain + ";" + CS.Manufacturer + ";" + CS.Model + ";")

# Retrieve some system informations on servers via WMI (serial number of server...)
for Bios in wmiObject.Win32_Bios():

   file.write(Bios.SerialNumber + ";"  + convert_size(int(CS.TotalPhysicalMemory)) + ";")

# Retrieve some system informations on servers via WMI (processor information of server)
for proc in wmiObject.Win32_Processor():

   file.write(proc.Caption + ";" + (str(proc.MaxClockSpeed)) + ";" + (str(proc.NumberOfCores)) + ";" + (str(proc.NumberOfLogicalProcessors)) + ";")



# Retrieve some system informations on servers via WMI (Storage information of server)
for disk in wmiObject.Win32_LogicalDisk(DriveType=3):
	
   disksize = disk.size
   diskdeviceid = disk.DeviceID
   diskfreeSpace = disk.freeSpace

   # convert size value to appropriete (KB or GB or TB...)
   volsize = convert_size(int(disksize))
   volfreespace = convert_size(int(diskfreeSpace))
   

   file.write(diskdeviceid + ";" + volsize + ";" + volfreespace + ";")

# Retrieve some system informations on servers via WMI (installed softwares on server)
for soft in wmiObject.Win32_Product():

   text = soft.caption + ", " + soft.vendor + ", " + soft.version
   file.write(text + " --- ")

file.close()