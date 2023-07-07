# AutoCOMfig
Simple powershell script that automatically sets a given device a COM port and a configuration

The idea of this project is to automatize device configuration in large scale projects, where you have to manage lots of computer using different devices that need setup (AKA mostly factories). It just configures the devices you plug when you run it but it is possible to call it with a task when a USB device is plugged, for example

This is basically an upgraded version of the following article, adding some safety checks, feedback messages and xml reading, to list the devices you want to configure and how to do it

https://syswow.blogspot.com/2013/03/change-device-com-port-via-powershell.html

## How to set up
1. Have both *.xml* and *.ps1* files at the same folder
2. Fill the *.xml* with your device info  
- Name: The name you want the device to be called after configuration
  
- HardwareID: Device HardwareID. You can look it up in; Device Manager --> (right click your device) Properties --> Details --> HardwareID
  
- Configuration: This determines Bits/s, Data bits, Parity, Stopping Bits and Flux control settings the device will be given. These values are located at "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Ports\COMX"
  
- Newport: COM port the device will use once configured

Example:

&lt;Device name="Kaba desktop reader 91 08" hardwareid="FTDIBUS\\\\COMPORT&amp;amp;VID_18D9&amp;amp;PID_0211" configuration="115200,n,8,1 " newport="COM9"/>

**NOTE THAT "CONFIGURATION" DATA HAS A SPACE AT THE END*** ALSO THAT WE USE "&amp;amp;" INSTEAD OF "&amp;" AND "\\\\" INSTEAD OF "\\"*

3. Execute the .ps1 as admin, it won't edit the necessary registry keys otherwise
