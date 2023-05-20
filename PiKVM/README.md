These scripts and this procedure are not for running on any machine you care about - YOU HAVE BEEN WARNED

It appears fairly easy to create a .img file for use with the PiKVM in Windows.
HOWEVER: The following procedure however involves Ventoy2Disk in /NoUSBCheck mode - which is  **dangerous** 

Only run this on a test rig, or a blank Hyper-V environment. 
Do not blame me if you wipe the wrong drive!

To do it manually get

 * Ventoy - https://github.com/ventoy/Ventoy/releases/download/v1.0.91/ventoy-1.0.91-windows.zip
 * Qemu-img - https://cloudbase.it/downloads/qemu-img-win-x64-2_3_0.zip

1) Create and Mount a VHD big enough for your ISO's (we'll call it alltheisos.vhd as an example)
2) Use Ventoy2Disk  Gui (Option >  Show all devices) or using the CLI switch "/NOUSBCheck"
3) Run Ventoy on the VHD drive
4) Once you have a drive letter, copy the ISO's over as a normal Ventoy install.
5) Once copied, detach and run 
6) run th command qemu-img  convert -O raw alltheisos.vhd

I have created a Powershell script that automates this for my own use.

These scripts are not for running on any machine you care about - YOU HAVE BEEN WARNED
