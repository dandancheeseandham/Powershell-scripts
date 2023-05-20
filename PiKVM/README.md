# Create .img files for PiKVM in Windows.

**These scripts and this procedure are not for running on any machine you care about - YOU HAVE BEEN WARNED**

This is to create a .img file for use with the PiKVM in Windows.
HOWEVER: The following procedure however involves Ventoy2Disk in /NOUSBCheck mode - which is  **dangerous** 

So only run this on a test rig, or a blank Hyper-V environment. 
**Do not blame me if you wipe the wrong drive!**

You can do the procedure manually by

Downloading
 * Ventoy - https://www.ventoy.net/en/download.html
 * qemu-img for Windows - https://cloudbase.it/qemu-img-windows/

Then:-

1) Create and Mount a VHD big enough for your ISO's (we'll call it alltheisos.vhd as an example)
2) Use Ventoy2Disk  Gui (Option >  Show all devices) or using the CLI switch "/NOUSBCheck"
3) Run Ventoy on the VHD drive
4) Once you have a drive letter, copy the ISO's over as a normal Ventoy install.
5) Once copied, detach and run 
6) Run the command qemu-img convert -O raw alltheisos.vhd alltheisos.img

----

I have created a Powershell scripts here that automates this. These are for my own use. :)
 Just run the script 1-GetPrograms.ps1 once, this will download Ventoy and Qemu-img, extract and create an ISO folder. This only needs to be run once on any single machine
 From then on, 
  * Place the ISO's you want in an image in the ISO folder
  * run 2-CreateIMG.ps1 and an image file called "Ventoy.img" will be created for you to upload to your PiKVM
 

**These scripts and this procedure are not for running on any machine you care about - YOU HAVE BEEN WARNED**