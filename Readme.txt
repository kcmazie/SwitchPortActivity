SwitchPortActivity
A PowerShell Script to track Cisco switch port activity over time on MS Excel

     File Name : SwitchPortActivity.ps1	 
        Author : Kenneth C. Mazie [kcmjr AT kcmjr.com]
               :
   Description : Tracks switch port status over time using MS Excel.  This is primarily 
               : intended for cases where your techs don't pull down x-connect jumpers in
               : IT closets after removing equipment.  The resulting spreadsheet gives you 
               : an idea of which ports have been abandoned.  The script was designed to 
               : access and parse Cisco switches.
               :
         Notes : Normal operation is with no command line options.  If pre-stored credentials 
               : are desired use this: https://www.powershellgallery.com/packages/CredentialsWithKey/1.10
               : Spreadsheet can be generated via a flat list of IP addresses or pulled from a master
               : copy.  The first column contains the list of IP addresses OR the "Switches" tab from a 
               : master inventory.  Each IP gets a dedicated worksheet labeled with the IP.  Column A 
               : is the port ID.  Cell "A1" is a ROUGH total port count.  The top row is the date.
               : The spreadsheet is color coded for readability:
               : - If a port at any time registers as connected the "A" colum gets high-lited.
               : - Connected ports are red.
               : - Not-connected ports are green (OK to unplug).
               : - A connection change is flagged in bold violet
               : - Disabled ports are tagged blue
               :
  Requirements : Plink.exe must be available in your path or the full path must be included in the 
               : commandline(s) below.  2 versions are used in case of version issues.  These are located
               : in the same folder and named according to version (see line 136 below). Excel must be 
               : available on the local PC.  SSH Keys must already be stored on the local PC through the
               : use of PuTTY or conenction will fail.  An option exists to add it below.
               :
      Switches : $Console - If Set to $true will display status during run (Defaults to $True)
               : $Debug - If set to $true adds extra output on screen (Defaults to $false)
               : $SafeUpdate - If set to $True backs up spreadsheet prior to updating.  Keeps 10 copies.
               :
      Warnings : Excel is set to be visible (can be changed) so don't mess with it while the script is 
               : running or it can crash.  Don't click in spreadsheet while running or the script will crash.
               :
         Legal : Public Domain. Modify and redistribute freely. No rights reserved.
               : SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF 
               : ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED.
               : That being said, feel free to ask if you have questions...
               :
       Credits : Code snippets and/or ideas came from many sources including but 
               : not limited to the following:
               :
Last Update by : Kenneth C. Mazie.
Change History : v1.00 - 04-16-23 - Original Draft
               : v1.10 - 05-18-23 - Adjusted coding for disabled ports when the description contains "bad".
