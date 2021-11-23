#***************************************************************************************************************
# Author: Damien VAN ROBAEYS
# Website: http://www.systanddeploy.com
# Twitter: https://twitter.com/syst_and_deploy
#***************************************************************************************************************
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | out-null

$Current_Folder = split-path $MyInvocation.MyCommand.Path
$Sources = $Current_Folder + "\" + "Sources\*"
If(test-path $Sources)
	{	
		$ProgData = $env:ProgramData
		$Destination_folder = "$ProgData\Windows_Sandbox_Editor"
		If(!(test-path $Destination_folder)){new-item $Destination_folder -type directory -force | out-null}
		copy-item $Sources $Destination_folder -force -recurse
		Get-Childitem -Recurse $Destination_folder | Unblock-file	
		
		$Get_Current_user = (gwmi win32_computersystem).username
		$Get_Current_user_Name = $Get_Current_user.Split("\")[1]		
		$User_Desktop_Profile = "C:\Users\$Get_Current_user_Name\Desktop"

		$Get_Tool_Shortcut = "$Destination_folder\Windows Sandbox Editor.lnk"		
		$Get_Desktop_Profile = [environment]::GetFolderPath('Desktop')
		copy-item $Get_Tool_Shortcut $Get_Desktop_Profile -Force

		[System.Windows.Forms.MessageBox]::Show("Windows Sandbox Editor has been installed.`nA shortcut has been added on your Desktop.")			
	}
Else
	{
		[System.Windows.Forms.MessageBox]::Show("It seems you don't have dowloaded all the folder structure.`nThe folder Sources is missing !!!")	
	}
