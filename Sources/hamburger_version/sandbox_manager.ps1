#========================================================================
#
# Author 	: Damien VAN ROBAEYS
# Website	: http://www.systanddeploy.com/
# Twitter	: https://twitter.com/syst_and_deploy
#
#========================================================================

[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')  				| out-null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | out-null
[System.Reflection.Assembly]::LoadFrom('assembly\MahApps.Metro.dll')      | out-null  
[System.Reflection.Assembly]::LoadFrom('assembly\System.Windows.Interactivity.dll') | out-null
[System.Reflection.Assembly]::LoadFrom('assembly\MahApps.Metro.IconPacks.dll')      | out-null  
[System.Windows.Forms.Application]::EnableVisualStyles()

#########################################################################
#                        Load Main Panel                                #
#########################################################################

$Global:pathPanel= split-path -parent $MyInvocation.MyCommand.Definition

function LoadXaml ($filename){
    $XamlLoader=(New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($filename)
    return $XamlLoader
}


$XamlMainWindow=LoadXaml($pathPanel+"\sandbox_manager.xaml")
$reader = (New-Object System.Xml.XmlNodeReader $XamlMainWindow)
$Form = [Windows.Markup.XamlReader]::Load($reader)


#########################################################################
#                        HAMBURGER VIEWS                                #
#########################################################################

#******************* Item in the menu  *****************
$Basic_infos      		= $Form.FindName("Basic_infos") 
$Mapped_folders      	= $Form.FindName("Mapped_folders") 
$Startup_commands   	= $Form.FindName("Startup_commands")
$Overview   			= $Form.FindName("Overview") 
$MonBouton   			= $Form.FindName("MonBouton") 



#******************* Generral controls  *****************
$TabMenuHamburger 	= $Form.FindName("TabMenuHamburger")

#******************* Load Other Views  *****************
$viewFolder = $pathPanel +"\views"

$XamlChildWindow = LoadXaml($viewFolder+"\Basic_infos.xaml")
$Childreader     = (New-Object System.Xml.XmlNodeReader $XamlChildWindow)
$Basic_infos_Xaml        = [Windows.Markup.XamlReader]::Load($Childreader)

$XamlChildWindow = LoadXaml($viewFolder+"\Mapped_folders.xaml")
$Childreader     = (New-Object System.Xml.XmlNodeReader $XamlChildWindow)
$Mapped_folders_Xaml        = [Windows.Markup.XamlReader]::Load($Childreader)

$XamlChildWindow = LoadXaml($viewFolder+"\Startup_commands.xaml")
$Childreader     = (New-Object System.Xml.XmlNodeReader $XamlChildWindow)
$Startup_commands_Xaml     = [Windows.Markup.XamlReader]::Load($Childreader)

$XamlChildWindow = LoadXaml($viewFolder+"\Overview.xaml")
$Childreader     = (New-Object System.Xml.XmlNodeReader $XamlChildWindow)
$Overview_Xaml    = [Windows.Markup.XamlReader]::Load($Childreader)

 
$Basic_infos.Children.Add($Basic_infos_Xaml)        	 | Out-Null
$Mapped_folders.Children.Add($Mapped_folders_Xaml)        	 | Out-Null
$Startup_commands.Children.Add($Startup_commands_Xaml)  			 | Out-Null    
$Overview.Children.Add($Overview_Xaml)    		 | Out-Null  
     

#******************************************************
# Initialize with the first value of Item Section *****
#******************************************************

$TabMenuHamburger.SelectedItem = $TabMenuHamburger.ItemsSource[0]

$Load_Sandbox = $form.FindName("Load_Sandbox")
$Create_Sandbox = $form.FindName("Create_Sandbox")

#########################################################################
#                        HAMBURGER EVENTS                               #
#########################################################################

#******************* Items Section  *******************

# Controls for infos part
$sandbox_name = $Basic_infos_Xaml.FindName("sandbox_name")
$sandbox_path = $Basic_infos_Xaml.FindName("sandbox_path")
$sandbox_path_textbox = $Basic_infos_Xaml.FindName("sandbox_path_textbox")
$Shared_Multiple_Folder_Path = $Basic_infos_Xaml.FindName("Shared_Multiple_Folder_Path")
$Shared_Multiple_Folder_Path_Textbox = $Basic_infos_Xaml.FindName("Shared_Multiple_Folder_Path_Textbox")
$Choose_Network = $Basic_infos_Xaml.FindName("Choose_Network")
$Enable_Network = $Basic_infos_Xaml.FindName("Enable_Network")
$Disable_Network = $Basic_infos_Xaml.FindName("Disable_Network")
$Choose_vpgu = $Basic_infos_Xaml.FindName("Choose_vpgu")
$Enable_vpgu = $Basic_infos_Xaml.FindName("Enable_vpgu")
$Disable_vpgu = $Basic_infos_Xaml.FindName("Disable_vpgu")
$Run_Sandbox = $Basic_infos_Xaml.FindName("Run_Sandbox")
$OS_Version = $Basic_infos_Xaml.FindName("OS_Version")
$Sandbox_Status = $Basic_infos_Xaml.FindName("Sandbox_Status")
$OS_Warning_Block = $Basic_infos_Xaml.FindName("OS_Warning_Block")
$OS_Warning = $Basic_infos_Xaml.FindName("OS_Warning")

$Shared_Folder_Path = $Mapped_folders_Xaml.FindName("Shared_Folder_Path")
$Shared_Folder_Path_Textbox = $Mapped_folders_Xaml.FindName("Shared_Folder_Path_Textbox")
$DataGrid_Folders = $Mapped_folders_Xaml.FindName("DataGrid_Folders")
# $OneOrMultipleFolder = $Mapped_folders_Xaml.FindName("OneOrMultipleFolder")
$ReadOnlyOrNot = $Mapped_folders_Xaml.FindName("ReadOnlyOrNot")

$DataGrid_Command = $Startup_commands_Xaml.FindName("DataGrid_Command")
$command_path = $Startup_commands_Xaml.FindName("command_path")
$command_path_textbox = $Startup_commands_Xaml.FindName("command_path_textbox")
$Remove_Command = $Startup_commands_Xaml.FindName("Remove_Command")

$Command_File_Type = $Startup_commands_Xaml.FindName("Command_File_Type")
$File_PS1 = $Startup_commands_Xaml.FindName("File_PS1")
$File_VBS = $Startup_commands_Xaml.FindName("File_VBS")
$File_EXE = $Startup_commands_Xaml.FindName("File_EXE")
$File_MSI = $Startup_commands_Xaml.FindName("File_MSI")
$Browse_File_ToRun = $Startup_commands_Xaml.FindName("Browse_File_ToRun")
$Browse_File_ToRun_TextBox = $Startup_commands_Xaml.FindName("Browse_File_ToRun_TextBox")
$File_ToRun_Parameters = $Startup_commands_Xaml.FindName("File_ToRun_Parameters")

$SandBox_Overview = $Overview_Xaml.FindName("SandBox_Overview")

$OS_Warning_Block.Visibility = "Collapsed"




########################################################################################################################################################################################################	
#*******************************************************************************************************************************************************************************************************
# 																		VARIABLES INITIALIZATION 
#*******************************************************************************************************************************************************************************************************
########################################################################################################################################################################################################

$object = New-Object -comObject Shell.Application  
$Global:Current_Folder =(get-location).path 



#************************************************************************** TABCONTROLS ACTIONS ***************************************************************************************	
$Global:Launch_mode = "Details"	
	
$TabMenuHamburger.add_ItemClick({  
   $TabMenuHamburger.Content = $TabMenuHamburger.SelectedItem
   $TabMenuHamburger.IsPaneOpen = $false	

})		
		

$MonBouton.Add_Click({
	$Theme = [MahApps.Metro.ThemeManager]::DetectAppStyle($form)
	
	$my_theme = ($Theme.Item1).name
	If($my_theme -eq "BaseLight")
		{
			If($OS_ReleaseID -ge "1903")
				{
					$OS_Version.Foreground = "white"
				}
			Else
				{
					$OS_Version.Foreground = "yellow"
					$OS_Warning.Foreground = "yellow"
				}

			If($OS_ReleaseID -lt "1903")
				{
					$Sandbox_Status.Foreground = "yellow"
				}

			If($WindowsFeatureState -eq "Enabled")
				{
					$Sandbox_Status.Foreground = "white"
				}
			Else
				{
					$Sandbox_Status.Foreground = "yellow"
				}	
			[MahApps.Metro.ThemeManager]::ChangeAppStyle($form, $Theme.Item2, [MahApps.Metro.ThemeManager]::GetAppTheme("BaseDark"));		
				
		}
	ElseIf($my_theme -eq "BaseDark")
		{		
			If($OS_ReleaseID -ge "1903")
				{
					$OS_Version.Foreground = "green"
				}
			Else
				{
					$OS_Version.Foreground = "red"
					$OS_Warning.Foreground = "red"
				}

			If($OS_ReleaseID -lt "1903")
				{
					$Sandbox_Status.Foreground = "red"
				}

			If($WindowsFeatureState -eq "Enabled")
				{
					$Sandbox_Status.Foreground = "green"
				}
			Else
				{
					$Sandbox_Status.Foreground = "red"
				}				
			[MahApps.Metro.ThemeManager]::ChangeAppStyle($form, $Theme.Item2, [MahApps.Metro.ThemeManager]::GetAppTheme("BaseLight"));			
		}		
})		
		
		
		
$Remove_Command.Add_Click({
	$command_path_textbox.Text = ""
})		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		

$Choose_Network.Add_SelectionChanged({	
    If ($Enable_Network.IsSelected -eq $true) 
		{					
			$Script:Sandbox_Network_Status = "Enable"	
		} 
	Else 
		{	
			$Script:Sandbox_Network_Status = "Disable"				
		}
})	



$Choose_vpgu.Add_SelectionChanged({	
    If ($Enable_vpgu.IsSelected -eq $true) 
		{					
			$Script:Sandbox_VPGU_Status = "Enable"					
		} 
	Else 
		{	
			$Script:Sandbox_VPGU_Status = "Disable"								
		}
})	



# $RunSandbox.Add_Click({	
	# If ($RunSandbox.IsChecked -eq $true)
		# {
			# $Script:Sandbox_RunAfterCreate_Status = $true					
		# }
	# Else
		# {
			# $Script:Sandbox_RunAfterCreate_Status = $false				
		# }		
# })



$sandbox_path.Add_Click({
	Add-Type -AssemblyName System.Windows.Forms
	$SandBoxFolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
	[void]$SandBoxFolderBrowser.ShowDialog()	
	$sandbox_path_textbox.Text = $SandBoxFolderBrowser.SelectedPath
	$Script:Sandbox_Final_Path = $SandBoxFolderBrowser.SelectedPath		
	If($sandbox_path_textbox.Text -ne "")
		{
			$sandbox_path_textbox.BorderBrush = "gray"
			$sandbox_path_textbox.BorderThickness = "1"		
		}
	Else
		{
			$sandbox_path_textbox.BorderBrush = "Red"
			$sandbox_path_textbox.BorderThickness = "1"		
		}			
})


[System.Windows.RoutedEventHandler]$EventonDataGrid = {
    $button =  $_.OriginalSource.Name
    $Script:resultObj = $DataGrid_Folders.CurrentItem
    If ($button -match "Edit" ){
        EditFolder -rowObj $resultObj
    }
    ElseIf ($button -match "Remove" ){   
        RemoveFolder -rowObj $resultObj

    }
}
$DataGrid_Folders.AddHandler([System.Windows.Controls.Button]::ClickEvent, $EventonDataGrid)	



# [System.Windows.RoutedEventHandler]$EventonDataGrid = {
    # $button =  $_.OriginalSource.Name
    # $Script:resultObj = $DataGrid_Command.CurrentItem
    # If ($button -match "Edit" ){
        # EditCommand -rowObj $resultObj
    # }
    # ElseIf ($button -match "Remove" ){   
        # RemoveCommand -rowObj $resultObj

    # }
# }
# $DataGrid_Command.AddHandler([System.Windows.Controls.Button]::ClickEvent, $EventonDataGrid)	






function LoadXml ($global:filename)
{
	$XamlLoader=(New-Object System.Xml.XmlDocument)
	$XamlLoader.Load($filename)
	return $XamlLoader
}
$xamlDialog  = LoadXml("$Current_Folder\Dialog.xaml")
$read=(New-Object System.Xml.XmlNodeReader $xamlDialog)
$DialogForm=[Windows.Markup.XamlReader]::Load($read)

$CustomDialog = [MahApps.Metro.Controls.Dialogs.CustomDialog]::new($form)
$CustomDialog.AddChild($DialogForm)

$SaveAndClose_Dialog = $DialogForm.FindName("SaveAndClose_Dialog")
$Cancel = $DialogForm.FindName("Cancel")
$Set_Folder_Path = $DialogForm.FindName("Set_Folder_Path")
$Set_Folder_ReadOnly = $DialogForm.FindName("Set_Folder_ReadOnly")
$Set_Folder_ReadOnly_True = $DialogForm.FindName("Set_Folder_ReadOnly_True")
$Set_Folder_ReadOnly_False = $DialogForm.FindName("Set_Folder_ReadOnly_False")
					
$SaveAndClose_Dialog.add_Click({
	$Save_Set_Folder_Path = $Set_Folder_Path.Text.ToString()
	If($Set_Folder_ReadOnly_True.IsSelected -eq $true)
		{
			$Save_Set_Folder_ReadOnly = "true"
		}
		Else	
		{
			$Save_Set_Folder_ReadOnly = "false"
		}
			
	$resultObj.Path = $Save_Set_Folder_Path	
	$resultObj.ReadOnly = $Save_Set_Folder_ReadOnly	
	$DataGrid_Folders.Items.Refresh();						
    $CustomDialog.RequestCloseAsync()
})


$Cancel.add_Click({
    $CustomDialog.RequestCloseAsync()
})

Function EditFolder($rowObj)
	{     
		$Global:SandboxToEdit_Path = $rowObj.Path	
		$Global:SandboxToEdit_Access = $rowObj.ReadOnly						
		
		$Set_Folder_Path.Text = $SandboxToEdit_Path
		If($SandboxToEdit_Access -eq "true")
			{
				$Set_Folder_ReadOnly_True.IsSelected = $true
			}
		Else
			{
				$Set_Folder_ReadOnly_False.IsSelected = $true
			}		

		[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMetroDialogAsync($form, $CustomDialog)

	}	

Function RemoveFolder($rowObj)
	{     
		$DataGrid_Folders.Items.Remove($rowObj);
	}	


	

	
	
# function LoadXml ($global:filename)
# {
	# $XamlLoader=(New-Object System.Xml.XmlDocument)
	# $XamlLoader.Load($filename)
	# return $XamlLoader
# }
# $xamlDialog_Command  = LoadXml("$Current_Folder\Dialog_Command.xaml")
# $read_command=(New-Object System.Xml.XmlNodeReader $xamlDialog_Command)
# $Command_DialogForm=[Windows.Markup.XamlReader]::Load($read_command)

# $CustomCommandDialog = [MahApps.Metro.Controls.Dialogs.CustomDialog]::new($form)
# $CustomCommandDialog.AddChild($Command_DialogForm)

# $SaveAndClose_Command_Dialog = $Command_DialogForm.FindName("SaveAndClose_Command_Dialog")
# $Cancel_Command = $Command_DialogForm.FindName("Cancel_Command")
# $Set_Command = $Command_DialogForm.FindName("Set_Command")

# $SaveAndClose_Command_Dialog.add_Click({
	# $Save_Set_Command = $Set_Command.Text.ToString()
			
	# $resultObj.Command = $Save_Set_Command	
	# $DataGrid_Command.Items.Refresh();						
    # $CustomCommandDialog.RequestCloseAsync()
# })


# $Cancel_Command.add_Click({
    # $CustomCommandDialog.RequestCloseAsync()
# })

# Function EditCommand($rowObj)
	# {     			
		# $Script:CommandToEdit_Path = $rowObj.Command			
		# $Set_Command.Text = $CommandToEdit_Path			
		# [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMetroDialogAsync($form, $CustomCommandDialog)
	# }		
	
	
# Function RemoveCommand($rowObj)
	# {     
		# $DataGrid_Command.Items.Remove($rowObj);
	# }		
	
	
	
	
	
	
	
	
	
	
	



$Shared_Folder_Path.Add_Click({
	Add-Type -AssemblyName System.Windows.Forms
	$ShareSandBoxFolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
	[void]$ShareSandBoxFolderBrowser.ShowDialog()	
	$Script:New_Sandbox_Path = $ShareSandBoxFolderBrowser.SelectedPath		
	
	If($New_Sandbox_Path -ne "")
		{
			If ($ReadOnlyOrNot.IsChecked -eq $true)
				{
					$Script:Shared_Folder_Access = "false"					
				}
			Else
				{
					$Script:Shared_Folder_Access = "true"				
				}

			# If ($OneOrMultipleFolder.IsChecked -eq $true)
				# {
					# $Get_Folder_Item = get-childitem $New_Sandbox_Path -Directory
					# ForEach($Folder in $Get_Folder_Item)
						# {
							# $MappedFolders_values = New-Object PSObject											
							# $Full_Folder_Path = $Folder.FullName 
							# $MappedFolders_values = $MappedFolders_values | Add-Member NoteProperty Path $Full_Folder_Path -passthru
							# $MappedFolders_values = $MappedFolders_values | Add-Member NoteProperty ReadOnly $Shared_Folder_Access -passthru							
							# $DataGrid_Folders.Items.Add($MappedFolders_values) > $null								
						# }
				# }
			# Else
				# {
					$MappedFolders_values = New-Object PSObject
					$MappedFolders_values = $MappedFolders_values | Add-Member NoteProperty Path $New_Sandbox_Path -passthru
					$MappedFolders_values = $MappedFolders_values | Add-Member NoteProperty ReadOnly $Shared_Folder_Access -passthru							
					$DataGrid_Folders.Items.Add($MappedFolders_values) > $null						
				# }					
		}
})





# $command_path.Add_Click({
	# $Script:New_Command = $command_path_textbox.Text.ToString()		
	
	# If($New_Command -ne "")
		# {
			# $Commands_values = New-Object PSObject
			# $Commands_values = $Commands_values | Add-Member NoteProperty Command $New_Command -passthru
			# $DataGrid_Command.Items.Add($Commands_values) > $null		
		# }
# })















$Win32_OperatingSystem = Get-ciminstance Win32_OperatingSystem 
$REG_OS_Version = get-itemproperty -path registry::"HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -erroraction 'silentlycontinue'
$Build_number = $Win32_OperatingSystem.buildnumber
$OS_ReleaseID =  $REG_OS_Version.ReleaseID
$OS_Version.Content = "$Build_number ($OS_ReleaseID)"
If($OS_ReleaseID -ge "1903")
	{
		$OS_Version.Foreground = "Green"
		$OS_Version.FontWeight = "Bold"	
		$OS_Warning_Block.Visibility = "Collapsed"		
	}
Else
	{
		$OS_Version.Foreground = "red"		
		$OS_Version.FontWeight = "Bold"		
		$OS_Warning_Block.Visibility = "Visible"	
		$OS_Warning.Content = "Requires at least 1903"
		$OS_Warning.Foreground = "red"				
		$OS_Warning.FontWeight = "Bold"				
	}


If($OS_ReleaseID -lt "1903")
	{
		$Sandbox_Status.Content = "Disabled"		
		$Sandbox_Status.Foreground = "red"
		$Sandbox_Status.FontWeight = "Bold"		
	}

$Sandbox_Feature = "Containers-DisposableClientVM"
$WindowsFeatureState = (Get-WindowsOptionalFeature -FeatureName $Sandbox_Feature -Online).State 
If($WindowsFeatureState -eq "Enabled")
	{
		$Sandbox_Status.Content = "Enabled"
		$Sandbox_Status.Foreground = "Green"
		$Sandbox_Status.FontWeight = "Bold"
	}
Else
	{
		$Sandbox_Status.Content = "Disabled"
		$Sandbox_Status.Foreground = "red"		
		$Sandbox_Status.FontWeight = "Bold"		
	}	

$Script:Load_WSB_Status = $false
$Load_Sandbox.Add_Click({
	$OpenFileDialog1 = New-Object System.Windows.Forms.OpenFileDialog
	$openfiledialog1.Filter = "WSB File (.wsb)|*.wsb;"
	$openfiledialog1.title = "Select the WSB file to upload"	
	$openfiledialog1.ShowHelp = $true	
	$OpenFileDialog1.initialDirectory = [Environment]::GetFolderPath("Desktop")
	$OpenFileDialog1.ShowDialog() | Out-Null	
	$Script:Sandbox_Final_Path = $OpenFileDialog1.filename	
	$Script:Sandbox_Short_Name = [System.IO.Path]::GetFileNameWithoutExtension($Sandbox_Final_Path)	
	$Script:Sandbox_Configuration = $Sandbox_Final_Path						
	$Input_Configuration = [xml] (Get-Content $Sandbox_Configuration)	
	
	$Load_Sandbox_Networking = $Input_Configuration.Configuration.Networking
	$Load_Sandbox_vpgu = $Input_Configuration.Configuration.VGpu
	$Load_Sandbox_MappedFolders = $Input_Configuration.Configuration.MappedFolders.MappedFolder
	$Load_Sandbox_Commands = $Input_Configuration.Configuration.LogonCommand.Command
	
	$sandbox_path_textbox.Text = $Sandbox_Final_Path	
	$sandbox_name.Text = $Sandbox_Short_Name


	If($Sandbox_Final_Path -ne $null)
		{
			$Script:Load_WSB_Status = $true
		}
			
	$Create_Sandbox.Content = "Save existing Sandbox"
	
	If($Load_Sandbox_Networking -eq "Enable")
		{
			$Enable_Network.IsSelected = $true
		}
	Else
		{
			$Disable_Network.IsSelected = $true		
		}
		
	If($Load_Sandbox_vpgu -eq "Enable")
		{
			$Enable_vpgu.IsSelected = $true
		}
	Else
		{
			$Disable_vpgu.IsSelected = $true		
		}
		
	foreach ($MappedFolder in $Load_Sandbox_MappedFolders)
		{
			$MappedFolders_values = New-Object PSObject
			$MappedFolders_values = $MappedFolders_values | Add-Member NoteProperty Path $MappedFolder.HostFolder -passthru
			$MappedFolders_values = $MappedFolders_values | Add-Member NoteProperty ReadOnly $MappedFolder.ReadOnly -passthru							
			$DataGrid_Folders.Items.Add($MappedFolders_values) > $null
		}		

	$command_path_textbox.Text = $Load_Sandbox_Commands		
})




$Browse_File_ToRun.Add_Click({
	$OpenFileDialog_FileToRun = New-Object System.Windows.Forms.OpenFileDialog
		
	If($File_PS1.IsSelected -eq $True)
		{
			$OpenFileDialog_FileToRun.Filter = "PS1 File (.ps1)|*.ps1;"
			$OpenFileDialog_FileToRun.title = "Select the PS1 file to upload"			
		}
	ElseIf($File_VBS.IsSelected -eq $True)
		{
			$OpenFileDialog_FileToRun.Filter = "VBS File (.vbs)|*.vbs;"
			$OpenFileDialog_FileToRun.title = "Select the VBS file to upload"			
		}
	ElseIf($File_EXE.IsSelected -eq $True)
		{
			$OpenFileDialog_FileToRun.Filter = "EXE File (.exe)|*.exe;"
			$OpenFileDialog_FileToRun.title = "Select the EXE file to upload"			
		}
	ElseIf($File_MSI.IsSelected -eq $True)
		{
			$OpenFileDialog_FileToRun.Filter = "MSI File (.msi)|*.msi;"
			$OpenFileDialog_FileToRun.title = "Select the MSI file to upload"			
		}		
		
	$OpenFileDialog_FileToRun.ShowHelp = $True	
	$OpenFileDialog_FileToRun.initialDirectory = [Environment]::GetFolderPath("Desktop")
	$OpenFileDialog_FileToRun.ShowDialog() | Out-Null	
	$Script:FileToRun_Full_Path = $OpenFileDialog_FileToRun.FileName		
	$Script:FileToRun_Final_Path = $OpenFileDialog_FileToRun.SafeFileName	
	$Browse_File_ToRun_TextBox.Text = $FileToRun_Final_Path
	$Script:Paramaters_Switches = $File_ToRun_Parameters.Text.ToString()

	If($New_Sandbox_Path -ne $null)
		{
			$FolderPath = (get-item $New_Sandbox_Path).Name
		}
	Else
		{
			$FolderPath = (get-item $FileToRun_Full_Path).DirectoryName
			$FolderPath = (get-item $FolderPath).Name		
		}
		
	$Sandbox_Shared_Path = "C:\Users\WDAGUtilityAccount\Desktop\$FolderPath"
	$Full_Startup_Path = "$Sandbox_Shared_Path\$FileToRun_Final_Path"
	$Full_Startup_Path = """$Full_Startup_Path"""	
	
	If($File_PS1.IsSelected -eq $True)
		{
			$Script:Startup_Command = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -sta -WindowStyle Hidden -noprofile -executionpolicy unrestricted -file $Full_Startup_Path $Paramaters_Switches"						
		}
	ElseIf($File_VBS.IsSelected -eq $True)
		{
			$Script:Startup_Command = "wscript.exe $Full_Startup_Path $Paramaters_Switches"					
		}
	ElseIf($File_EXE.IsSelected -eq $True)
		{
			$Script:Startup_Command = "start """" $Full_Startup_Path " + $Paramaters_Switches					
		}
	ElseIf($File_MSI.IsSelected -eq $True)
		{
			$Script:Startup_Command = "start """" $Full_Startup_Path " + $Paramaters_Switches					
		}	
		
	$command_path_textbox.Text = $Startup_Command		
})




$File_ToRun_Parameters.Add_TextChanged({		
	$command_path_textbox.Text = $Startup_Command + $File_ToRun_Parameters.Text.ToString()
})


Function Generate_Sandbox_Overview
	{
		If($Enable_Network.IsSelected -eq $True)
			{
				$Overview_Network = "Enable"
			}
		Else
			{
				$Overview_Network = "Disable"
			}	

		If($Enable_vpgu.IsSelected -eq $True)
			{
				$Overview_vpgu = "Enable"
			}
		Else
			{
				$Overview_vpgu = "Disable"
			}				

		$Script:Overview_Sandbox = ""
		$Overview_Sandbox += "<Configuration>"	
		$Overview_Sandbox += "`n    <VGpu>$Overview_vpgu</VGpu>"					
		$Overview_Sandbox += "`n    <Networking>$Overview_Network</Networking>"		
		$Overview_Sandbox += "`n    <MappedFolders>"	
		ForEach($Folder in $DataGrid_Folders.items)
			{
				$Folder_Path = $Folder.path
				$Folder_ReadOnly = $Folder.ReadOnly
				
				$Overview_Sandbox += "`n        <MappedFolder>"	
				$Overview_Sandbox += "`n              <HostFolder>$Folder_Path</HostFolder>"		
				$Overview_Sandbox += "`n              <ReadOnly>$Folder_ReadOnly</ReadOnly>"																										
				$Overview_Sandbox += "`n        </MappedFolder>"																		
			}		
		$Overview_Sandbox += "`n    </MappedFolders>"	

		$Overview_Commandline = $command_path_textbox.Text.ToString()
		$Overview_Sandbox += "`n    <LogonCommand>"			
		$Overview_Sandbox += "`n    <Command>$Overview_Commandline</Command>"							
		$Overview_Sandbox += "`n    </LogonCommand>"			

		$Overview_Sandbox += "`n</Configuration>"	
		$SandBox_Overview.Text = $Overview_Sandbox

		If($Export -eq $true)
			{
				$SandboxToCreate_Path = $sandbox_path_textbox.Text.ToString()
				$SandboxToCreate_Name = $sandbox_name.Text.ToString()	
				$Sandbox_full_path = "$SandboxToCreate_Path\$SandboxToCreate_Name.wsb"				

				If($Load_WSB_Status -eq $True)
					{
						$Sandbox_To_Run = $SandboxToCreate_Path
						$Overview_Sandbox | out-file $Sandbox_To_Run 
					}
				Else
					{
						$Sandbox_To_Run = $Sandbox_full_path					
						$Overview_Sandbox | out-file $Sandbox_To_Run 
					}

				If ($Run_Sandbox.IsChecked -eq $True)
					{
						[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Form, "Success :-)", "Your Sandbox has been created and will be launched automatically.")																												
						& $Sandbox_To_Run	
					}	
				Else
					{
						[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Form, "Success :-)", "Your Sandbox has been created")																												
					}
			}
	}
	
	
	
$Script:Export = $false
$Create_Sandbox.Add_Click({	
	$SandboxToCreate_Path = $sandbox_path_textbox.Text.ToString()
	If($SandboxToCreate_Path -eq "")
		{
			$sandbox_path_textbox.BorderBrush = "Red"
			$sandbox_path_textbox.BorderThickness = "1"	
			[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Form, "Oops :-(", "Please specify a sandbox path")								
		}
	Else
		{
			$Script:Export = $true			
			Generate_Sandbox_Overview				
		}
			

})
	

	
	
$TabMenuHamburger.add_ItemClick({  
   $TabMenuHamburger.Content = $TabMenuHamburger.SelectedItem
   $TabMenuHamburger.IsPaneOpen = $false	
	If ($TabMenuHamburger.SelectedItem.Label -eq "Overview")
		{
			Generate_Sandbox_Overview				
		}	

})		
		
		
		
		
		
		
		
		
		
		
		
		


$Form.ShowDialog() | Out-Null
	 
	 