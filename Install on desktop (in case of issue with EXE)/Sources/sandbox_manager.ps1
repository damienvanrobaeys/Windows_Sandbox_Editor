#================================================================================================================
#
# Author 		 : Damien VAN ROBAEYS
#
#================================================================================================================

$Global:Current_Folder = split-path $MyInvocation.MyCommand.Path


[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')  				| out-null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') 				| out-null
[System.Reflection.Assembly]::LoadFrom("$Current_Folder\MahApps.Metro.dll")       				| out-null
[System.Reflection.Assembly]::LoadFrom("$Current_Folder\MahApps.Metro.IconPacks.dll")      | out-null  

function LoadXml ($global:filename)
{
    $XamlLoader=(New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($filename)
    return $XamlLoader
}

# Load MainWindow
$XamlMainWindow=LoadXml("$Current_Folder\sandbox_manager.xaml")
$Reader=(New-Object System.Xml.XmlNodeReader $XamlMainWindow)
$Form=[Windows.Markup.XamlReader]::Load($Reader)

[System.Windows.Forms.Application]::EnableVisualStyles()


########################################################################################################################################################################################################	
#*******************************************************************************************************************************************************************************************************
# 																		BUTTONS AND LABELS INITIALIZATION 
#*******************************************************************************************************************************************************************************************************
########################################################################################################################################################################################################

#************************************************************************** DATAGRID *******************************************************************************************************************
$Sandbox_TabControl = $form.FindName("Sandbox_TabControl")
$sandbox_name = $form.FindName("sandbox_name")
$sandbox_path = $form.FindName("sandbox_path")
$sandbox_path_textbox = $form.FindName("sandbox_path_textbox")
$Shared_Multiple_Folder_Path = $form.FindName("Shared_Multiple_Folder_Path")
$Shared_Multiple_Folder_Path_Textbox = $form.FindName("Shared_Multiple_Folder_Path_Textbox")
$Choose_Network = $form.FindName("Choose_Network")
$Enable_Network = $form.FindName("Enable_Network")
$Disable_Network = $form.FindName("Disable_Network")
$Choose_vpgu = $form.FindName("Choose_vpgu")
$Enable_vpgu = $form.FindName("Enable_vpgu")
$Disable_vpgu = $form.FindName("Disable_vpgu")

$Enable_Clipboard = $form.FindName("Enable_Clipboard")
$Disable_Clipboard = $form.FindName("Disable_Clipboard")

$Enable_ProtectedClient = $form.FindName("Enable_ProtectedClient")
$Disable_ProtectedClient = $form.FindName("Disable_ProtectedClient")

$Enable_Pinter = $form.FindName("Enable_Pinter")
$Disable_Printer = $form.FindName("Disable_Printer")

$Choose_Printer = $form.FindName("Choose_Printer")
$Choose_ProtectedClient = $form.FindName("Choose_ProtectedClient")
$Choose_Clipboard = $form.FindName("Choose_Clipboard")
$Choose_Memory = $form.FindName("Choose_Memory")

$Run_Sandbox = $form.FindName("Run_Sandbox")
 
$dark = $form.FindName("dark")
$light = $form.FindName("light")
$MonBouton = $form.FindName("MonBouton")

$Shared_Folder_Path = $form.FindName("Shared_Folder_Path")
$Shared_Folder_Path_Textbox = $form.FindName("Shared_Folder_Path_Textbox")
$DataGrid_Folders = $form.FindName("DataGrid_Folders")
$ReadOnlyOrNot = $form.FindName("ReadOnlyOrNot")

$Command_File_Type = $form.FindName("Command_File_Type")
$File_PS1 = $form.FindName("File_PS1")
$File_VBS = $form.FindName("File_VBS")
$File_EXE = $form.FindName("File_EXE")
$File_MSI = $form.FindName("File_MSI")
$Browse_File_ToRun = $form.FindName("Browse_File_ToRun")
$Browse_File_ToRun_TextBox = $form.FindName("Browse_File_ToRun_TextBox")
$command_path = $form.FindName("command_path")
$command_path_textbox = $form.FindName("command_path_textbox")
$Remove_Command = $form.FindName("Remove_Command")
$File_ToRun_Parameters = $form.FindName("File_ToRun_Parameters")

$Load_Sandbox = $form.FindName("Load_Sandbox")
$Create_Sandbox = $form.FindName("Create_Sandbox")

$SandBox_Overview = $form.FindName("SandBox_Overview")

$MonBouton.Add_Click({
	$Theme = [MahApps.Metro.ThemeManager]::DetectAppStyle($form)
	
	$my_theme = ($Theme.Item1).name
	If($my_theme -eq "BaseLight")
		{
			[MahApps.Metro.ThemeManager]::ChangeAppStyle($form, $Theme.Item2, [MahApps.Metro.ThemeManager]::GetAppTheme("BaseDark"));		
				
		}
	ElseIf($my_theme -eq "BaseDark")
		{					
			[MahApps.Metro.ThemeManager]::ChangeAppStyle($form, $Theme.Item2, [MahApps.Metro.ThemeManager]::GetAppTheme("BaseLight"));			
		}		
})

$Remove_Command.Add_Click({
	$command_path_textbox.Text = ""
})

$Choose_Printer.Add_SelectionChanged({	
    If ($Enable_Pinter.IsSelected -eq $true) 
		{					
			$Script:Sandbox_Printer_Status = "Enable"	
		} 
	Else 
		{	
			$Script:Sandbox_Printer_Status = "Disable"				
		}
})	

$Choose_ProtectedClient.Add_SelectionChanged({	
    If ($Enable_ProtectedClient.IsSelected -eq $true) 
		{					
			$Script:Sandbox_ProtectedClient_Status = "Enable"	
		} 
	Else 
		{	
			$Script:Sandbox_ProtectedClient_Status = "Disable"				
		}
})	

$Choose_Clipboard.Add_SelectionChanged({	
    If ($Enable_Clipboard.IsSelected -eq $true) 
		{					
			$Script:Sandbox_Clipboard_Status = "Enable"	
		} 
	Else 
		{	
			$Script:Sandbox_Clipboard_Status = "Disable"				
		}
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
	If($Set_Folder_ReadOnly_True.IsSelected -eq $True)
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
				$Set_Folder_ReadOnly_True.IsSelected = $True
			}
		Else
			{
				$Set_Folder_ReadOnly_False.IsSelected = $True
			}		

		[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMetroDialogAsync($form, $CustomDialog)

	}	

Function RemoveFolder($rowObj)
	{     
		$DataGrid_Folders.Items.Remove($rowObj);
	}	


$Shared_Folder_Path.Add_Click({
	Add-Type -AssemblyName System.Windows.Forms
	$ShareSandBoxFolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
	[void]$ShareSandBoxFolderBrowser.ShowDialog()	
	$Script:New_Sandbox_Path = $ShareSandBoxFolderBrowser.SelectedPath		
	
	If($New_Sandbox_Path -ne "")
		{
			If ($ReadOnlyOrNot.IsChecked -eq $True)
				{
					$Script:Shared_Folder_Access = "false"					
				}
			Else
				{
					$Script:Shared_Folder_Access = "true"				
				}

			$MappedFolders_values = New-Object PSObject
			$MappedFolders_values = $MappedFolders_values | Add-Member NoteProperty Path $New_Sandbox_Path -passthru
			$MappedFolders_values = $MappedFolders_values | Add-Member NoteProperty ReadOnly $Shared_Folder_Access -passthru							
			$DataGrid_Folders.Items.Add($MappedFolders_values) > $null						
		}
})

$Script:Load_WSB_Status = $False
$Load_Sandbox.Add_Click({
	$OpenFileDialog1 = New-Object System.Windows.Forms.OpenFileDialog
	$openfiledialog1.Filter = "WSB File (.wsb)|*.wsb;"
	$openfiledialog1.title = "Select the WSB file to upload"	
	$openfiledialog1.ShowHelp = $True	
	$OpenFileDialog1.initialDirectory = [Environment]::GetFolderPath("Desktop")
	$OpenFileDialog1.ShowDialog() | Out-Null	
	$Script:Sandbox_Final_Path = $OpenFileDialog1.filename	
	$Script:Sandbox_Short_Name = [System.IO.Path]::GetFileNameWithoutExtension($Sandbox_Final_Path)	
	$Script:Sandbox_Configuration = $Sandbox_Final_Path						
	$Input_Configuration = [xml] (Get-Content $Sandbox_Configuration)	
	
	$Load_Sandbox_Networking = $Input_Configuration.Configuration.Networking
	$Load_Sandbox_vpgu = $Input_Configuration.Configuration.VGpu	
	$Load_Sandbox_ClipboardRedirection = $Input_Configuration.Configuration.ClipboardRedirection
	$Load_Sandbox_ProtectedClient = $Input_Configuration.Configuration.ProtectedClient
	$Load_Sandbox_PrinterRedirection = $Input_Configuration.Configuration.PrinterRedirection
	$Load_Sandbox_MemoryInMB = $Input_Configuration.Configuration.MemoryInMB	
	$Load_Sandbox_MappedFolders = $Input_Configuration.Configuration.MappedFolders.MappedFolder
	$Load_Sandbox_Commands = $Input_Configuration.Configuration.LogonCommand.Command
	$Load_Sandbox_MemoryInMB = $Input_Configuration.Configuration.MemoryInMB
	
	$sandbox_path_textbox.Text = $Sandbox_Final_Path	
	$sandbox_name.Text = $Sandbox_Short_Name

	If($Load_Sandbox_MemoryInMB -ne $null)
		{
			$Choose_Memory.Value = $Load_Sandbox_MemoryInMB
		}

	If($Sandbox_Final_Path -ne $null)
		{
			$Script:Load_WSB_Status = $True
		}
			
	$Create_Sandbox.Content = "Save existing Sandbox"
	If(($Load_Sandbox_ClipboardRedirection -eq $null) -or ($Load_Sandbox_ClipboardRedirection -eq "Enable"))
		{
			$Enable_Clipboard.IsSelected = $True			
		}
	ElseIf($Load_Sandbox_ClipboardRedirection -eq "Disable")
		{
			$Disable_Clipboard.IsSelected = $True		
		}	
	
	
	If(($Load_Sandbox_ProtectedClient -eq "Enable") -or ($Load_Sandbox_ProtectedClient -eq $null))
		{
			$Enable_ProtectedClient.IsSelected = $True
		}
	ElseIf($Load_Sandbox_ProtectedClient -eq "Disable")
		{
			$Disable_ProtectedClient.IsSelected = $True		
		}	
		
	If(($Load_Sandbox_PrinterRedirection -eq "Enable") -or ($Load_Sandbox_PrinterRedirection -eq $null))
		{
			$Enable_Pinter.IsSelected = $True
		}
	ElseIf($Load_Sandbox_PrinterRedirection -eq "Disable")
		{
			$Disable_Printer.IsSelected = $True		
		}				
	
	If(($Load_Sandbox_Networking -eq "Enable") -or ($Load_Sandbox_Networking -eq $null))
		{
			$Enable_Network.IsSelected = $True
		}
	ElseIf($Load_Sandbox_Networking -eq "Disable")
		{
			$Disable_Network.IsSelected = $True		
		}
		
	If(($Load_Sandbox_vpgu -eq "Enable") -or ($Load_Sandbox_vpgu -eq $null))
		{
			$Enable_vpgu.IsSelected = $True
		}
	ElseIf($Load_Sandbox_vpgu -eq "Disable")
		{
			$Disable_vpgu.IsSelected = $True		
		}
		
	foreach ($MappedFolder in $Load_Sandbox_MappedFolders)
		{
			$MappedFolders_values = New-Object PSObject
			$MappedFolders_values = $MappedFolders_values | Add-Member NoteProperty Path $MappedFolder.HostFolder -passthru
			$MappedFolders_values = $MappedFolders_values | Add-Member NoteProperty ReadOnly $MappedFolder.ReadOnly -passthru							
			$DataGrid_Folders.Items.Add($MappedFolders_values) > $null
		}		

	$command_path_textbox.Text = $Load_Sandbox_Commands	
	
	If($command_path_textbox.Text -like "*.ps1*")
		{
			$File_PS1.IsSelected = $true
		}
	ElseIf ($command_path_textbox.Text -like "*.vbs*")
		{
			$File_VBS.IsSelected = $true		
		}
	ElseIf ($command_path_textbox.Text -like "*.exe*")
		{
			$File_EXE.IsSelected = $true		
		}
	ElseIf ($command_path_textbox.Text -like "*.msi*")
		{
			$File_MSI.IsSelected = $true		
		}		
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
			$Script:Startup_Command = "$Full_Startup_Path $Paramaters_Switches"						
		}
	ElseIf($File_MSI.IsSelected -eq $True)
		{
			$Script:Startup_Command = "msiexec /i $Full_Startup_Path " + $Paramaters_Switches			
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

		If($Enable_Clipboard.IsSelected -eq $True)
			{
				$Overview_Clipboard = "Enable"
			}
		Else
			{
				$Overview_Clipboard = "Disable"
			}		

		If($Enable_ProtectedClient.IsSelected -eq $True)
			{
				$Overview_ProtectedClient = "Enable"
			}
		Else
			{
				$Overview_ProtectedClient = "Disable"
			}		

		If($Enable_Pinter.IsSelected -eq $True)
			{
				$Overview_Printer = "Enable"
			}
		Else
			{
				$Overview_Printer = "Disable"
			}				
			
		$Overview_MemoryInMB = $Choose_Memory.value

		$Script:Overview_Sandbox = ""
		$Overview_Sandbox += "<Configuration>"	
		$Overview_Sandbox += "`n    <VGpu>$Overview_vpgu</VGpu>"					
		$Overview_Sandbox += "`n    <Networking>$Overview_Network</Networking>"		
		$Overview_Sandbox += "`n    <ClipboardRedirection>$Overview_Clipboard</ClipboardRedirection>"		
		$Overview_Sandbox += "`n    <ProtectedClient>$Overview_ProtectedClient</ProtectedClient>"		
		$Overview_Sandbox += "`n    <PrinterRedirection>$Overview_Printer</PrinterRedirection>"		
		$Overview_Sandbox += "`n    <MemoryInMB>$Overview_MemoryInMB</MemoryInMB>"				
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
			$sandbox_path_textbox.BorderBrush = "red"
			$sandbox_path_textbox.BorderThickness = "1"		
			[MahApps.Metro.Controls.Dialogs.DialogManager]::ShowMessageAsync($Form, "Oops :-(", "Please specify a sandbox path")								
		}
	Else
		{
			$Script:Export = $true			
			Generate_Sandbox_Overview				
		}		
})
	

$Sandbox_TabControl.Add_SelectionChanged({	
	If ($Sandbox_TabControl.SelectedItem.Header -eq "Overview")
		{
			Generate_Sandbox_Overview		
		}		
})


$Form.ShowDialog() | Out-Null

