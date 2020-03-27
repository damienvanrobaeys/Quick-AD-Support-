[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')  	 | out-null
[System.Windows.Forms.Application]::EnableVisualStyles()

$Current_Folder =(get-location).path 
$QADS_Folder = "$Current_Folder\Quick_AD_Support"
$ProgData = $env:PROGRAMDATA
copy-item $QADS_Folder $ProgData -force -recurse

$QADS_Final_Path = "$ProgData\Quick_AD_Support"
If (test-path $QADS_Final_Path)
	{
		Try
			{
				Get-ChildItem $QADS_Final_Path -Recurse | Unblock-File	
				[System.Windows.Forms.MessageBox]::Show("Folder has been copied and unblocked. `nThe tool is now in your systray. `n`nNote: The tool won't be opened directly, run it through the systray !!!")			
				cd "C:\ProgramData\Quick_AD_Support"
				powershell -windowstyle hidden ".\QADS.ps1"				
			}
		Catch
			{}
	}
Else
	{
		[System.Windows.Forms.MessageBox]::Show("The folder $QADS_Final_Path does not exist")				
	}


