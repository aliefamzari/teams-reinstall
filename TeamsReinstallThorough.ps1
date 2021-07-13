##### Script based on the great works of XSIOL. #####
##### Formatting, commenting, modifications and merging of the scripts done by TOBKO. #####
##### Mod to unantended remote reinstallation by ALMAZ ####

$ErrorActionPreference="SilentlyContinue"
$ConfirmPreference="False"
#Admin check
$pc_name = Read-Host -Prompt "Enter PC Name" 
$user = "$env:USERNAME";
$group = "Administrators";
$groupObj =[ADSI]"WinNT://$pc_name/$group,group" 
$membersObj = @($groupObj.psbase.Invoke("Members")) 
$members = ($membersObj | ForEach-Object {$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)})
If ($members -contains $user) {
      Write-Host "$user is admin on $pc_name"
 } Else {
        Write-Host "$user is not admin on $pc_name"
        exit 
}
#Prepping variables for remote invoke
$pcname=Read-Host "Enter PC Name"
$logonuser=Get-WmiObject -ComputerName $pcname -Class Win32_ComputerSystem | Select-Object UserName
$sam=$logonuser.UserName.split('\')[1]
$rENVLAD="c:\users\$sam\localappdata"
$rENVAD="c:\Users\$sam\AppData\Roaming"
$rENVUP="c:\users\$sam"


    ##################################
    ##################################
    ##### Teams Uninstall Script #####
    #####        By XSIOL        #####
    ##################################
    ##################################
    
    # Stops Microsoft Teams
    Write-Host "Stopping Teams Process" -ForegroundColor Yellow

    try{
        Get-Process -ProcessName Teams  | Stop-Process -Force
        Start-Sleep -Seconds 3
        Write-Host "Teams Process Sucessfully Stopped" -ForegroundColor Green
    }
    
    catch{
        echo $_
    }

    # Starts uninstall
    $TeamsPath = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Microsoft', 'Teams')
    $TeamsUpdateExePath = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'Microsoft', 'Teams', 'Update.exe')
    try
    {
        if (Test-Path -Path $TeamsUpdateExePath) {
            Write-Host "Uninstalling Teams process"
            # Uninstall app
            $proc = Start-Process -FilePath $TeamsUpdateExePath -ArgumentList "-uninstall -s" -PassThru
            $proc.WaitForExit()
        }
        if (Test-Path -Path $TeamsPath) {
            Write-Host "Deleting Teams directory"
            Remove-Item �Path $TeamsPath -Recurse
        }
    }
    catch
    {
        Write-Error -ErrorRecord $_
    }
    
    ##############################
    ##############################
    ##### Cache Clear Script #####
    #####      By XSIOL      #####
    ##############################
    ##############################
    
    Write-Host "Clearing Teams Disk Cache" -ForegroundColor Yellow
    
    try{
        Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\application cache\cache"  | Remove-Item  
        Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\blob_storage"  | Remove-Item  
        Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\databases"  | Remove-Item  
        Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\cache"  | Remove-Item  
        Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\gpucache"  | Remove-Item  
        Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\Indexeddb"  | Remove-Item  
        Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\Local Storage"  | Remove-Item  
        Get-ChildItem -Path $env:APPDATA\"Microsoft\teams\tmp"  | Remove-Item  
        Write-Host "Teams Disk Cache Cleaned" -ForegroundColor Green
    }
    
    catch{
        echo $_
    }
    
    Write-Host "Stopping Chrome Process" -ForegroundColor Yellow

    try{
        Get-Process -ProcessName Chrome  | Stop-Process -Force 
        Start-Sleep -Seconds 3
        Write-Host "Chrome Process Sucessfully Stopped" -ForegroundColor Green
    }
    
    catch{
        echo $_
    }

    Write-Host "Clearing Chrome Cache" -ForegroundColor Yellow
    
    try{
        Get-ChildItem -Path $env:LOCALAPPDATA"\Google\Chrome\User Data\Default\Cache"  | Remove-Item  
        Get-ChildItem -Path $env:LOCALAPPDATA"\Google\Chrome\User Data\Default\Cookies" -File  | Remove-Item  
        Get-ChildItem -Path $env:LOCALAPPDATA"\Google\Chrome\User Data\Default\Web Data" -File  | Remove-Item  
        Write-Host "Chrome Cleaned" -ForegroundColor Green
    }
    
    catch{
        echo $_
    }
    
    Write-Host "Stopping IE Process" -ForegroundColor Yellow
    
    try{
        Get-Process -ProcessName MicrosoftEdge  | Stop-Process -Force 
        Get-Process -ProcessName MSEdge  | Stop-Process -Force 
        Get-Process -ProcessName IExplore  | Stop-Process -Force 
        Write-Host "Internet Explorer and Edge Processes Sucessfully Stopped" -ForegroundColor Green
    }
    
    catch{
        echo $_
    }

    Write-Host "Clearing IE & Edge Cache" -ForegroundColor Yellow
    
    try{
        RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 8
        RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 2
        Start-Sleep 3
        Write-Host "IE and Edge Cleaned" -ForegroundColor Green
    }
    
    catch{
        echo $_
    }

    Write-Host "Cleanup Complete..." -ForegroundColor Green


    ##################################
    ##################################
    ##### Teams Reinstall Script #####
    #####        By TOBKO        ##### 
    ##################################
    ##################################
    
    # Make a Teams install that automatically downloads the installer
    $ExeFolder = "$ENV:USERPROFILE\Downloads"
    $DownloadSource = "https://go.microsoft.com/fwlink/p/?LinkID=869426&clcid=0x409&culture=en-us&country=US&lm=deeplink&lmsrc=groupChatMarketingPageWeb&cmpid=directDownloadWin64"
    $ExeDestination = "$ExeFolder\Teams_windows_x64.exe"

    If([System.IO.File]::Exists($ExeDestination) -eq $false){
        Write-Host "Downloading Teams, please wait." -ForegroundColor Red
        Invoke-WebRequest $DownloadSource -OutFile $ExeDestination
    }
    Else{
        Write-Host "Installer file already present in Downloads folder. Skipping download." -ForegroundColor Red
    }

    Write-Host "Installing Teams" -ForegroundColor Magenta
    $proc = Start-Process -FilePath $ExeDestination -ArgumentList "-s" -PassThru
    $proc.WaitForExit()

    Start-Sleep 5

    Write-Host "Checking install" -ForegroundColor Magenta
    $proc = Start-Process -FilePath $ExeDestination -ArgumentList "-s" -PassThru
    $proc.WaitForExit()

    Start-Process -FilePath $env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe
    
    Stop-Process -Id $PID
}

else{
    Write-Host "Not a valid input, stopping script"
    Start-Sleep -s 6
    Stop-Process -Id $PID
}