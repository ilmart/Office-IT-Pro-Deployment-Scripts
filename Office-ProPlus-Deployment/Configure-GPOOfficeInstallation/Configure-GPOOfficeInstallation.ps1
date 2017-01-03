try {
$enumDef = "
using System;
       [FlagsAttribute]
       public enum Bitness
       {
          Both = 0,
          v32 = 1,
          v64 = 2
       }
"
Add-Type -TypeDefinition $enumDef -ErrorAction SilentlyContinue
} catch { }

try {
$enumDef = "
using System;
       [FlagsAttribute]
       public enum OfficeBranch
       {
          FirstReleaseCurrent = 0,
          Current = 1,
          FirstReleaseBusiness = 2,
          Business = 3,
          CMValidation = 4
       }
"
Add-Type -TypeDefinition $enumDef -ErrorAction SilentlyContinue
} catch { }

try {
$enumDef = "
using System;
       [FlagsAttribute]
       public enum OfficeChannel
       {
          FirstReleaseCurrent = 0,
          Current = 1,
          FirstReleaseDeferred = 2,
          Deferred = 3
       }
"
Add-Type -TypeDefinition $enumDef -ErrorAction SilentlyContinue
} catch { }

try {
$enum = "
using System;
 
    [FlagsAttribute]
    public enum GPODeploymentType
    {
        DeployWithScript = 0,
        DeployWithConfigurationFile = 1,
        DeployWithInstallationFile = 2
    }
"
Add-Type -TypeDefinition $enum -ErrorAction SilentlyContinue
} catch { }

function Download-GPOOfficeChannelFiles() {
<#
.SYNOPSIS
Downloads the Office Click-to-Run files into the specified folder.

.DESCRIPTION
Downloads the Office 365 ProPlus installation files to a specified file path.

.PARAMETER Channels
The update channel. Current, Deferred, FirstReleaseDeferred, FirstReleaseCurrent

.PARAMETER OfficeFilesPath
This is the location where the source files will be downloaded.

.PARAMETER Languages
All office languages are supported in the ll-cc format "en-us"

.PARAMETER Bitness
Downloads the bitness of Office Click-to-Run "v32, v64, Both"

.PARAMETER Version
You can specify the version to download. 16.0.6868.2062. Version information can be found here https://technet.microsoft.com/en-us/library/mt592918.aspx

.EXAMPLE
Download-GPOOfficeChannelFiles -OfficeFilesPath D:\OfficeChannelFiles

.EXAMPLE
Download-GPOOfficeChannelFiles -OfficeFilesPath D:\OfficeChannelFiles -Channels Deferred -Bitness v32

.EXAMPLE
Download-GPOOfficeChannelFiles -OfficeFilesPath D:\OfficeChannelFiles -Bitness v32 -Channels Deferred,FirstReleaseDeferred -Languages en-us,es-es,ja-jp
#>

    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter()]
        [OfficeChannel[]] $Channels = @(1,2,3),

        [Parameter(Mandatory=$true)]
	    [String]$OfficeFilesPath = $NULL,

        [Parameter()]
        [ValidateSet("en-us","ar-sa","bg-bg","zh-cn","zh-tw","hr-hr","cs-cz","da-dk","nl-nl","et-ee","fi-fi","fr-fr","de-de","el-gr","he-il","hi-in","hu-hu","id-id","it-it",
                    "ja-jp","kk-kz","ko-kr","lv-lv","lt-lt","ms-my","nb-no","pl-pl","pt-br","pt-pt","ro-ro","ru-ru","sr-latn-rs","sk-sk","sl-si","es-es","sv-se","th-th",
                    "tr-tr","uk-ua","vi-vn")]
        [string[]] $Languages = ("en-us"),

        [Parameter()]
        [Bitness] $Bitness = 0,

        [Parameter()]
        [string] $Version = $NULL
        
    )

    Process {
       if (Test-Path "$PSScriptRoot\Download-OfficeProPlusChannels.ps1") {
         . "$PSScriptRoot\Download-OfficeProPlusChannels.ps1"
       } else {
       <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Dependency file missing: $PSScriptRoot\Download-OfficeProPlusChannels.ps1"
         throw "Dependency file missing: $PSScriptRoot\Download-OfficeProPlusChannels.ps1"
       }

       $ChannelList = @("FirstReleaseCurrent", "Current", "FirstReleaseDeferred", "Deferred")
       $ChannelXml = Get-ChannelXml -FolderPath $OfficeFilesPath -OverWrite $true

       foreach ($Channel in $ChannelList) {
         if ($Channels -contains $Channel) {

            $selectChannel = $ChannelXml.UpdateFiles.baseURL | Where {$_.branch -eq $Channel.ToString() }
            $latestVersion = Get-ChannelLatestVersion -ChannelUrl $selectChannel.URL -Channel $Channel
            $ChannelShortName = ConvertChannelNameToShortName -ChannelName $Channel

            if ($Version) {
               $latestVersion = $Version
            }

            Download-OfficeProPlusChannels -TargetDirectory $OfficeFilesPath  -Channels $Channel -Version $latestVersion -UseChannelFolderShortName $true -Languages $Languages -Bitness $Bitness

            $cabFilePath = "$env:TEMP/ofl.cab"
            Copy-Item -Path $cabFilePath -Destination "$OfficeFilesPath\ofl.cab" -Force

            $latestVersion = Get-ChannelLatestVersion -ChannelUrl $selectChannel.URL -Channel $Channel -FolderPath $OfficeFilesPath -OverWrite $true 
         }
       }
    }
}

Function Configure-GPOOfficeDeployment {
<#
.SYNOPSIS
Configures an Office deployment using Group Policy

.DESCRIPTION
Configures the folders and files to deploy Office using Group Policy

.PARAMETER Channel
The update channel to deploy.

.PARAMETER Bitness
The architecture of the update channel.

.PARAMETER OfficeSourceFilesPath
The path to the required deployment files.

.PARAMETER MoveSourceFiles
By default, the installation files will be moved to the source folder. Set this to $false to copy the installation files.

.EXAMPLE
Configure-GPOOfficeDeployment -Channel Current -Bitness 64 -OfficeSourceFilesPath D:\OfficeChannelFiles

.EXAMPLE
Configure-GPOOfficeDeployment -Channel Current -Bitness 64 -OfficeSourceFilesPath D:\OfficeChannelFiles -MoveSourceFiles $false
#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (      
        [Parameter()]
        [OfficeChannel]$Channel,

        [Parameter()]
        [Bitness]$Bitness = "v32",

        [Parameter()]
	    [string]$OfficeSourceFilesPath,

        [Parameter()]
        [string]$MoveSourceFiles = $true
    )

    Begin
    {
	    $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location        
    }

    Process 
    {
        Try{
            $cabFilePath = "$OfficeSourceFilesPath\ofl.cab"
            if(Test-Path $cabFilePath){
                Copy-Item -Path $cabFilePath -Destination "$PSScriptRoot\ofl.cab" -Force
            }

            $ChannelXml = Get-ChannelXml -FolderPath $OfficeSourceFilesPath -OverWrite $false
           
            $selectChannel = $ChannelXml.UpdateFiles.baseURL | Where {$_.branch -eq $Channel.ToString() }
            $latestVersion = Get-ChannelLatestVersion -ChannelUrl $selectChannel.URL -Channel $Channel -FolderPath $OfficeFilesPath -OverWrite $false
        
            $ChannelShortName = ConvertChannelNameToShortName -ChannelName $Channel
            $LargeDrv = Get-LargestDrive
        
            $Path = CreateOfficeChannelShare -Path "$LargeDrv\OfficeDeployment"
        
            $ChannelPath = "$Path\$Channel"
            $LocalPath = "$LargeDrv\OfficeDeployment"
            $LocalChannelPath = "$LargeDrv\OfficeDeployment\SourceFiles"
        
            [System.IO.Directory]::CreateDirectory($LocalChannelPath) | Out-Null
                   
            if($OfficeSourceFilesPath) {
                $officeFileChannelPath = "$OfficeSourceFilesPath\$ChannelShortName"
                $officeFileTargetPath = "$LocalChannelPath"

                [string]$oclVersion = $NULL
                if ($officeFileChannelPath) {
                    if (Test-Path -Path "$officeFileChannelPath\Office\Data") {
                       $oclVersion = Get-LatestVersion -UpdateURLPath $officeFileChannelPath
                    }
                }

                if ($oclVersion) {
                   $latestVersion = $oclVersion
                }

                if (!(Test-Path -Path $officeFileChannelPath)) {
                    <# write log#>
                    $lineNum = Get-CurrentLineNumber    
                    $filName = Get-CurrentFileName 
                    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Channel Folder Missing: $officeFileChannelPath - Ensure that you have downloaded the Channel you are trying to deploy"
                    throw "Channel Folder Missing: $officeFileChannelPath - Ensure that you have downloaded the Channel you are trying to deploy"
                }

                [System.IO.Directory]::CreateDirectory($officeFileTargetPath) | Out-Null

                if ($MoveSourceFiles) {
                    Move-Item -Path $officeFileChannelPath -Destination $officeFileTargetPath -Force
                } else {
                    Copy-Item -Path $officeFileChannelPath -Destination $officeFileTargetPath -Recurse -Force
                }

                $cabFilePath = "$OfficeSourceFilesPath\ofl.cab"
                if (Test-Path $cabFilePath) {
                    Copy-Item -Path $cabFilePath -Destination "$LocalPath\ofl.cab" -Force
                }
            } else {
                if(Test-Path -Path "$LocalChannelPath\Office") {
                    Remove-Item -Path "$LocalChannelPath\Office" -Force -Recurse
                }
            }
        
            $cabFilePath = "$env:TEMP/ofl.cab"
            if(!(Test-Path $cabFilePath)) {
                Copy-Item -Path "$LocalPath\ofl.cab" -Destination $cabFilePath -Force
            }

            CreateMainCabFiles -LocalPath $LocalPath -ChannelShortName $ChannelShortName -LatestVersion $latestVersion

            $DeploymentFilePath = "$PSSCriptRoot\DeploymentFiles\*.*"
            if (Test-Path -Path $DeploymentFilePath) {
                Copy-Item -Path $DeploymentFilePath -Destination "$LocalPath" -Force -Recurse
            } else {
                <# write log#>
                $lineNum = Get-CurrentLineNumber    
                $filName = Get-CurrentFileName 
                WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Deployment folder missing: $DeploymentFilePath"
                throw "Deployment folder missing: $DeploymentFilePath"
            }
        } Catch{}
    }        
}

Function Create-GPOOfficeDeployment {
<#
.SYNOPSIS
Configures an Office deployment using Group Policy

.DESCRIPTION
Configures an existing Group Policy Object to deploy Office 365 ProPlus 

.PARAMETER GroupPolicyName
The name of the Group Policy Object

.PARAMETER DeploymentType
Choose between DeployWithScript or DeployWithConfigurationFile. DeployWithScript will deploy a dynamic installation
using the target computer's existing Office installation. DeployWithConfigurationFile will deploy a standard Office installation
to all of the targeted computers.

.PARAMETER ScriptName
The name of the deployment script if the DeploymentType is DeployWithScript. If ScriptName is not specified the 
GPO-OfficeDeploymentScript.ps1 will be used.

.PARAMETER OfficeDeploymentFileName
The name of an Office installation file to deploy. An Office install MSI or EXE can be generated using the
Microsoft Office ProPlus Install Toolkit which can be downloaded from http://officedev.github.io/Office-IT-Pro-Deployment-Scripts/XmlEditor.html

.PARAMETER Channel
The update channel to install.

.PARAMETER Bitness
The update channel bit to install.

.PARAMETER ConfigurationXML
The name of a custom (ODT) configuration.xml file if DeploymentTYpe is set to DeployWithConfigurationFile. If you plan on using a custom xml
for the deployment make sure to copy the file to the DeploymentFiles folder before running Configure-GPOOfficeDeployment, or copy the file
to OfficeDeployment if Configure-GPOOfficeDeployment has already been ran.

.PARAMETER WaitForInstallToFinish
While Office is installing PowerShell will remain open until the installation is finished.

.PARAMETER InstallProofingTools
Set this value to $true to include the Proofing Tools exe with the deployment.

.EXAMPLE 
Create-GPOOfficeDeployment -GroupPolicyName DeployCurrentChannel64Bit -DeploymentType DeployWithScript -Channel Current -Bitness 64

.EXAMPLE
Create-GPOOfficeDeployment -GroupPolicyName DeployDeferredChannel32Bit -DeploymentType DeployWithConfigurationFile -Channel Current -Bitness 64 -ConfigurationXML Config-Deferred-32bit.xml

.EXAMPLE
Create-GPOOfficeDeployment -GroupPolicyName DeployWithMSI -DeploymentType DeployWithInstallationFile -OfficeDeploymentFileName OfficeProPlus.msi
#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
	    [Parameter(Mandatory=$True)]
	    [string]$GroupPolicyName,
	
        [Parameter()]
	    [GPODeploymentType]$DeploymentType = 0,
        
        [Parameter()]
        [string]$ScriptName,
              
        [Parameter()]
        [OfficeChannel]$Channel,

        [Parameter()]
        [Bitness]$Bitness,

        [Parameter()]
        [string]$ConfigurationXML = $null,

        [Parameter()]
        [string]$OfficeDeploymentFileName,

        [Parameter()]
        [bool]$WaitForInstallToFinish = $true,

        [Parameter()]
        [bool]$InstallProofingTools = $false,

        [Parameter()]
        [bool]$Quiet = $true
    )

    Begin
    {
	    $currentExecutionPolicy = Get-ExecutionPolicy
	    Set-ExecutionPolicy Unrestricted -Scope Process -Force  
        $startLocation = Get-Location        
    }

    Process
    {
        $Root = [ADSI]"LDAP://RootDSE"
        $DomainPath = $Root.Get("DefaultNamingContext")

        Write-Host "Configuring Group Policy to Install Office Click-To-Run"
        Write-Host
        <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Configuring Group Policy to Install Office Click-To-Run"

        Write-Host "Searching for GPO: $GroupPolicyName..." -NoNewline
        <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Searching for GPO: $GroupPolicyName..."
	    $gpo = Get-GPO -Name $GroupPolicyName
	
	    if(!$gpo -or ($gpo -eq $null))
	    {
            <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "The GPO $GroupPolicyName could not be found."
		    Write-Error "The GPO $GroupPolicyName could not be found."
	    }

        Write-Host "GPO Found"
        <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "GPO Found"

        Write-Host "Modifying GPO: $GroupPolicyName..." -NoNewline
        <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Modifying GPO: $GroupPolicyName..."

	    $baseSysVolPath = "$env:LOGONSERVER\sysvol"

	    $domain = $gpo.DomainName
        $gpoId = $gpo.Id.ToString()

        $adGPO = [ADSI]"LDAP://CN={$gpoId},CN=Policies,CN=System,$DomainPath"
    	
	    $gpoPath = "{0}\{1}\Policies\{{{2}}}" -f $baseSysVolPath, $domain, $gpoId
	    $relativePathToScriptsFolder = "Machine\Scripts"
	    $scriptsPath = "{0}\{1}" -f $gpoPath, $relativePathToScriptsFolder

        $createDir = [system.io.directory]::CreateDirectory($scriptsPath) 

	    $gptIniFileName = "GPT.ini"
	    $gptIniFilePath = ".\$gptIniFileName"
   
	    Set-Location $scriptsPath
	
	    #region PSSCripts.ini
	    $psScriptsFileName = "psscripts.ini"
        $scriptsFileName = "scripts.ini"

	    $psScriptsFilePath = ".\$psScriptsFileName"
        $scriptsFilePath = ".\$scriptsFileName"

	    $encoding = 'Unicode' #[System.Text.Encoding]::Unicode

	    if(!(Test-Path $psScriptsFilePath))
	    {
		    $baseContent = "`r`n[ScriptsConfig]`r`nStartExecutePSFirst=true`r`n[Startup]"
		    $baseContent | Out-File -FilePath $psScriptsFilePath -Encoding unicode -Force
		
		    $file = Get-ChildItem -Path $psScriptsFilePath
		    $file.Attributes = $file.Attributes -bor ([System.IO.FileAttributes]::Hidden).value__
	    }

	    if(!(Test-Path $scriptsFilePath))
	    {
            "" | Out-File -FilePath $scriptsFilePath -Encoding unicode -Force

		    $file = Get-ChildItem -Path $scriptsFilePath
		    $file.Attributes = $file.Attributes -bor ([System.IO.FileAttributes]::Hidden).value__
        }
	
	    $content = Get-Content -Encoding $encoding -Path $psScriptsFilePath

	    $length = $content.Length

	    $newContentLength = $length + 2

	    $newContent = New-Object System.String[] ($newContentLength)

	    $pattern = [string]"\[\w+\]"

	    $startUpIndex = 0
	    $nextIndex = 0
	    $startUpFound = $false

	    foreach($s in $content)
	    {
		    if($s -match $pattern)
		    {
		       if($startUpFound)
		       {
			      $nextIndex = $content.IndexOf($s) - 1
			      break
		       }
		       else
		       {
				    if($s -eq "[Startup]")
				    {
					    $startUpIndex = $content.IndexOf($s)
					    $startUpFound = $true
				    }
		       }
		    }
	    }

	    if($startUpFound -and ($nextIndex -eq 0))
	    {
		    $nextIndex = $content.Count - 1;
	    }
	
	    $lastEntry = [string]$content[$nextIndex]

	    $num = [regex]::Matches($lastEntry, "\d+")[0].Value   
	
	    if($num)
	    {
		    $lastScriptIndex = [Convert]::ToInt32($num)
	    }
	    else
	    {
		    $lastScriptIndex = 0
		    $nextScriptIndex = 0
	    }
	
	    if($lastScriptIndex -gt 0)
	    {
		    $nextScriptIndex = $lastScriptIndex + 1
	    }

	    for($i=0; $i -le $nextIndex; $i++)
	    {
		    $newContent[$i] = $content[$i]
	    }
                      
        $ChannelShortName = ConvertChannelNameToShortName -ChannelName $Channel
        $LargeDrv = Get-LargestDrive        
        $OfficeDeploymentLocalPath = "$LargeDrv\OfficeDeployment"
        $OfficeDeploymentShare = Get-WmiObject Win32_Share | ? {$_.Name -like "OfficeDeployment$"}
        $OfficeDeploymentName = $OfficeDeploymentShare.Name
        $OfficeDeploymentUNC = "\\" + $OfficeDeploymentShare.PSComputerName + "\$OfficeDeploymentName" 
        
        if($Bitness -like "v64"){
            $Bit = "64"
        } else {
            $Bit = "32"
        } 
               
        if($DeploymentType -eq "DeployWithConfigurationFile")
        {
            if(!$ScriptName){$ScriptName = "DeployConfigFile.ps1"}

            $newContent[$nextIndex+1] = "{0}CmdLine={1}" -f $nextScriptIndex, $ScriptName

            if($WaitForInstallToFinish -eq $false){
	            $newContent[$nextIndex+2] = "{0}Parameters=-OfficeDeploymentPath {1} -ConfigFileName {2} -WaitForInstallToFinish {3} -Channel {4} -Bitness {5}" -f $nextScriptIndex, $OfficeDeploymentUNC, $ConfigurationXML, $WaitForInstallToFinish, $Channel, $Bit
                if($InstallProofingTools -eq $true){
                    $newContent[$nextIndex+2] = "{0}Parameters=-OfficeDeploymentPath {1} ConfigFileName {2} -WaitForInstallToFinish {3} -InstallProofingTools {4} -Channel {5} -Bitness {6}" -f $nextScriptIndex, $OfficeDeploymentUNC, $ConfigurationXML, $WaitForInstallToFinish, $InstallProofingTools, $Channel, $Bit
                }
            } else {
                if($InstallProofingTools -eq $true){
                    $newContent[$nextIndex+2] = "{0}Parameters=-OfficeDeploymentPath {1} -ConfigFileName {2} -InstallProofingTools {3} -Channel {4} -Bitness {5}" -f $nextScriptIndex, $OfficeDeploymentUNC, $ConfigurationXML, $InstallProofingTools, $Channel, $Bit
                } else {
                    $newContent[$nextIndex+2] = "{0}Parameters=-OfficeDeploymentPath {1} -ConfigFileName {2} -Channel {3} -Bitness {4}" -f $nextScriptIndex, $OfficeDeploymentUNC, $ConfigurationXML, $Channel, $Bit
                }
            }
        } elseif ($DeploymentType -eq "DeployWithScript") 
        {
            if(!$ScriptName){$ScriptName = "GPO-OfficeDeploymentScript.ps1"}

            $newContent[$nextIndex+1] = "{0}CmdLine={1}" -f $nextScriptIndex, $ScriptName

            if($Channel -eq $null -and $Bitness -eq $null){
                $newContent[$nextIndex+2] = "{0}Parameters=-OfficeDeploymentPath {1}" -f $nextScriptIndex, $OfficeDeploymentUNC
            }
            elseif($Channel -eq $null){
                $newContent[$nextIndex+2] = "{0}Parameters=-OfficeDeploymentPath {1} -Bitness {2}" -f $nextScriptIndex, $OfficeDeploymentUNC, $Bit
            }
            elseif($Bitness -eq $null){
                $newContent[$nextIndex+2] = "{0}Parameters=-OfficeDeploymentPath {1} -Channel {2}" -f $nextScriptIndex, $OfficeDeploymentUNC, $Channel
            } else {
                $newContent[$nextIndex+2] = "{0}Parameters=-OfficeDeploymentPath {1} -Channel {2} -Bitness {3}" -f $nextScriptIndex, $OfficeDeploymentUNC, $Channel, $Bit
            }

        } elseif($DeploymentType -eq "DeployWithInstallationFile")
        {
            if(!$ScriptName){$ScriptName = "DeployOfficeInstallationFile.ps1"}
            if(!$OfficeDeploymentFileName){$OfficeDeploymentFileName = "OfficeProPlus.msi"}
            
            $Quiet = Convert-Bool $Quiet
            
            $newContent[$nextIndex+1] = "{0}CmdLine={1}" -f $nextScriptIndex, $ScriptName
            $newContent[$nextIndex+2] = "{0}Parameters=-OfficeDeploymentPath {1} -OfficeDeploymentFileName {2} -Quiet {3}" -f $nextScriptIndex, $OfficeDeploymentUNC, $OfficeDeploymentFileName, $Quiet

        }

	    for($i=$nextIndex; $i -lt $length; $i++)
	    {
		    $newContent[$i] = $content[$i]
	    }

	    $newContent | Set-Content -Encoding $encoding -Path $psScriptsFilePath -Force
	    #endregion
	
	    #region Place the script to attach in the StartUp Folder
        $LargeDrv = Get-LargestDrive 
	    $setupExeSourcePath = "$LargeDrv\OfficeDeployment\$ScriptName"
	    $setupExeTargetPath = "$scriptsPath\StartUp"
        $setupExeTargetPathShutdown = "$scriptsPath\ShutDown"

        $createDir = [system.io.directory]::CreateDirectory($setupExeTargetPath) 
        $createDir = [system.io.directory]::CreateDirectory($setupExeTargetPathShutdown) 
	
	    Copy-Item -Path $setupExeSourcePath -Destination $setupExeTargetPath -Force
	    #endregion
	
	    #region Update GPT.ini
	    Set-Location $gpoPath   

	    $encoding = 'ASCII' #[System.Text.Encoding]::ASCII
	    $gptIniContent = Get-Content -Encoding $encoding -Path $gptIniFilePath
	
        [int]$newVersion = 0
	    foreach($s in $gptIniContent)
	    {
		    if($s.StartsWith("Version"))
		    {
			    $index = $gptIniContent.IndexOf($s)

			    #Write-Host "Old GPT.ini Version: $s"

			    $num = ($s -split "=")[1]

			    $ver = [Convert]::ToInt32($num)

			    $newVer = $ver + 1

			    $s = $s -replace $num, $newVer.ToString()

			    #Write-Host "New GPT.ini Version: $s"

                $newVersion = $s.Split('=')[1]

			    $gptIniContent[$index] = $s
			    break
		    }
	    }

        [System.Collections.ArrayList]$extList = New-Object System.Collections.ArrayList

        Try {
           $currentExt = $adGPO.get('gPCMachineExtensionNames')
        } Catch { 

        }

        if ($currentExt) {
            $extSplit = $currentExt.Split(']')

            foreach ($extGuid in $extSplit) {
              if ($extGuid) {
                if ($extGuid.Length -gt 0) {
                    $addItem = $extList.Add($extGuid.Replace("[", "").ToUpper())
                }
              }
            }
        }

        $extGuids = @("{42B5FAAE-6536-11D2-AE5A-0000F87571E3}{40B6664F-4972-11D1-A7CA-0000F87571E3}")

        foreach ($extGuid in $extGuids) {
            if (!$extList.Contains($extGuid)) {
              $addItem = $extList.Add($extGuid)
            }
        }

        foreach ($extAddGuid in $extList) {
           $newGptExt += "[$extAddGuid]"
        }

        $adGPO.put('versionNumber',$newVersion)
        $adGPO.put('gPCMachineExtensionNames',$newGptExt)
        $adGPO.CommitChanges()
    
	    $gptIniContent | Set-Content -Encoding $encoding -Path $gptIniFilePath -Force
	
        Write-Host "GPO Modified"
        Write-Host ""
        Write-Host "The Group Policy '$GroupPolicyName' has been modified to install Office at Workstation Startup." -BackgroundColor DarkBlue
        Write-Host "Once Group Policy has refreshed on the Workstations then Office will install on next startup if the computer has access to the Network Share." -BackgroundColor DarkBlue
        <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "GPO Modified"
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "The Group Policy '$GroupPolicyName' has been modified to install Office at Workstation Startup."
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Once Group Policy has refreshed on the Workstations then Office will install on next startup if the computer has access to the Network Share."

    }

    End 
    {      
       $setLocation = Set-Location $startLocation
    }
}

function DownloadFile($url, $targetFile) {

  for($t=1;$t -lt 10; $t++) {
   try {
       $uri = New-Object "System.Uri" "$url"
       $request = [System.Net.HttpWebRequest]::Create($uri)
       $request.set_Timeout(15000) #15 second timeout

       $response = $request.GetResponse()
       $totalLength = [System.Math]::Floor($response.get_ContentLength()/1024)
       $responseStream = $response.GetResponseStream()
       $targetStream = New-Object -TypeName System.IO.FileStream -ArgumentList $targetFile.replace('/','\'), Create
       $buffer = new-object byte[] 8192KB
       $count = $responseStream.Read($buffer,0,$buffer.length)
       $downloadedBytes = $count

       while ($count -gt 0)
       {
           $targetStream.Write($buffer, 0, $count)
           $count = $responseStream.Read($buffer,0,$buffer.length)
           $downloadedBytes = $downloadedBytes + $count
           Write-Progress -id 3 -ParentId 2 -activity "Downloading file '$($url.split('/') | Select -Last 1)'" -status "Downloaded ($([System.Math]::Floor($downloadedBytes/1024))K of $($totalLength)K): " -PercentComplete ((([System.Math]::Floor($downloadedBytes/1024)) / $totalLength)  * 100)
       }

       Write-Progress -id 3 -ParentId 2 -activity "Finished downloading file '$($url.split('/') | Select -Last 1)'"

       $targetStream.Flush()
       $targetStream.Close()
       $targetStream.Dispose()
       $responseStream.Dispose()
       break;
   } catch {
     $strError = $_.Message
     if ($t -ge 9) {
        throw
     }
   }
   Start-Sleep -Milliseconds 500
  }
}

function PurgeOlderVersions([string]$targetDirectory, [int]$numVersionsToKeep, [array]$channels){
    Write-Host "Checking for Older Versions"
    <# write log#>
    $lineNum = Get-CurrentLineNumber    
    $filName = Get-CurrentFileName 
    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Checking for Older Versions"
                         
    for($k = 0; $k -lt $channels.Count; $k++)
    {
        [array]$totalVersions = @()#declare empty array so each folder can be purged of older versions individually
        [string]$channelName = $channels[$k]
        [string]$shortChannelName = ConvertChannelNameToShortName -ChannelName $channelName
        [string]$branchName = ConvertChannelNameToBranchName -ChannelName $channelName
        [string]$channelName2 = ConvertBranchNameToChannelName -BranchName $channelName

        $folderList = @($channelName, $shortChannelName, $channelName2, $branchName)

        foreach ($folderName in $folderList) {
            $directoryPath = $TargetDirectory.ToString() + '\'+ $folderName +'\Office\Data'

            if (Test-Path -Path $directoryPath) {
               break;
            }
        }

        if (Test-Path -Path $directoryPath) {
            Write-Host "`tChannel: $channelName2"
            <# write log#>
            $lineNum = Get-CurrentLineNumber    
            $filName = Get-CurrentFileName 
            WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Channel: $channelName2"
             [bool]$versionsToRemove = $false

            $files = Get-ChildItem $directoryPath  
            Foreach($file in $files)
            {        
                if($file.GetType().Name -eq 'DirectoryInfo')
                {
                    $totalVersions+=$file.Name
                }
            }

            #check if number of versions is greater than number of versions to hold onto, if not, then we don't need to do anything
            if($totalVersions.Length -gt $numVersionsToKeep)
            {
                #sort array in numerical order
                $totalVersions = $totalVersions | Sort-Object 
               
                #delete older versions
                $numToDelete = $totalVersions.Length - $numVersionsToKeep
                for($i = 1; $i -le $numToDelete; $i++)#loop through versions
                {
                     $versionsToRemove = $true
                     $removeVersion = $totalVersions[($i-1)]
                     Write-Host "`t`tRemoving Version: $removeVersion"
                     <# write log#>
                    $lineNum = Get-CurrentLineNumber    
                    $filName = Get-CurrentFileName 
                    WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Removing Version: $removeVersion"
                     
                     Foreach($file in $files)#loop through files
                     {  #array is 0 based

                        if($file.Name.Contains($removeVersion))
                        {                               
                            $folderPath = "$directoryPath\$file"

                             for($t=1;$t -lt 5; $t++) {
                               try {
                                  Remove-Item -Recurse -Force $folderPath
                                  break;
                               } catch {
                                 if ($t -ge 4) {
                                    throw
                                 }
                               }
                             }
                        }
                     }
                }

            }

            if (!($versionsToRemove)) {
                Write-Host "`t`tNo Versions to Remove"
                 <# write log#>
                $lineNum = Get-CurrentLineNumber    
                $filName = Get-CurrentFileName 
                WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "No Versions to Remove"
            }
        }


    }    
      
}

function ConvertChannelNameToShortName {
    Param(
       [Parameter()]
       [string] $ChannelName
    )
    Process {
       if ($ChannelName.ToLower() -eq "FirstReleaseCurrent".ToLower()) {
         return "FRCC"
       }
       if ($ChannelName.ToLower() -eq "Current".ToLower()) {
         return "CC"
       }
       if ($ChannelName.ToLower() -eq "FirstReleaseDeferred".ToLower()) {
         return "FRDC"
       }
       if ($ChannelName.ToLower() -eq "Deferred".ToLower()) {
         return "DC"
       }
       if ($ChannelName.ToLower() -eq "Business".ToLower()) {
         return "DC"
       }
       if ($ChannelName.ToLower() -eq "FirstReleaseBusiness".ToLower()) {
         return "FRDC"
       }
    }
}

function ConvertChannelNameToBranchName {
    Param(
       [Parameter()]
       [string] $ChannelName
    )
    Process {
       if ($ChannelName.ToLower() -eq "FirstReleaseCurrent".ToLower()) {
         return "FirstReleaseCurrent"
       }
       if ($ChannelName.ToLower() -eq "Current".ToLower()) {
         return "Current"
       }
       if ($ChannelName.ToLower() -eq "FirstReleaseDeferred".ToLower()) {
         return "FirstReleaseBusiness"
       }
       if ($ChannelName.ToLower() -eq "Deferred".ToLower()) {
         return "Business"
       }
       if ($ChannelName.ToLower() -eq "Business".ToLower()) {
         return "Business"
       }
       if ($ChannelName.ToLower() -eq "FirstReleaseBusiness".ToLower()) {
         return "FirstReleaseBusiness"
       }
    }
}

function ConvertBranchNameToChannelName {
    Param(
       [Parameter()]
       [string] $BranchName
    )
    Process {
       if ($BranchName.ToLower() -eq "FirstReleaseCurrent".ToLower()) {
         return "FirstReleaseCurrent"
       }
       if ($BranchName.ToLower() -eq "Current".ToLower()) {
         return "Current"
       }
       if ($BranchName.ToLower() -eq "FirstReleaseDeferred".ToLower()) {
         return "FirstReleaseDeferred"
       }
       if ($BranchName.ToLower() -eq "Deferred".ToLower()) {
         return "Deferred"
       }
       if ($BranchName.ToLower() -eq "Business".ToLower()) {
         return "Deferred"
       }
       if ($BranchName.ToLower() -eq "FirstReleaseBusiness".ToLower()) {
         return "FirstReleaseDeferred"
       }
    }
}


function UpdateConfigurationXml() {
    [CmdletBinding()]	
    Param
	(
		[Parameter(Mandatory=$true)]
		[String]$Path = "",

		[Parameter(Mandatory=$true)]
		[String]$Channel = "",

		[Parameter(Mandatory=$true)]
		[String]$Bitness,

		[Parameter()]
		[String]$SourcePath = $NULL,
		
        [Parameter()]
		[String]$Language
        
	) 
    Process {
	  $doc = [Xml] (Get-Content $Path)

      $addNode = $doc.Configuration.Add
      $languageNode = $addNode.Product.Language

      if ($addNode.OfficeClientEdition) {
          $addNode.OfficeClientEdition = $Bitness
      } else {
          $addNode.SetAttribute("OfficeClientEdition", $Bitness)
      }

      if ($addNode.Channel) {
          $addNode.Channel = $Channel
      } else {
          $addNode.SetAttribute("Channel", $Channel)
      }

      if ($addNode.SourcePath) {
          $addNode.SourcePath = $SourcePath
      } else {
          $addNode.SetAttribute("SourcePath", $SourcePath)
      }

      if($Language){
          if ($languageNode.ID){
              if($languageNode.ID -contains $Language) {
                  Write-Host "$Language already exists in the xml"
                  <# write log#>
                $lineNum = Get-CurrentLineNumber    
                $filName = Get-CurrentFileName 
                WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "$Language already exists in the xml"
              } else {
                  $newLanguageElement = $doc.CreateElement("Language")
                  $newLanguage = $doc.Configuration.Add.Product.AppendChild($newLanguageElement)
                  $newLanguage.SetAttribute("ID", $Language)
              }
          } else {
              $languageNode.SetAttribute("ID", $language)
          }
     }

      $doc.Save($Path)
    }
}

function CreateMainCabFiles() {
    [CmdletBinding()]	
    Param
	(
		[Parameter(Mandatory=$true)]
		[String]$LocalPath = "",

        [Parameter(Mandatory=$true)]
        [String] $ChannelShortName,

        [Parameter(Mandatory=$true)]
        [String] $LatestVersion
	) 
    Process {
        $versionFile321 = "$LocalPath\$ChannelShortName\Office\Data\v32_$LatestVersion.cab"
        $v32File1 = "$LocalPath\$ChannelShortName\Office\Data\v32.cab"

        $versionFile641 = "$LocalPath\$ChannelShortName\Office\Data\v64_$LatestVersion.cab"
        $v64File1 = "$LocalPath\$ChannelShortName\Office\Data\v64.cab"

        $versionFile322 = "$LocalPath\SourceFiles\$ChannelShortName\Office\Data\v32_$LatestVersion.cab"
        $v32File2 = "$LocalPath\SourceFiles\$ChannelShortName\Office\Data\v32.cab"

        $versionFile642 = "$LocalPath\SourceFiles\$ChannelShortName\Office\Data\v64_$LatestVersion.cab"
        $v64File2 = "$LocalPath\SourceFiles\$ChannelShortName\Office\Data\v64.cab"

        if (Test-Path -Path $versionFile321) {
            Copy-Item -Path $versionFile321 -Destination $v32File1 -Force
        }

        if (Test-Path -Path $versionFile641) {
            Copy-Item -Path $versionFile641 -Destination $v64File1 -Force
        }

        if (Test-Path -Path $versionFile322) {
            Copy-Item -Path $versionFile322 -Destination $v32File2 -Force
        }

        if (Test-Path -Path $versionFile642) {
            Copy-Item -Path $versionFile642 -Destination $v64File2 -Force
        }
    }
}

function CheckIfVersionExists() {
    [CmdletBinding()]	
    Param
	(
	   [Parameter(Mandatory=$True)]
	   [String]$Version,

		[Parameter()]
		[String]$Channel
    )
    Begin
    {
        $startLocation = Get-Location
    }
    Process {
       LoadCMPrereqs

       $VersionName = "$Channel - $Version"

       $packageName = "Office 365 ProPlus"

       $existingPackage = Get-CMPackage | Where { $_.Name -eq $packageName -and $_.Version -eq $Version }
       if ($existingPackage) {
         return $true
       }

       return $false
    }
}

function CreateOfficeChannelShare() {
    [CmdletBinding()]	
    Param
	(
        [Parameter()]
        [String]$Name = "OfficeDeployment$",

        [Parameter()]
        [String]$Path = "$env:SystemDrive\OfficeDeployment"
	) 
    
    if (!(Test-Path $Path)) { 
      $addFolder = New-Item $Path -type Directory 
    }
    
    $ACL = Get-ACL $Path

    $identity = New-Object System.Security.Principal.NTAccount  -argumentlist ("$env:UserDomain\$env:UserName") 
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule -argumentlist ($identity,"FullControl","ContainerInherit, ObjectInherit","None","Allow")

    $addAcl = $ACL.AddAccessRule($accessRule) | Out-Null

    $identity = New-Object System.Security.Principal.NTAccount -argumentlist ("$env:UserDomain\Domain Admins") 
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule -argumentlist ($identity,"FullControl","ContainerInherit, ObjectInherit","None","Allow")
    $addAcl = $ACL.AddAccessRule($accessRule) | Out-Null

    $identity = "Everyone"
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule -argumentlist ($identity,"Read","ContainerInherit, ObjectInherit","None","Allow")
    $addAcl = $ACL.AddAccessRule($accessRule) | Out-Null

    Set-ACL -Path $Path -ACLObject $ACL | Out-Null
    
    $share = Get-WmiObject -Class Win32_share | Where {$_.name -eq "$Name"}
    if (!$share) {
       Create-FileShare -Name $Name -Path $Path | Out-Null
    }

    $sharePath = "\\$env:COMPUTERNAME\$Name"
    return $sharePath
}

function GetSupportedPlatforms([String[]] $requiredPlatformNames){
    $computerName = $env:COMPUTERNAME
    #$assignedSite = $([WmiClass]"\\$computerName\ROOT\ccm:SMS_Client").getassignedsite()
    $siteCode = Get-Site  
    $filteredPlatforms = Get-WmiObject -ComputerName $computerName -Class SMS_SupportedPlatforms -Namespace "root\sms\site_$siteCode" | Where-Object {$_.IsSupported -eq $true -and  $_.OSName -like 'Win NT' -and ($_.OSMinVersion -match "6\.[0-9]{1,2}\.[0-9]{1,4}\.[0-9]{1,4}" -or $_.OSMinVersion -match "10\.[0-9]{1,2}\.[0-9]{1,4}\.[0-9]{1,4}") -and ($_.OSPlatform -like 'I386' -or $_.OSPlatform -like 'x64')}

    $requiredPlatforms = $filteredPlatforms| Where-Object {$requiredPlatformNames.Contains($_.DisplayText) } #| Select DisplayText, OSMaxVersion, OSMinVersion, OSName, OSPlatform | Out-GridView

    $supportedPlatforms = @()

    foreach($p in $requiredPlatforms)
    {
        $osDetail = ([WmiClass]("\\$computerName\root\sms\site_$siteCode`:SMS_OS_Details")).CreateInstance()    
        $osDetail.MaxVersion = $p.OSMaxVersion
        $osDetail.MinVersion = $p.OSMinVersion
        $osDetail.Name = $p.OSName
        $osDetail.Platform = $p.OSPlatform

        $supportedPlatforms += $osDetail
    }

    $supportedPlatforms
}

function CreateDownloadXmlFile([string]$Path, [string]$ConfigFileName){
	#1 - Set the correct version number to update Source location
	$sourceFilePath = "$path\$configFileName"
    $localSourceFilePath = ".\$configFileName"

    Set-Location $PSScriptRoot

    if (Test-Path -Path $localSourceFilePath) {   
	  $doc = [Xml] (Get-Content $localSourceFilePath)

      $addNode = $doc.Configuration.Add
	  $addNode.OfficeClientEdition = $bitness

      $doc.Save($sourceFilePath)
    } else {
      $doc = New-Object System.XML.XMLDocument

      $configuration = $doc.CreateElement("Configuration");
      $a = $doc.AppendChild($configuration);

      $addNode = $doc.CreateElement("Add");
      $addNode.SetAttribute("OfficeClientEdition", $bitness)
      if ($Version) {
         if ($Version.Length -gt 0) {
             $addNode.SetAttribute("Version", $Version)
         }
      }
      $a = $doc.DocumentElement.AppendChild($addNode);

      $addProduct = $doc.CreateElement("Product");
      $addProduct.SetAttribute("ID", "O365ProPlusRetail")
      $a = $addNode.AppendChild($addProduct);

      $addLanguage = $doc.CreateElement("Language");
      $addLanguage.SetAttribute("ID", "en-us")
      $a = $addProduct.AppendChild($addLanguage);

	  $doc.Save($sourceFilePath)
    }
}

function CreateUpdateXmlFile([string]$Path, [string]$ConfigFileName, [string]$Bitness, [string]$Version){
    $newConfigFileName = $ConfigFileName -replace '\.xml'
    $newConfigFileName = $newConfigFileName + "$Bitness" + ".xml"

    Copy-Item -Path ".\$ConfigFileName" -Destination ".\$newConfigFileName"
    $ConfigFileName = $newConfigFileName

    $testGroupFilePath = "$path\$ConfigFileName"
    $localtestGroupFilePath = ".\$ConfigFileName"

	$testGroupConfigContent = [Xml] (Get-Content $localtestGroupFilePath)

	$addNode = $testGroupConfigContent.Configuration.Add
	$addNode.OfficeClientEdition = $bitness
    $addNode.SourcePath = $path	

	$updatesNode = $testGroupConfigContent.Configuration.Updates
	$updatesNode.UpdatePath = $path
	$updatesNode.TargetVersion = $version

	$testGroupConfigContent.Save($testGroupFilePath)
    return $ConfigFileName
}

function DownloadBits() {
    [CmdletBinding()]	
    Param
	(
	    [Parameter()]
	    [OfficeBranch]$Branch = $null
	)

    $DownloadScript = "$PSScriptRoot\Download-OfficeProPlusBranch.ps1"
    if (Test-Path -Path $DownloadScript) {
       



    }
}

Function GetScriptRoot() {
 process {
     [string]$scriptPath = "."

     if ($PSScriptRoot) {
       $scriptPath = $PSScriptRoot
     } else {
       $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
     }

     return $scriptPath
 }
}

function GetQueryStatus(){
Param(
    [string]$SiteCode,
    [string]$PkgID
)

    $query = Get-WmiObject –NameSpace Root\SMS\Site_$SiteCode –Class SMS_DistributionDPStatus –Filter "PackageID='$PkgID'" | Select Name, MessageID, MessageState, LastUpdateDate

    if ($query -eq $null)
    {  
    <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "PackageID not found"
        throw "PackageID not found"
    }

    foreach ($objItem in $query){

        $DPName = $objItem.Name
        $UpdDate = [System.Management.ManagementDateTimeconverter]::ToDateTime($objItem.LastUpdateDate)

        switch ($objItem.MessageState)
        {
            1         {$Status = "Success"}
            2         {$Status = "In Progress"}
            3         {$Status = "Failed"}
            4         {$Status = "Error"}
        }

        switch ($objItem.MessageID)
        {
            2300      {$Message = "Content is beginning to process"}
            2301      {$Message = "Content has been processed successfully"}
            2303      {$Message = "Failed to process package"}
            2311      {$Message = "Distribution Manager has successfully created or updated the package"}
            2303      {$Message = "Content was successfully refreshed"}
            2323      {$Message = "Failed to initialize NAL"}
            2324      {$Message = "Failed to access or create the content share"}
            2330      {$Message = "Content was distributed to distribution point"}
            2342      {$Message = "Content is beginning to distribute"}
            2354      {$Message = "Failed to validate content status file"}
            2357      {$Message = "Content transfer manager was instructed to send content to Distribution Point"}
            2360      {$Message = "Status message 2360 unknown"}
            2370      {$Message = "Failed to install distribution point"}
            2371      {$Message = "Waiting for prestaged content"}
            2372      {$Message = "Waiting for content"}
            2376      {$Message = "Distribution Manager created a snapshot for content"}
            2380      {$Message = "Content evaluation has started"}
            2381      {$Message = "An evaluation task is running. Content was added to Queue"}
            2382      {$Message = "Content hash is invalid"}
            2383      {$Message = "Failed to validate content hash"}
            2384      {$Message = "Content hash has been successfully verified"}
            2391      {$Message = "Failed to connect to remote distribution point"}
            2397      {$Message = "Detail will be available after the server finishes processing the messages"}
            2398      {$Message = "Content Status not found"}
            8203      {$Message = "Failed to update package"}
            8204      {$Message = "Content is being distributed to the distribution Point"}
            8211      {$Message = "Failed to update package"}
        }

        $Displayvalue = showTaskStatus -Operation $Status -Status $Message -DateTime $UpdDate

    }

    return $Displayvalue
}

function showTaskStatus() {
    [CmdletBinding()]
    Param(
        [Parameter()]
        [string] $Operation = "",

        [Parameter()]
        [string] $Status = "",

        [Parameter()]
        [string] $DateTime = ""
    )

    $Result = New-Object –TypeName PSObject 
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Operation" -Value $Operation
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "Status" -Value $Status
    Add-Member -InputObject $Result -MemberType NoteProperty -Name "DateTime" -Value $DateTime
    return $Result
}

function Get-CurrentLineNumber {
    $MyInvocation.ScriptLineNumber
}

function Get-CurrentFileName{
    $MyInvocation.ScriptName.Substring($MyInvocation.ScriptName.LastIndexOf("\")+1)
}

function Get-CurrentFunctionName {
    (Get-Variable MyInvocation -Scope 1).Value.MyCommand.Name;
}

Function WriteToLogFile() {
    param( 
      [Parameter(Mandatory=$true)]
      [string]$LNumber,
      [Parameter(Mandatory=$true)]
      [string]$FName,
      [Parameter(Mandatory=$true)]
      [string]$ActionError
    )
    try{
        $headerString = "Time".PadRight(30, ' ') + "Line Number".PadRight(15,' ') + "FileName".PadRight(60,' ') + "Action"
        $stringToWrite = $(Get-Date -Format G).PadRight(30, ' ') + $($LNumber).PadRight(15, ' ') + $($FName).PadRight(60,' ') + $ActionError

        #check if file exists, create if it doesn't
        $getCurrentDatePath = "C:\Windows\Temp\" + (Get-Date -Format u).Substring(0,10)+"OfficeAutoScriptLog.txt"
        if(Test-Path $getCurrentDatePath){#if exists, append  
             Add-Content $getCurrentDatePath $stringToWrite
        }
        else{#if not exists, create new
             Add-Content $getCurrentDatePath $headerString
             Add-Content $getCurrentDatePath $stringToWrite
        }
    } catch [Exception]{
        Write-Host $_
        $fileName = $_.InvocationInfo.ScriptName.Substring($_.InvocationInfo.ScriptName.LastIndexOf("\")+1)
        WriteToLogFile -LNumber $_.InvocationInfo.ScriptLineNumber -FName $fileName -ActionError $_
    }
}

Function Convert-Bool() {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        [Parameter(Mandatory=$true)]
        [bool] $value
    )

    $newValue = "$" + $value.ToString()
    return $newValue 
}

$scriptPath = GetScriptRoot

$shareFunctionsPath = "$scriptPath\SharedFunctions.ps1"
if ($scriptPath.StartsWith("\\")) {
} else {
    if (!(Test-Path -Path $shareFunctionsPath)) {
    <# write log#>
        $lineNum = Get-CurrentLineNumber    
        $filName = Get-CurrentFileName 
        WriteToLogFile -LNumber $lineNum -FName $filName -ActionError "Missing Dependency File SharedFunctions.ps1"    
        throw "Missing Dependency File SharedFunctions.ps1"    
    }
}
. $shareFunctionsPath