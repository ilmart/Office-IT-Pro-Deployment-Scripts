param(
    [Parameter()]
    [string]$OfficeDeploymentPath,
    
	[Parameter(Mandatory=$true)]
	[String]$OfficeDeploymentFileName = $NULL,

    [Parameter()]
    [bool]$InstallSilently = $true
)

Set-Location $OfficeDeploymentPath

$DeploymentFile = "$OfficeDeploymentPath\$OfficeDeploymentFileName"

if($InstallSilently -eq $true){
    $arguments = " /i $DeploymentFile /qn"
} else {
    $arguments = " /i $DeploymentFile"
}

Start-Process $arguments