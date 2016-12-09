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
    $args = " /i $DeploymentFile /qn"
} else {
    $args = " /i $DeploymentFile"
}

[diagnostics.process]::Start("msiexec", $args).WaitForExit()