param(
    [Parameter()]
    [string]$OfficeDeploymentPath,
    
	[Parameter(Mandatory=$true)]
	[String]$OfficeDeploymentFileName = $NULL,

    [Parameter()]
    [bool]$Quiet = $true
)

$ActionFile = "$OfficeDeploymentPath\$OfficeDeploymentFileName"

if($OfficeDeploymentFileName.EndsWith("msi")){
    if($Quiet -eq $true){
        $argList = "/qn /norestart"
    } else {
        $argList = "/norestart"
    }

    $cmdLine = """$ActionFile"" $argList"
    $cmd = "cmd /c msiexec /i $cmdLine"
} elseif($OfficeDeploymentFileName.EndsWith("exe")){
    if($Quiet -eq $true){
        $argList = "/silent"
    }

    $cmd = "$ActionFile $argList"
}

Invoke-Expression $cmd