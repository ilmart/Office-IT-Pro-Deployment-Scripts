param(
    [Parameter()]
    [string]$OfficeDeploymentPath,
    
	[Parameter(Mandatory=$true)]
	[String]$OfficeDeploymentFileName = $NULL,

    [Parameter()]
    [string]$Quiet = "True"
)

$ActionFile = "$OfficeDeploymentPath\$OfficeDeploymentFileName"

if($OfficeDeploymentFileName.EndsWith("msi")){
    if($Quiet -eq "True"){
        $argList = "/qn /norestart"
    } else {
        $argList = "/norestart"
    }

    $cmdLine = """$ActionFile"" $argList"
    $cmd = "cmd /c msiexec /i $cmdLine"
} elseif($OfficeDeploymentFileName.EndsWith("exe")){
    if($Quiet -eq "True"){
        $argList = "/silent"
    }

    $cmd = "$ActionFile $argList"
}

Invoke-Expression $cmd