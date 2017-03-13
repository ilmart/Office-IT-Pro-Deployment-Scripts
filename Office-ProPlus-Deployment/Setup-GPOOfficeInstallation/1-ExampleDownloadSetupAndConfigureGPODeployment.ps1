# Importing all required functions
. $scriptPath\Setup-GPOOfficeInstallation.ps1

# Set the Office channel files path, bitness, channels, and languages
$OfficeFilesPath = "C:\OfficeChannelFiles"
$Bitness = @("v64")
$Channels = @("Current")
$Languages = @("en-us")

# Download the channel files
Download-GPOOfficeChannelFiles -OfficeFilesPath $OfficeFilesPath -Bitness $Bitness -Channels $Channels -Languages $Languages

# Configure the deployment files
Configure-GPOOfficeDeployment -Channel Current -Bitness 64 -OfficeSourceFilesPath $OfficeFilesPath -MoveSourceFiles $true

#------------------------------------------------------------------------------------------------------------
#   Customize Deployment Script - Uncomment and modify the code below to customize this deployment script
#------------------------------------------------------------------------------------------------------------

  #### ------- Deploy Office using a script ------- ####
  # Create-GPOOfficeDeployment -GroupPolicyName DeployCurrentChannel64Bit -DeploymentType DeployWithScript -Channel Current -Bitness 64
  
  #### ------- Deploy Office using a standard configuration xml file ------- ####
  # Create-GPOOfficeDeployment -GroupPolicyName DeployCurrentChannel64Bit -DeploymentType DeployWithConfigurationFile -Channel Current -Bitness 64 -ConfigurationXML CurrentChannelDeployment.xml
  
  #### ------- Deploy Office using an installation file ------- ####
  # Create-GPOOfficeDeployment -GroupPolicyName DeployCurrentChannelWithMSI -DeploymentType DeployWithInstallationFile -OfficeDeploymentFileName OfficeProPlus.msi

#------------------------------------------------------------------------------------------------------------