[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ScriptFromGitHub = Invoke-WebRequest https://raw.githubusercontent.com/volvo64/dcgScripts/master/SPLA/SPLAScriptv1.ps1
Invoke-Expression $($ScriptFromGitHub.Content)