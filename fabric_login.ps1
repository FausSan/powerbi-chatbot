Install-Module -Name MicrosoftPowerBIMgmt 

Connect-PowerBIServiceAccount
 
$accessToken = Get-PowerBIAccessToken
 
$uri="https://api.powerbi.com/v1.0/myorg/groups/33aec4ac-1ca0-4c4f-9701-26995e5f64e4/datasets/b0032ad6-16a4-4667-8bd8-50b4ed1f455a"
 
Invoke-RestMethod -Uri $uri –Headers $accessToken –Method get -ContentType application/json
