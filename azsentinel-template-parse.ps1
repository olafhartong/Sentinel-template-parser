$Excelfile = 'sentinel-rule-templates.xlsx'
$RulePath = 'azsentinel-rule-templates.json'
$Workspace = '[changeme]'

Get-AzSentinelAlertRuleTemplates -WorkspaceName $Workspace | ConvertTo-Json -Depth 3| Out-File $RulePath
$dataset=Get-Content -Path $RulePath | ConvertFrom-Json

$Collection =@()
foreach ($object in $dataset)
    {        
    $Props = @{
        'ID' = $object.'name'
        'Name' = $object.'displayName'
        'Query' = $object.'query'
        'Severity' = $object.'severity'
        'Platform' = $object.'requiredDataConnectors'.'connectorId'
        'DataType' = $object.'requiredDataConnectors'.'dataTypes'
        'Tactic' = $object.'tactics'
        'Description' = $object.'description'
        }
    $TotalObjects = New-Object PSCustomObject -Property $Props
    if ($TotalObjects.Revoked -notcontains $true){
        $Collection += $TotalObjects }
    }

$Collection | Select-Object  @{Name ="ID"; Expression={$_.ID | Select-Object -Index 0 }},@{Name ="Name"; Expression={$_.Name -join ","}},@{Name="Platform";Expression={$_.'Platform' -join ","}},@{Name="DataType";Expression={$_.'DataType' -join ","}},@{Name="Tactic";Expression={$_.'Tactic' -join ","}},@{Name="Description";Expression={$_.'Description' -join ","}},@{Name="Query";Expression={$_.'Query' -join ","}} | Sort-Object ID | Export-Excel $Excelfile -WorksheetName Templates
