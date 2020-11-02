$Excelfile = 'sentinel-rule-templates.xlsx'
$RulePath = 'sentinel-rule-templates.json'
$dataset=Get-Content -Path $RulePath | ConvertFrom-Json | Select-Object -ExpandProperty value

$Collection =@()
foreach ($object in $dataset)
    {        
    $Props = @{
        'ID' = $object.'name'
        'Name' = $object.'properties'.'displayName'
        'Query' = $object.'properties'.'query'
        'Severity' = $object.'properties'.'severity'
        'Platform' = $object.'properties'.'requiredDataConnectors'.'connectorId'
        'DataType' = $object.'properties'.'requiredDataConnectors'.'dataTypes'
        'Tactic' = $object.'properties'.'tactics'
        'Description' = $object.'properties'.'description'
        }
    $TotalObjects = New-Object PSCustomObject -Property $Props
    if ($TotalObjects.Revoked -notcontains $true){
        $Collection += $TotalObjects }
    }

$Collection | Select-Object  @{Name ="ID"; Expression={$_.ID | Select-Object -Index 0 }},@{Name ="Name"; Expression={$_.Name -join ","}},@{Name="Platform";Expression={$_.'Platform' -join ","}},@{Name="DataType";Expression={$_.'DataType' -join ","}},@{Name="Tactic";Expression={$_.'Tactic' -join ","}},@{Name="Description";Expression={$_.'Description' -join ","}},@{Name="Query";Expression={$_.'Query' -join ","}} | Sort-Object ID | Export-Excel $Excelfile -WorksheetName Templates      