Import-Module Visio
New-VisioApplication
New-VisioDocument
$basic_u = Open-VisioDocument basic_u.vss
$master = Get-VisioMaster "Rectangle" $basic_u
$shape = New-VisioShape $master 3,3
Set-VisioShapeText "Hello World"