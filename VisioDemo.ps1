#demo
$server = 'localhost'
$shares = gwmi win32_share -ComputerName $server

Connect-visApplication
$app = Get-visApplication
$page = Get-visPage

Add-visStencil basic_u.vssx

#create server shape
$servershape = Add-visShape -ShapeName square -xPos 4.25 -yPos 5 -ShapeText $server

#create share shape

$shareShapes = $shares | %{ Add-visShape -ShapeName square -xPos 0 -yPos 3 -ShapeText "$($_.name)`n$($_.path)" }
Set-visShapeDistribution -shape $shareShapes  -type Horizontal -space .5 -startX 2 -startY 3


$connections = $shareShapes | Add-visShapeConnection -FromShape $servershape

$connLayer = $connections | Add-visShapeToLayer -Layer ConnLayer -Force 

Get-visShape | Select-visShape | Set-visShapeAutoAlignment
sleep 5
## connectors are added to a default "connector" layer
Switch-visLayerVisibility $connLayer,"connector"
