import-module D:\ps\wc\Modules\VisioAutomation
Connect-visApplication
Add-visStencil basic_u.vssx
##create two base shapes to work with
$s1 = Add-visShape -ShapeName square -ShapeText "text1"
$s2 = Add-visShape -ShapeName square -ShapeText "text2"
$sel = Get-visShape | Select-visShape

#create shape to attach to base
$top = Add-visShape -ShapeName rectangle -ShapeText "im the top!"
$top.Height = .5

$bottom = Add-visShape -ShapeName rectangle -ShapeText "im the bottom!"
$bottom.Height = .5
$bottom.Width = .75

$left = Add-visShape -ShapeName rectangle -ShapeText "im the left!"
$left.height = 1

$right = Add-visShape -ShapeName rectangle -ShapeText "im the right!"
$right.Height = 1

#set base shape position
Set-visShapeDistribution -shapes ($s1,$s2) -type Horizontal -startX 3 -startY 4

#attach shapes
Attach-visShape -Shape $top -Selection ($s1,$s2) -Side top -Alignment Stretch
Attach-visShape -Shape $bottom -Selection ($s1,$s2) -side Bottom -Alignment RightOrBottom
Attach-visShape -Shape $left -Selection ($s1,$s2) -Side Left -Alignment LeftOrTop
Attach-visShape -Shape $right -Selection ($s1,$s2) -Side right -Alignment Stretch



