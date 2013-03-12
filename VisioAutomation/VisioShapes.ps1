
##generic shape functions
<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Get-visShape
{
    [CmdletBinding()]
    Param
    (
        # Shape name to find
        [Parameter(Position=0)]
        $Name,
        [switch]$WithConnectors
    )

    $shapes = (Get-visPage).shapes | %{$_}
    if($WithConnectors)
    {
        $shapes
    }
    else
    {
        $shapes | ? {$_.master.name -ne "Dynamic connector"}
    }
}

<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Get-VisConnector
{
    [CmdletBinding()]
    Param()
    (Get-visPage).shapes | %{$_} | ? {$_.master.name -eq "Dynamic connector"}
}

<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Add-visShape
{
    [CmdletBinding()]
    Param
    (
        # Name of Shape to add
        [Parameter(Mandatory=$true, Position=0)]
        [string]
        $ShapeName,
        # pick a shape from a certain stencil
        [Parameter(Mandatory=$false, Position=1)]
        [string]
        $StencilName,
        #set X value drop point 
        [Parameter(Mandatory=$false, Position=2)]
        [double]
        $xPos=0,
        [Parameter(Mandatory=$false, Position=3)]
        [double]
        $yPos=0,
        [Parameter(Mandatory=$false, Position=4)]
        [string]
        $ShapeText
    )

    if($StencilName)
    {
        $sten = $script:visApplication.Documents | ?{ $_.type -eq 2} | ? {$_.title -like "*$StencilName*" -or $_.name -like "*$StencilName*"}
    }
    else
    {
        $sten = $script:visApplication.Documents | ?{ $_.type -eq 2}
    }

    $obj =  $sten | %{$_.masters} | ? { $_.name -like "*$ShapeName*"} | select -first 1

    $shape = (Get-visPage).drop($obj,$xPos,$yPos)
    if($ShapeText)
    {
        $Shape.text = $ShapeText
    }
    $shape
}

<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Set-visShapeProperty
{
    [CmdletBinding()]
    Param
    (
        # Shape to modify property on
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $shape,
        # Property to set
        [string]
        $Property,
        # Value of property
        [string]
        $Value
    )

    Begin
    {
    }
    Process
    {
        if([bool]$shape.cellExists($property,0))
        {
            $shape.cells($property).FormulaU = "`"$value`""
        }
        elseif([bool]$shape.cellExists("prop.$property",0))
        {
            $shape.cells("prop.$property").FormulaU = "`"$value`""
        }
        else
        {
            Write-Error "Property not found!"
        }
    }
    End
    {
    }
}

<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Select-visShape
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $Shape
    )

    Begin
    {
        $Selection = (Get-visApplication).activewindow.selection
        $Selection.DeselectAll()
    }
    Process
    {
        $Selection.Select($shape,2)
    }
    End
    {
        ,$Selection
    }
}

<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Group-visShape
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Selection
    )

    $selection.group()
}

##change/move shape
<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
Function Set-visShapeDistribution
{
param(
$shapes,
[ValidateSet("Vertical","Horizontal")]
$type,
$space=0,
[double]
$startX,
[double]
$startY
)
    if($type -like "v*")
    {
        $prop = "piny"
        $staticprop = "pinx"
        $difftype = "height"
        [double]$mover = $startY
        $loc = $startX
    }
    else
    {
        $prop = "pinx"
        $staticprop = "piny"
        $difftype = "width"
        [double]$mover = $startx
        $loc = $startY
    }

    foreach($shape in $shapes)
    {
        $thisdiff = $shape.cells($difftype).resultiu / 2
        if($lastdiff -ne $null)
        { 
            $mover = $mover + $thisdiff + $lastdiff + $space
        }
        $shape.cells($prop).ResultIU = $mover
        $shape.cells($staticprop).ResultIU = $loc
        
        $lastdiff = $thisdiff

    }
}

<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Set-visShapeAlignmentOn
{
[CmdletBinding()]
param(
[paramter(manditory=$true)]
$shapes,
# The Page inch location you'd like to align on 
$AlignLocation, 
[ValidateSet("Vertical","Horizontal")]
$type)

$prop = ""
if($typ -eq "Vertical")
{$prop="piny"}
else
{$prop="pinx"}
foreach($shape in $shapes){
$shape.cells($prop) = $AlignLocation
}

}

<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Set-visShapeAutoAlignment
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]
        $Selection
    )

    $selection.LayoutIncremental([int][VisioAutomation.LayoutIncrementalType]::Align,[int][VisioAutomation.LayoutHorzAlignType]::Default,[int][VisioAutomation.LayoutVertAlignType]::Default, 0,0,65);

}


<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Set-visShapePinPoint
{
[CmdletBinding()]
param(
#shape to change pin type on
[parameter(position=0)]
$Shape,
#pin type = default is center/center
[parameter(position=1)]
[VisioAutomation.PinPoint]$pinpoint="CenterCenter")

    switch($pinpoint)
    {
    "CenterCenter" { $shape.cells("locpinx").formula = "Width * 0.5";  $shape.cells("locpiny").formula = "Height * 0.5"}
    "CenterLeft" { $shape.cells("locpinx").formula = 0;  $shape.cells("locpiny").formula = "Height * 0.5"}
    "CenterRight" { $shape.cells("locpinx").formula = "Width";  $shape.cells("locpiny").formula = "Height * 0.5"}
    
    "TopCenter" { $shape.cells("locpinx").formula = "Width * 0.5";  $shape.cells("locpiny").formula = "Height"}
    "TopLeft" { $shape.cells("locpinx").formula = 0;  $shape.cells("locpiny").formula = "Height"}
    "TopRight" { $shape.cells("locpinx").formula = "Width";  $shape.cells("locpiny").formula = "Height"}
    
    "BottomCenter" { $shape.cells("locpinx").formula = "Width * 0.5";  $shape.cells("locpiny").formula = 0}
    "BottomRight" { $shape.cells("locpinx").formula = "Width";  $shape.cells("locpiny").formula = 0}
    "BottomLeft" { $shape.cells("locpinx").formula = 0;  $shape.cells("locpiny").formula = 0}
    }
}

<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Set-visShapePosition
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Shape,

        # Param2 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [Double]
        $xPos,
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [double]
        $yPos
    )

    Begin
    {
    }
    Process
    {
        $shape.Cells("pinx") = $xPos
        $shape.Cells("piny") = $yPos
    }
    End
    {
    }
}

<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Remove-visShape
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,Position=0)]
        $Shape
    )

    Begin
    {
    }
    Process
    {
        $shape.delete()
    }
    End
    {
    }
}

<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Attach-visShape
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $Shape,

        # The selection of items to attach it too
        
        $Selection,
        [VisioAutomation.Side]
        $Side,
        [VisioAutomation.Alignment]
        $Alignment
    )

    Begin
    {
    }
    Process
    {
        #todo: check for pinposition
        #todo: support multiple shapes
        ########## assumes center pin!
        #All maxes
        $maxtop = $Selection | select -ExpandProperty top | sort | select -last 1
        $maxbottom = $Selection | select -ExpandProperty bottom | sort | select -first 1
        $maxleft = $Selection | select -ExpandProperty left | sort | select -first 1
        $maxright = $Selection | select -ExpandProperty right | sort | select -Last 1
        $width = $maxright - $maxleft
        $height = $maxtop - $maxbottom
        
        
        switch($Alignment)
        {
            "LeftOrTop"
            {
                if([int]$side -le 1) #top/bottom - Left applies, set x
                {
                    $shape.pinx = $maxleft + ($shape.width /2)
                }
                else #left/right - top applies, set y
                {
                    $shape.piny = $maxtop - ($shape.Height /2)
                    
                }
            }
            "RightOrBottom"
            {
                if([int]$side -le 1) #top/bottom
                {
                    $shape.pinx = $maxright - ($shape.width /2)
                }
                else
                {
                    $shape.piny = $maxbottom + ($shape.Height /2)
                }
            }
            "Stretch" 
            {
                if([int]$side -le 1) #top/bottom - right applies
                {
                    $shape.Width = $width
                    $shape.pinx = $maxleft + (($maxright - $maxleft) /2)

                }
                else #left/right applies
                {
                    $shape.Height = $height
                    $shape.Piny = $maxbottom + (($maxtop - $maxbottom) /2)
                }
            }
        }

        ##side
        switch($side)
        {
            "Top"
            {
                $shape.piny = $maxtop + ($shape.Height /2)
            }
            "Bottom"
            {
                $Shape.Piny = $maxbottom - ($shape.Height /2)
            }
            "Left"
            {
                $shape.pinx = $maxleft - ($shape.width /2)
            }
            "Right"
            {
                $Shape.PinX = $maxright + ($shape.width /2)
            }
        }
    }
    End
    {
    }
}

Export-ModuleMember -Function * 