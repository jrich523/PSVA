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
function Get-visPage
{
    [CmdletBinding()]
    Param
    (
        # Page index to get
        [Parameter(Position=0)]
        [int]
        $Index
    )

    if($index)
    {
        $pages = $script:visDocument.pages | %{$_}
        return $pages[$Index+1]
        
    }
    else
    {
        return $script:visApplication.ActivePage
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
function New-visLayer
{
    [CmdletBinding()]
    Param
    (
        # Name for the new layer
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Name
    )

    (Get-visPage).Layers.Add($name)
    
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
function Get-visLayer
{
    [CmdletBinding()]
    Param
    (
        # Name of the layer to get
        [Parameter(ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Name
    )
    $layers = Get-visPage | select -exp layers
    if($Name)
    {
        $layers | ?{$_.name -eq "$name"}
    }
    else
    {
        $layers
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
function Add-visShapeToLayer
{
    [CmdletBinding()]
    Param
    (
        # Name of the layer you'd like to add the shape to
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Layer,

        # Shape object to add to the Layer
        [Parameter(Mandatory=$true,
            ValueFromPipelineByPropertyName=$true,
            ValueFromPipeline=$true,
            Position=1)]
        $Shape,

        # If the layer isnt there, create it
        [switch]
        $Force
    )

    Begin
    {
        #is it a layer object?
        if($Layer.ObjectType -eq [VisioAutomation.ObjectTypes]::Layer)
        {
            Write-Verbose "Layer object passed in"
            $LayerObj =  $Layer
        }
        #no, then get it by name (assumption of string)
        else
        {
            Write-Verbose "Looking for layer by name"
            $LayerObj = Get-visLayer $Layer
        }
        # it isnt an object or a known name, so create it.

        if((-not $LayerObj) -and $Force)
        {
            Write-Verbose "Creating new layer"
            $layerobj = New-visLayer $Layer
        }
        elseif(-not $LayerObj -and -not $force)
        {
            Write-Error "LAYER NOT FOUND! If you would like it to auto create then please specify Force" -ErrorAction Stop
        }

    }
    Process
    {
        $LayerObj.add($shape,1) #Zero to remove subshapes from any previous layer assignments; non-zero to preserve layer assignments.

    }
    End
    {
        $layerobj
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
function Switch-visLayerVisibility
{
    [CmdletBinding()]
    Param
    (
        # Name of layer to toggle
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $Layer
    )

    Begin
    {

    }
    Process
    {
        
        
        foreach($layerItem in $layer)
        {
            if($LayerItem.ObjectType -eq [VisioAutomation.ObjectTypes]::Layer)
        {
            Write-Verbose "Layer object passed in"
            $LayerObj =  $layerItem
        }
        #no, then get it by name (assumption of string)
        else
        {
            Write-Verbose "Looking for layer by name"
            $LayerObj = Get-visLayer $layerItem
        }
        # it isnt an object or a known name write error
        if((-not $LayerObj))
        {
            Write-Error "LAYER NOT FOUND! If you would like it to auto create then please specify Force"
        }

            if($LayerObj.CellsC(4).resultiu -eq 1)
            {
                $LayerObj.CellsC(4).resultiu=0
            }
            else
            {
                $LayerObj.CellsC(4).resultiu=1
            }
        }
    }
    End
    {
    }
}


Export-ModuleMember -Function *