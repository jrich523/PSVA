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
function Connect-visApplication
{
    [CmdletBinding()]
    
    Param
    (
        #File to open
        [Parameter(Position=0)]
        [string]$path,
        # Force New instance
        [switch]$New
    )

    $isOpen = if(gps "visio" -ErrorAction SilentlyContinue){$true}else{$false}
    if($new -or -not $isOpen)
    {
        $script:visApplication = New-Object -ComObject Visio.Application
    }
    else
    {
        $script:visApplication = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Visio.Application")
    }

    $documents = $script:visApplication.Documents

    $script:visStencilShapes = $documents | ? {$_.type -eq 2} | %{$_.masters}
    
    if($path)
    {
        $script:visActiveDocument = $documents.Add($path)
    }
    elseif($new -or -not $isOpen)
    {
        $script:visActiveDocument = $documents.Add("")
    }
    else
    {
        $script:visActiveDocument = $script:visApplication.ActiveDocument
    }

    $script:visActivePage = $script:visApplication.ActivePage
    ## load by default, has Dynamic Connector
    $basflo = Add-visStencil basflo_u.vssx -PassThru
    $script:visMasterConn = $basflo.Masters.ItemU("Dynamic connector")
    
}

Function Disconnect-visApplication{
##TODO: maybe prompt to save?
$script:visApplication.close()}

Function Get-visApplication { return $script:visApplication}
Export-ModuleMember -Function *