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
function Add-visStencil
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true)]
        $Path,
        #Pass back the Stencil Document object
        [switch]$PassThru
    )

    $stencil = $script:visApplication.Documents.OpenEx($path,4)
    if($PassThru){$stencil}
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
function Get-visStencil
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$false, Position=0)]
        $Title
    )

    if($title)
    {
        $script:visApplication.documents | ? {$_.type -eq 2 -and $_.title -like "*$title*"} |  %{$_}
    }
    else
    {
        $script:visApplication.documents | ? {$_.type -eq 2} | %{$_}
    }
}


Export-ModuleMember -Function *