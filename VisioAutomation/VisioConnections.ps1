#Connections

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
function Add-visShapeConnection
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $ToShape,

        # Param2 help description
        #[VisioAutomation.AutoConnectDir]
        #$FromDirection,

        # param3
        [Parameter(Mandatory=$true, Position=1)]
        $FromShape
    )

    Begin
    {
        $page = Get-visPage
    }
    Process
    {
        if(-not $script:visMasterConn){$script:visMasterConn = Get-visStencil | %{$_.masters} | ? {$_.nameU -eq "Dynamic connector"} | select -first 1}
        $conn = $page.drop($script:visMasterConn,0,0)
        $cb = $conn.Cells("beginx")
        $ce = $conn.Cells("endx")
        $cb.GlueTo($FromShape.cells("pinx"))
        $ce.GlueTo($ToShape.cells('pinx'))
        $conn
    }
    End
    {
    }
}


Export-ModuleMember -Function *