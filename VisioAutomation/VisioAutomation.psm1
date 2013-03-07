
#. $psScriptRoot\VisioEnums.ps1  #loaded by module psd1
. $psScriptRoot\VisioApplication.ps1
. $psScriptRoot\VisioDocument.ps1
. $psScriptRoot\VisioPage.ps1
. $psScriptRoot\VisioShapes.ps1
. $psScriptRoot\VisioStencil.ps1
. $psScriptRoot\VisioConnections.ps1


$myInvocation.MyCommand.ScriptBlock.Module.OnRemove = { 
    Clear-FormatData
}