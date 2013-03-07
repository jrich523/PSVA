## enums

Function New-Enum {
param($name,$hash)
### thanks jaykul
if(!("VisioAutomation.$name" -as [type]))
{
Add-Type -TypeDefinition @"
namespace VisioAutomation { public enum $name {
$( $hash.Keys | % { "$_ = {0}," -f $hash.$_ } )
}}
"@
}
}

#Used: 
New-Enum AutoConnectDir @{Down=2;Left=3;Right=4;Up=1;None=0}
#Used: Selectionion.align
New-Enum HorizontalAlignTypes @{Center=2;Left=1;Right=3;None=0}
#Used: Seletionion.align
New-Enum VerticalAlignTypes @{Bottom=3;Middle=2;top=1;None=0}
#Used: Selection.distribute
New-Enum DistributeTypes @{HorCenter=2;HorLeft=1;HorRight=3;HorSpace=0;VertBottom=7;VertMiddle=6;VertSpace=4;VertTop=5}
#Used: Selection.LayoutIncremental
New-Enum LayoutIncrementalType @{Align=1;Space=2}
New-Enum LayoutHorzAlignType @{None=0;Default=1;Left=2;Center=3;Right=4}
New-Enum LayoutVertAlignType @{None=0;Default=1;Top=2;Middle=3;Bottom=4}
##used to set pinpoint, Custom Type data and a function, maybe drop the function
New-Enum PinPoint @{CenterCenter = 0;CenterLeft = 1;CenterRight = 2;TopCenter = 3;TopLeft = 4;TopRight = 5;BottomCenter = 6;BottomRight = 7;BottomLeft = 8;Custom = 9}


New-Enum ObjectTypes @{
Addon=31;
Addons=32;
ApplicationSettings=51;
App=3;
Cell=4;
Chars=5;
Color=29;
Colors=30;
Connect=8;
Connects=9;
ContainerProperties=60;
Curve=42;
DataColumn=56;
DataColumns=55;
DataConnection=54;
DataRecordset=53;
DataRecordsetChangedEvent=57;
DataRecordsets=52;
Doc=10;
Docs=11;
EventList=34;
Event=33;
Font=27;
Fonts=28;
Global=36;
GraphicItem=59;
GraphicItems=58;
Hyperlink=37;
Hyperlinks=43;
KeyboardEvent=50;
Layer=25;
Layers=26;
MasterShortcut=47;
MasterShortcuts=46;
Master=12;
Masters=13;
MouseEvent=49;
MovedSelectionEvent=62;
MSGWrap=48;
OLEObject=39;
OLEObjects=38;
Page=14;
Pages=15;
Path=41;
Paths=40;
RelatedShapePairEvent=61;
Row=45;
Section=44;
Selection=16;
ServerPublishOptions=63;
Shape=17;
Shapes=18;
Style=19;
Styles=20;
Unknown=1;
Validation=64;
ValidationIssue=70;
ValidationIssues=69;
ValidationRule=68;
ValidationRules=67;
ValidationRuleSet=66;
ValidationRuleSets=65;
Window=21;
Windows=22;
}