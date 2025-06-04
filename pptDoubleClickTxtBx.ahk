/*

AutoHotkey v2.0 script to create a textbox in PowerPoint when doubleclicking on a slide.
Intended to mimic the behavior seen in the macOS version of PowerPoint

Author: mayonster
Creation date: 6/3/2025
Revision: A
Revision Date: 6/3/2025
Revision Notes: Initial issue.


*/
#Requires AutoHotkey v2.0

;Global variables
doubleClickThreshold := 300
lastClickTime := 0

~LButton:: {
    global lastClickTime, doubleClickThreshold  ;Declare globals used inside this function

    if WinActive("ahk_exe POWERPNT.EXE") { ;Sets powerpoint as the only application to sniff for a doubleclick
        currentTime := A_TickCount
        if (currentTime - lastClickTime <= doubleClickThreshold) {
        
            PPT := ComObjActive("PowerPoint.Application") ;Sets PowerPoint as the active application
            PPT.CommandBars.ExecuteMso("TextBoxInsert") ;Uses the XML processing instruction to activate textbox insert function

            MouseGetPos(&xpos, &ypos) ;Gets the position you double clicked on

            ;Adjust the size of the textbox you make
            drawTargetX := xpos + 250 ;X Coordinate
            drawTargetY := ypos + 75 ;Y Coordinate

            Sleep(90) ;Necessary delay for PowerPoint to register draw textbox mode, tunable

            SendInput "{Click xpos ypos Down}"
            MouseMove(drawTargetX, drawTargetY, 0)
            Send "{CLick drawTargetX drawTargetY Up}"
            MouseMove(xpos, ypos, 0) ;I found i had to move the mouse as if creating a textbox manually for this to work
        }
        lastClickTime := currentTime
    }
}




/*
 /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\ 
( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )
 > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ < 
 /\_/\                                                                            .x+=:.        s                              /\_/\ 
( o.o )                                   ..                                     z`    ^%      :8                             ( o.o )
 > ^ <     ..    .     :                 @L                 u.      u.    u.        .   <k    .88                  .u    .     > ^ < 
 /\_/\   .888: x888  x888.        u     9888i   .dL   ...ue888b   x@88k u@88c.    .@8Ned8"   :888ooo      .u     .d88B :@8c    /\_/\ 
( o.o ) ~`8888~'888X`?888f`    us888u.  `Y888k:*888.  888R Y888r ^"8888""8888"  .@^%8888"  -*8888888   ud8888.  ="8888f8888r  ( o.o )
 > ^ <    X888  888X '888>  .@88 "8888"   888E  888I  888R I888>   8888  888R  x88:  `)8b.   8888    :888'8888.   4888>'88"    > ^ < 
 /\_/\    X888  888X '888>  9888  9888    888E  888I  888R I888>   8888  888R  8888N=*8888   8888    d888 '88%"   4888> '      /\_/\ 
( o.o )   X888  888X '888>  9888  9888    888E  888I  888R I888>   8888  888R   %8"    R88   8888    8888.+"      4888>       ( o.o )
 > ^ <    X888  888X '888>  9888  9888    888E  888I u8888cJ888    8888  888R    @8Wou 9%   .8888Lu= 8888L       .d888L .+     > ^ < 
 /\_/\   "*88%""*88" '888!` 9888  9888   x888N><888'  "*888*P"    "*88*" 8888" .888888P`    ^%888*   '8888c. .+  ^"8888*"      /\_/\ 
( o.o )    `~    "    `"`   "888*""888"   "88"  888     'Y"         ""   'Y"   `   ^"F        'Y"     "88888%       "Y"       ( o.o )
 > ^ <                       ^Y"   ^Y'          88F                                                     "YP'                   > ^ < 
 /\_/\                                         98"                                                                             /\_/\ 
( o.o )                                      ./"                                                                              ( o.o )
 > ^ <                                      ~`                                                                                 > ^ < 
 /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\  /\_/\ 
( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )( o.o )
 > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ <  > ^ < 
 */