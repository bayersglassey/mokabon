Option _Explicit
Option Base 0

Const TILES = 100
Const TILE_WIDTH = 16

Dim tiles(TILES, TILE_WIDTH, TILE_WIDTH) As Integer



Randomize Timer
_Title "Gameboy"
_ScreenMove _Middle
_Limit 32




Dim Image&(4) ' 4 image surfaces
Dim Count% '    image counter

'* prepare image surfaces

For Count% = 1 To 4 '                                  cycle 4 times
    Image&(Count%) = _NewImage(640, 480, 32) '         create a new surface image
    _Dest Image&(Count%) '                              make the surface the destination
    Cls '                                               clear the surface
    Locate 2, 2 '                                       position text cursor
    Print "This is image number"; Count% '              print the surface number
    Circle (Count% * 100, 300), 50 '                    draw a circle on the surface
Next Count%
Count% = 1 '                                           reset image counter

'* display each surface

Do '                                                   main program loop
    Screen Image&(Count%) '                             use image as current screen
    Locate 4, 2 '                                       position text cursor
    Print "Press ENTER to switch to the next screen." ' print directions
    Print " Press ESC to exit."
    Sleep '                                             wait for key stroke
    Count% = Count% + 1 '                              increment image counter
    If Count% = 5 Then Count% = 1 '                     keep count within limits
Loop Until _KeyDown(27) '                              leave when escape key pressed
System '                                               return to operating system


