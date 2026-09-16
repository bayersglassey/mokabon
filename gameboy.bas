Option _Explicit ' Need to declare all variables (with Dim) before using them
Option Base 1 ' Array indexes start at 1, not 0

' Frames per second (how fast the animation is)
Const FPS = 30

' Size of the screen in pixels
Const ScreenWidth = 160
Const ScreenHeight = 144

' Images we have loaded
Dim Shared TitleImage As Long
Dim Shared TilesetsImage As Long
Dim Shared CharactersImage As Long
Dim Shared FontImage As Long
Dim Shared MenusImage As Long
Dim Shared PokemonImage As Long
Dim Shared PortraitsImage As Long
TitleImage = _LoadImage("img/title.png", 32)
TilesetsImage = _LoadImage("img/tilesets.png", 32)
CharactersImage = _LoadImage("img/characters.png", 32)
FontImage = _LoadImage("img/font.png", 32)
MenusImage = _LoadImage("img/menus.png", 32)
PokemonImage = _LoadImage("img/pokemon.png", 32)
PortraitsImage = _LoadImage("img/portraits.png", 32)

' Represents an (X, Y) coordinate
Type XYPair
    X As Long
    Y As Long
End Type

' Maps ASCII characters onto (X, Y) coordinates of tiles
Dim Shared CharacterMapEntries(1, 127) As XYPair
SetCharacterMap 1, 0, 0, "ABCDEFGHIJKLMNOP"
SetCharacterMap 1, 0, 1, "QRSTUVWXYZ():;[]"
SetCharacterMap 1, 0, 2, "abcdefghijklmnop"
SetCharacterMap 1, 0, 3, "qrstuvwxyz"
SetCharacterMap 1, 0, 4, " "
SetCharacterMap 1, 0, 6, "'"
SetCharacterMap 1, 3, 6, "-"
SetCharacterMap 1, 6, 6, "?!."
SetCharacterMap 1, 3, 7, "/,"
SetCharacterMap 1, 6, 7, "0123456789"
SetCharacterMap 1, 13, 8, ":"

Type Tileset
    ' A tileset is a grid of "tiles" stored within an image

    Image As Long ' An open image, see TitleImage, TilesetsImage, etc

    ' (X, Y) coordinates of the top-left corner of this tileset within
    ' its image
    StartX As Long
    StartY As Long

    ' Width and height of the tiles inside this tileset
    TileWidth As Long
    TileHeight As Long

    ' How much to add to the (X, Y) coordinates of a tile to get to the
    ' ones next to it
    AddX As Long
    AddY As Long

    ' Index into first dimension of CharacterMapEntries, or 0 if this
    ' tileset isn't for a font
    CharacterMap As Long
End Type

Dim Shared FontTileset As Tileset
FontTileset.Image = FontImage
FontTileset.TileWidth = 8
FontTileset.TileHeight = 8
FontTileset.AddX = 9
FontTileset.AddY = 9
FontTileset.StartX = 8
FontTileset.StartY = 24
FontTileset.CharacterMap = 1 ' Use the first character map

' Variables used when writing text to the screen using a font
Dim Shared WriteTileset As Tileset
Dim Shared WriteX As Long
Dim Shared WriteY As Long
WriteTileset = FontTileset


' #################################################################
' # DECLARATIONS RELATED TO THE MAP

' Used by subroutines Parse, NextToken
Dim Shared ParseText As String
Dim Shared Token As String

' The image on which we draw the map
Dim Shared MapImage As _Unsigned Long

' Set up the map tilesets, which all use the same image, but whose tiles
' come from different offsets within that image.
' NOTE: the bounds 0 to 18 are just the map numbers given in
' "img/tilesets.png", we didn't invent them
Dim Shared MapTilesets(0 To 18) As Tileset
MapTilesets(0).Image = TilesetsImage
MapTilesets(0).TileWidth = 8
MapTilesets(0).TileHeight = 8
MapTilesets(0).AddX = 8
MapTilesets(0).AddY = 8
MapTilesets(0).StartX = 2
MapTilesets(0).StartY = 176
InitializeMapTileset 1, 234
InitializeMapTileset 2, 292
InitializeMapTileset 3, 350
InitializeMapTileset 4, 408
InitializeMapTileset 5, 466
InitializeMapTileset 6, 524
InitializeMapTileset 7, 582
InitializeMapTileset 8, 640
InitializeMapTileset 9, 698
InitializeMapTileset 10, 756
InitializeMapTileset 11, 814
InitializeMapTileset 12, 872
InitializeMapTileset 13, 930
InitializeMapTileset 14, 988
InitializeMapTileset 15, 1046
InitializeMapTileset 16, 1104
InitializeMapTileset 17, 1162
InitializeMapTileset 18, 1220

Dim MapFilename As String

' The current tileset, a copy of one of the elements of MapTilesets
Dim Shared MapTileset As Tileset

Type MapTile
    ' (X, Y) coordinates into the tileset, that is, into MapTileset
    X As Long
    Y As Long
End Type

Const MapTileWidth = 8
Const MapTileHeight = 8

Dim Shared MapWidth As Long
Dim Shared MapHeight As Long
ReDim Shared MapTiles(0, 0) As MapTile

Dim Shared MapScrollX As Long
Dim Shared MapScrollY As Long

Dim Shared MapZoom As Long
MapZoom = 2

LoadMap "maps/test0.txt"


' ########################################################################
' # THAT'S THE END OF ALL THE DECLARATIONS!
' # NOW WE ACTUALLY CREATE A WINDOW AND START THE GAME!

' Set up random number generator
Randomize Timer

' Set up the window/screen
Screen _NewImage(640, 480, 32)
_Title "Gameboy"
_ScreenMove _Middle

' Enter the main loop!..
Do
    Cls 0 ' Clear the screen
    RenderMap
    WriteAt 0, 0
    WriteText "Hello world!"

    ' While the H key is being held down, show the "help" message
    If _KeyDown(Asc("h")) Then
        _Dest 0
        Locate 2, 2
        Print "Keyboard controls:"
        Print " H: show this help"
        Print " Escape: quit the program"
    End If

    _Display ' Show whatever we've drawn on the screen

    ' Make sure the animation doesn't go faster than our intended
    ' frames-per-second (FPS)
    _Limit FPS
Loop Until _KeyDown(27) ' Quit if escape key is pressed

System ' Close the program without saying "Press any key..."


' ########################################################################
' # FUNCTION AND SUBROUTINE DEFINITIONS

Sub SetTilesetClearColor(T As Tileset)
    ' Set the "clear color", i.e. the transparent color, for the given
    ' tileset's image
    _Source T.Image ' Set the source image to be used by Point
    _ClearColor Point(T.StartX, T.StartY), T.Image
End Sub

Sub InitializeMapTileset(I As Long, StartY As Long)
    MapTilesets(I) = MapTilesets(0)
    MapTilesets(I).StartY = StartY
End Sub

Sub SetCharacterMap(CharacterMap As Long, X As Long, Y As Long, Text As String)
    Dim I As Integer
    Dim Ch As String
    For I = 1 To Len(Text)
        Ch = Mid$(Text, I, 1) ' Get the next character from Text
        CharacterMapEntries(CharacterMap, Asc(Ch)).X = X
        CharacterMapEntries(CharacterMap, Asc(Ch)).Y = Y
        X = X + 1
    Next
End Sub

Sub WriteAt(X As Long, Y As Long)
    WriteX = X
    WriteY = Y
End Sub

Sub RenderTile( _
    Tileset As Tileset, TileX As Long, TileY As Long, _
    X As Long, Y As Long, ExtraX As Long, ExtraY As Long _
)
    ' NOTE: this subroutine expects _Dest to already be set!..

    ' Tile width and height
    Dim TileW As Long
    Dim TileH As Long
    TileW = Tileset.TileWidth
    TileH = Tileset.TileHeight

    ' Source and destination (X, Y) coordinates
    Dim SrcX As Long
    Dim SrcY As Long
    Dim DstX As Long
    Dim DstY As Long

    SrcX = Tileset.StartX + TileX * Tileset.AddX
    SrcY = Tileset.StartY + TileY * Tileset.AddY
    DstX = X * TileW + ExtraX
    DstY = Y * TileH + ExtraY
    _PutImage _
        (DstX, DstY)-(DstX + TileW - 1, DstY + TileH - 1), _
        Tileset.Image, _Dest, _
        (SrcX, SrcY)-(SrcX + TileW - 1, SrcY + TileH - 1)
End Sub

Sub WriteText(Text As String)
    Dim I As Integer
    Dim Ch As Integer
    Dim Entry As XYPair

    _Dest 0 ' Write to the screen

    ' Now loop over the characters in the text, and draw each character on
    ' the screen, using the tiles in WriteTileset
    For I = 1 To Len(Text)
        Ch = Asc(Mid$(Text, I, 1)) ' Get the next character from Text
        Entry = CharacterMapEntries(WriteTileset.CharacterMap, Ch)
        RenderTile WriteTileset, Entry.X, Entry.Y, WriteX, WriteY, 0, 0
        WriteX = WriteX + 1
    Next
End Sub

Sub Die(Message As String)
    Print Message
    End
End Sub

Sub Parse(Text As String)
    ParseText = Text
End Sub

Sub NextToken
    Dim I As Long
    I = Instr(ParseText, " ")
    If I > 0 Then
        Token = Left$(ParseText, I)
        ParseText = Mid$(ParseText, I + 1)
    Else
        Token = ParseText
        ParseText = ""
    End If
    'Print "Parsed token: [" + Token + "]"
End Sub

Sub LoadMap(Filename As String)
    Dim File As Long
    Dim Text As String
    Dim X As Long
    Dim Y As Long
    Dim I As Long

    File = FreeFile
    Open Filename For Input As File
    Do Until Eof(File)
        Line Input #File, Text
        If Left$(Text, 7) = "tileset" Then
            Parse Text
            NextToken
            NextToken
            MapTileset = MapTilesets(Val(Token))
        ElseIf Left$(Text, 5) = "tiles" Then
            Parse Text
            NextToken
            NextToken
            MapWidth = Val(Token)
            NextToken
            MapHeight = Val(Token)
            ReDim MapTiles(0 To MapWidth - 1, 0 To MapHeight - 1) As MapTile
            If MapImage > 0 Then _FreeImage MapImage
            MapImage = _NewImage(MapWidth * MapTileWidth, _
                MapHeight * MapTileHeight, 32)
            For Y = 0 To MapHeight - 1
                Line Input #File, Text
                Parse Text
                For X = 0 To MapWidth - 1
                    NextToken
                    I = VAL("&H" + Token)
                    MapTiles(X, Y).X = I Mod 16
                    MapTiles(X, Y).Y = Int(I / 16)
                Next
            Next
        ElseIf Text = "" Then
            ' Empty line, ignore it!
        Else
            Die "Don't know what to do with line: " + Text
        End If
    Loop
    Close File

    RenderMapImage
End Sub

Sub RenderMapTile(X As Long, Y As Long)
    _Dest MapImage
    RenderTile MapTileset, MapTiles(X, Y).X, MapTiles(X, Y).Y, X, Y, 0, 0
End Sub

Sub RenderMapImage
    ' Render the map's tiles onto the map's image
    Dim X As Long
    Dim Y As Long
    For Y = 0 To MapHeight - 1
        For X = 0 To MapWidth - 1
            RenderMapTile X, Y
        Next
    Next
End Sub

Sub RenderMap
    ' Copy the map onto the screen, taking into account scrolling and zooming
    _PutImage _
        (MapScrollX * MapZoom, MapScrollY * MapZoom) - ( _
            (MapScrollX + MapWidth * MapTileWidth) * MapZoom - 1, _
            (MapScrollY + MapHeight * MapTileHeight) * MapZoom - 1 _
        ), _
        MapImage, 0
End Sub
