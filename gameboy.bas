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

' Set up the map tilesets, which all use the same image, but whose tiles
' come from different offsets within that image.
Dim Shared MapTilesets(0 To 18) As Tileset
MapTilesets(0).Image = TilesetsImage
MapTilesets(0).TileWidth = 8
MapTilesets(0).TileHeight = 8
MapTilesets(0).AddX = 8
MapTilesets(0).AddY = 8
MapTilesets(0).StartX = 3
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

' The current tileset, a copy of one of the elements of MapTilesets
Dim Shared MapTileset As Tileset
MapTileset = MapTilesets(0)

Type MapTile
    ' (X, Y) coordinates into the tileset, that is, into MapTileset
    XY As XYPair
End Type

Dim Shared MapWidth As Long
Dim Shared MapHeight As Long
ReDim Shared MapTiles(MapWidth, MapHeight) As MapTile


' #################################################################
' # THAT'S THE END OF ALL THE DECLARATIONS!
' # NOW WE ACTUALLY CREATE A WINDOW AND START THE GAME!

' Set up random number generator
Randomize Timer

' Set up the window/screen
Screen _NewImage(640, 480, 32)
_Title "Gameboy"
_ScreenMove _Middle

WriteText "Hello world!"

' Enter the main loop!..
Do
    ' Make sure the animation doesn't go faster than our intended
    ' frames-per-second (FPS)
    _Limit FPS

    'RenderMap

    If _KeyDown(Asc("h")) Then
        Locate 2, 2
        Print "HELP!"
        Do: _Limit FPS: Loop While _KeyDown(Asc("h"))
    End If
Loop Until _KeyDown(27) ' Quit if escape key is pressed

System ' Close the program without saying "Press any key..."


' ######################################################################################
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

Sub WriteText(Text As String)
    Dim I As Integer
    Dim Ch As Integer
    Dim Entry As XYPair

    ' Source and destination (X, Y) coordinates
    Dim SrcX As Long
    Dim SrcY As Long
    Dim DstX As Long
    Dim DstY As Long
    Dim TileW As Long
    Dim TileH As Long

    ' Set up some variables...
    _Source WriteTileset.Image
    TileW = WriteTileset.TileWidth
    TileH = WriteTileset.TileHeight

    ' Now loop over the characters in the text, and draw each character on
    ' the screen, using the tiles in WriteTileset
    For I = 1 To Len(Text)
        Ch = Asc(Mid$(Text, I, 1)) ' Get the next character from Text
        Entry = CharacterMapEntries(WriteTileset.CharacterMap, Ch)
        SrcX = WriteTileset.StartX + Entry.X * WriteTileset.AddX
        SrcY = WriteTileset.StartY + Entry.Y * WriteTileset.AddY
        DstX = WriteX * TileW
        DstY = WriteY * TileH
        _PutImage (DstX, DstY)-(DstX + TileW, DstY + TileH), _Source, 0, (SrcX, SrcY)-(SrcX + TileW, SrcY + TileH)
        WriteX = WriteX + 1
    Next
End Sub
