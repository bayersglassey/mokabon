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

Type XYPair
    X As Long
    Y As Long
End Type

' Maps ASCII characters onto (X, Y) coordinates of tiles
Dim Shared CharacterMapEntries(1, 127) As XYPair

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

    ' Number of characters in a row of the tileset
    Width As Long

    ' How much to add to the (X, Y) coordinates of a tile to get to the
    ' ones next to it
    AddX As Long
    AddY As Long

    ' Index into first dimension of CharacterMapEntries
    CharacterMap As Long
End Type

Dim Shared FontTileset As Tileset

' Variables used when writing text to the screen using a font
Dim Shared WriteTileset As Tileset
Dim Shared WriteX As Long
Dim Shared WriteY As Long


' #################################################################
' # NOW LET'S START ACTUALLY DOING STUFF!

' Set up random number generator
Randomize Timer

' Open some images
TitleImage = _LoadImage("img/title.png", 32)
TilesetsImage = _LoadImage("img/tilesets.png", 32)
CharactersImage = _LoadImage("img/characters.png", 32)
FontImage = _LoadImage("img/font.png", 32)
MenusImage = _LoadImage("img/menus.png", 32)
PokemonImage = _LoadImage("img/pokemon.png", 32)
PortraitsImage = _LoadImage("img/portraits.png", 32)

' Where to find the English font's tiles
FontTileset.Image = FontImage
FontTileset.TileWidth = 8
FontTileset.TileHeight = 8
FontTileset.AddX = 9
FontTileset.AddY = 9
FontTileset.Width = 16
FontTileset.StartX = 8
FontTileset.StartY = 24
FontTileset.CharacterMap = 1 ' Use the first character map
SetCharacterMap FontTileset, 0, 0, "ABCDEFGHIJKLMNOP"
SetCharacterMap FontTileset, 0, 1, "QRSTUVWXYZ():;[]"
SetCharacterMap FontTileset, 0, 2, "abcdefghijklmnop"
SetCharacterMap FontTileset, 0, 3, "qrstuvwxyz"
SetCharacterMap FontTileset, 0, 4, " "
SetCharacterMap FontTileset, 0, 6, "'"
SetCharacterMap FontTileset, 3, 6, "-"
SetCharacterMap FontTileset, 6, 6, "?!."
SetCharacterMap FontTileset, 3, 7, "/,"
SetCharacterMap FontTileset, 6, 7, "0123456789"
SetCharacterMap FontTileset, 13, 8, ":"

' Choose which tileset to use when "writing" text to the screen
WriteTileset = FontTileset

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

Sub WriteAt(X As Long, Y As Long)
    WriteX = X
    WriteY = Y
End Sub

Sub SetCharacterMap(T As Tileset, X As Long, Y As Long, Text As String)
    Dim I As Integer
    Dim Ch As String
    For I = 1 To Len(Text)
        Ch = Mid$(Text, I, 1) ' Get the next character from Text
        CharacterMapEntries(T.CharacterMap, Asc(Ch)).X = X
        CharacterMapEntries(T.CharacterMap, Asc(Ch)).Y = Y
        X = X + 1
    Next
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
