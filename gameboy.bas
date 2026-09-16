Option _Explicit ' Need to declare all variables (with Dim) before using them
Option Base 1 ' Array indexes start at 1, not 0

Const True = -1
Const False = 0

' Frames per second (how fast the animation is)
Const FPS = 30

' Set up random number generator
Randomize Timer

' Size of the gameboy's screen in pixels
Const TrueScreenWidth = 160
Const TrueScreenHeight = 144

' We "zoom" the screen, that is, stretch it when rendering it.
' So, if ScreenZoom is 3, then each pixel on the gameboy's screen becomes
' a 3x3 square on the computer's screen.
Const ScreenZoom = 3
Const ScreenWidth = TrueScreenWidth * ScreenZoom
Const ScreenHeight = TrueScreenHeight * ScreenZoom

' The image on the game boy's screen
Dim Shared ScreenImage As Long
ScreenImage = _NewImage(TrueScreenWidth, TrueScreenHeight, 32)

' Load some images, generally for use as tilesets
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
SetCharacterMap 1, 0, 7, "$*"
SetCharacterMap 1, 3, 7, "/,"
SetCharacterMap 1, 6, 7, "0123456789"
SetCharacterMap 1, 13, 8, ":"

' Width and height of tiles, in pixels
Const TileWidth = 8
Const TileHeight = 8

' A tileset is a grid of "tiles" stored within an image
Type Tileset
    Image As Long ' An open image, see TitleImage, TilesetsImage, etc

    ' (X, Y) coordinates of the top-left corner of this tileset within
    ' its image
    StartX As Long
    StartY As Long

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
FontTileset.AddX = 9
FontTileset.AddY = 9
FontTileset.StartX = 8
FontTileset.StartY = 24
FontTileset.CharacterMap = 1 ' Use the first character map

' Variables used when writing text to the screen using a font (see
' WriteAt, WriteText)
Dim Shared WriteX As Long
Dim Shared WriteY As Long

' Used by subroutines Parse, NextToken
Dim Shared ParseText As String
Dim Shared Token As String


' #################################################################
' # DECLARATIONS RELATED TO THE MAP

' The image on which we draw the map
Dim Shared MapImage As Long

' Set up the map tilesets, which all use the same image, but whose tiles
' come from different offsets within that image.
' NOTE: the bounds 0 to 18 are just the map numbers given in
' "img/tilesets.png", we didn't invent them
Const MapTilesetLower = 0
Const MapTilesetUpper = 18
Dim Shared MapTilesets(MapTilesetLower To MapTilesetUpper) As Tileset
MapTilesets(0).Image = TilesetsImage
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

' The file to which we will save the map
Dim MapFilename As String

' Index into MapTilesets
Dim Shared MapTilesetNumber As Long

' Map tiles aren't the same as the regular tiles stored in a Tileset.
' Each map tile is actually a 2x2 square of regular tiles, plus some
' information affecting whether you can walk onto it, etc.
Const MapTileWidth = TileWidth * 2
Const MapTileHeight = TileHeight * 2
Type MapTile
    Solid As Integer
    HasPokemon As Integer

    ' Top/bottom left/right: tile indexes, to be interpreted as (X, Y)
    ' coordinates of tiles within MapTilesets(I), using "Mod 16" to get the
    ' X and "/ 16" to get the Y
    TL As Long
    TR As Long
    BL As Long
    BR As Long
End Type

' NOTE: MaxMapTiles is an arbitrary number, just big enough to support
' all map tile files we ever try to load (see LoadMapTiles)
Const MaxMapTiles = 50
Dim Shared MapTiles(0 To MaxMapTiles - 1) As MapTile

' Width and height of the map, in "map tiles" (see the MapTile type).
' The elements of Map are indices into MapTiles.
Dim Shared MapWidth As Long
Dim Shared MapHeight As Long
ReDim Shared Map(MapWidth - 1, MapHeight - 1) As Long

' Load the map!..
' NOTE: MapFilename might change later, if the user wants to save the map
' to a different file.
MapFilename = "maps/test0.txt"
LoadMap MapFilename


' #################################################################
' # DECLARATIONS RELATED TO CHARACTERS


' #################################################################
' # DECLARATIONS RELATED TO THE PLAYER

Dim Shared PlayerX As Long
Dim Shared PlayerY As Long
Dim Shared PlayerScrollX As Long
Dim Shared PlayerScrollY As Long


' ########################################################################
' # THAT'S THE END OF ALL THE DECLARATIONS!
' # NOW WE ACTUALLY CREATE A WINDOW AND START THE GAME!

' Set up the window/screen
Screen _NewImage(ScreenWidth, ScreenHeight, 32)
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

    ' Show whatever we've drawn on the screen
    RenderScreen

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
    X As Long, Y As Long _
)
    ' NOTE: this subroutine expects _Dest to already be set!..

    ' Source and destination (X, Y) coordinates
    Dim SrcX As Long
    Dim SrcY As Long
    Dim DstX As Long
    Dim DstY As Long

    SrcX = Tileset.StartX + TileX * Tileset.AddX
    SrcY = Tileset.StartY + TileY * Tileset.AddY
    DstX = X * TileWidth
    DstY = Y * TileHeight
    _PutImage _
        (DstX, DstY)-(DstX + TileWidth - 1, DstY + TileHeight - 1), _
        Tileset.Image, _Dest, _
        (SrcX, SrcY)-(SrcX + TileWidth - 1, SrcY + TileHeight - 1)
End Sub

Sub WriteText(Text As String)
    Dim I As Integer
    Dim Ch As Integer
    Dim Entry As XYPair

    _Dest ScreenImage ' Write to the game boy's screen

    ' Now loop over the characters in the text, and draw each character on
    ' the screen, using the tiles in FontTileset
    For I = 1 To Len(Text)
        Ch = Asc(Mid$(Text, I, 1)) ' Get the next character from Text
        Entry = CharacterMapEntries(FontTileset.CharacterMap, Ch)
        RenderTile FontTileset, Entry.X, Entry.Y, WriteX, WriteY
        WriteX = WriteX + 1
    Next
End Sub

Sub Die(Message As String)
    _Dest 0 ' Print to the screen
    Cls ' Clear the screen
    Locate 1, 1
    Print "*** ERROR ***"
    Print Message
    _Display ' Show the message
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
    '_Dest = 0: Print "Parsed token: [" + Token + "]"
End Sub

Sub LoadMapTiles
    Dim Filename As String
    Dim File As Long
    Dim Text As String
    Dim LineNumber As Long
    Dim ParsingBottom As Long

    ' E.g. "tilesets/0.txt"
    Filename = "tilesets/" + LTrim$(Str$(MapTilesetNumber)) + ".txt"

    ' Index into MapTiles
    Dim I As Long
    I = 0

    File = FreeFile
    Open Filename For Input As File
    Do Until Eof(File)
        Line Input #File, Text
        LineNumber = LineNumber + 1
        If Text = "" Or Left$(Text, 1) = "#" Then
            ' Empty line or comment, ignore it!
        ElseIf Text = "solid" Then
            MapTiles(I).Solid = True
        ElseIf Text = "pokemon" Then
            MapTiles(I).HasPokemon = True
        ElseIf ParsingBottom Then
            ' Parsing bottom two tiles of this map tile
            MapTiles(I).BL = Val("&H" + Left$(Text, 2))
            MapTiles(I).BR = Val("&H" + Mid$(Text, 4, 2))
            ParsingBottom = False
            I = I + 1
        Else
            ' Parsing top two tiles of this map tile
            MapTiles(I).TL = Val("&H" + Left$(Text, 2))
            MapTiles(I).TR = Val("&H" + Mid$(Text, 4, 2))
            ParsingBottom = True
        End If
    Loop
    Close File

    If ParsingBottom Then
        Die "Hit end of file while still parsing map tile" + Str$(I)
    End If
End Sub

Sub LoadMap(Filename As String)
    Dim File As Long
    Dim Text As String
    Dim LineNumber As Long
    Dim X As Long
    Dim Y As Long

    MapTilesetNumber = 0

    File = FreeFile
    Open Filename For Input As File
    Do Until Eof(File)
        Line Input #File, Text
        LineNumber = LineNumber + 1
        If Text = "" Or Left$(Text, 1) = "#" Then
            ' Empty line or comment, ignore it!
        ElseIf Left$(Text, 7) = "tileset" Then
            Parse Text
            NextToken ' ignore "tileset"
            NextToken
            MapTilesetNumber = Val(Token)
        ElseIf Left$(Text, 5) = "tiles" Then
            Parse Text
            NextToken ' ignore "tiles"
            NextToken
            MapWidth = Val(Token)
            NextToken
            MapHeight = Val(Token)
            ReDim Map(0 To MapWidth - 1, 0 To MapHeight - 1) As Long
            If MapImage > 0 Then _FreeImage MapImage
            MapImage = _NewImage(MapWidth * MapTileWidth, _
                MapHeight * MapTileHeight, 32)
            For Y = 0 To MapHeight - 1
                Line Input #File, Text
                Parse Text
                For X = 0 To MapWidth - 1
                    NextToken
                    Map(X, Y) = Val("&H" + Token)
                Next
            Next
        Else
            Die "Don't know what to do with line" + Str$(LineNumber) _
                + ": [" + Text + "]"
        End If
    Loop
    Close File

    LoadMapTiles

    PlayerX = MapWidth / 2
    PlayerY = MapHeight / 2
    PlayerScrollX = 0
    PlayerScrollY = 0

    RenderMapImage
End Sub

' Render a map tile, that is, a 2x2 square of tiles
Sub RenderMapTile(X As Long, Y As Long)
    Dim MapTile As MapTile
    MapTile = MapTiles(Map(X, Y))

    ' Number encoding an (X, Y) coordinate into the map's tileset
    Dim XY As Long

    ' Render tiles onto MapImage
    _Dest MapImage

    ' Top-left tile
    XY = MapTile.TL
    RenderTile MapTilesets(MapTilesetNumber), _
        XY Mod 16, Int(XY / 16), X * 2 + 0, Y * 2 + 0

    ' Top-right tile
    XY = MapTile.TR
    RenderTile MapTilesets(MapTilesetNumber), _
        XY Mod 16, Int(XY / 16), X * 2 + 1, Y * 2 + 0

    ' Bottom-left tile
    XY = MapTile.BL
    RenderTile MapTilesets(MapTilesetNumber), _
        XY Mod 16, Int(XY / 16), X * 2 + 0, Y * 2 + 1

    ' Bottom-right tile
    XY = MapTile.BR
    RenderTile MapTilesets(MapTilesetNumber), _
        XY Mod 16, Int(XY / 16), X * 2 + 1, Y * 2 + 1
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
    ' Copy the map onto the game boy's screen
    Dim MapScrollX As Long
    Dim MapScrollY As Long
    MapScrollX = TrueScreenWidth / 2 - PlayerX * MapTileWidth - PlayerScrollX
    MapScrollY = TrueScreenHeight / 2 - PlayerY * MapTileHeight - PlayerScrollY
    _PutImage _
        (MapScrollX, MapScrollY) - ( _
            MapScrollX + MapWidth * MapTileWidth - 1, _
            MapScrollY + MapHeight * MapTileHeight - 1 _
        ), _
        MapImage, ScreenImage
End Sub

Sub RenderScreen
    ' Draw the game boy's screen (that is, ScreenImage) on the actual screen
    ' (that is, the program's window)
    _PutImage (0, 0)-(ScreenWidth - 1, ScreenHeight - 1), ScreenImage, 0
    _Display
End Sub
