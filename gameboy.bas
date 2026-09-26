Option _Explicit ' Need to declare all variables (with Dim) before using them
Option Base 1 ' Array indexes start at 1, not 0

Const True = -1
Const False = 0
Dim Shared QUOTE As String
QUOTE = Chr$(34) ' Literal quote character (")

' Frames per second (how fast the animation is)
Const FPS = 30

' Set to True to jump everywhere instead of walking... for debugging purposes!
Const ALWAYS_JUMP = False

' Set up random number generator
Randomize Timer

' Used by KeyPressed
Dim Shared PrevKeyCode As Long

' Keyboard key codes
Const UpCode = 18432
Const DownCode = 20480
Const LeftCode = 19200
Const RightCode = 19712
Const EnterCode = 13
Const EscapeCode = 27
Const F5Code = 16128
Const F7Code = 16640

' Size of the gameboy's screen in pixels
Const TrueScreenWidth = 160
Const TrueScreenHeight = 144

' We "zoom" the screen, that is, stretch it when rendering it.
' So, if ScreenZoom is 3, then each pixel on the gameboy's screen becomes
' a 3x3 square on the computer's screen.
Const ScreenZoom = 3
Const ScreenWidth = TrueScreenWidth * ScreenZoom
Const ScreenHeight = TrueScreenHeight * ScreenZoom

' The program's current "mode", e.g. whether you're walking around, or
' editing the map, etc
Dim Shared Mode As String
Const GAME_MODE = ""
Const MAP_EDITOR_MODE = "Map Editor"
Const MAP_SCROLL_MODE = "Map Scrolling Tool"
Const MAP_RESIZE_MODE = "Map Resizing Tool"
Const TILE_SELECTOR_MODE = "Tile Selector"
Const CHARACTER_EDITOR_MODE = "Character Editor"
Mode = GAME_MODE

' When Mode = TILE_SELECTOR_MODE, we render MapTiles as a grid, and this
' is the width of that grid (in map tiles).
Const TileSelectorWidth = 8

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

    ' Width and height of tiles within this tileset
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

' Variables used when writing text to the screen using a font (see
' WriteAt, WriteText)
Dim Shared WriteX As Long
Dim Shared WriteY As Long
Dim Shared WriteStartX As Long
Dim Shared WriteWidth As Long

' Used by subroutines Parse, NextToken
Dim Shared ParseText As String
Dim Shared Token As String

' Used by parsing subroutines
Dim Shared LineNumber As Long

' General-purpose global variables for subroutines to return information
' about rectangles into.
' Top-left corner is (RectangleX1, RectangleY1)
' Bottom-left corner is (RectangleX2, RectangleY2)
Dim Shared RectangleX1 As Long
Dim Shared RectangleY1 As Long
Dim Shared RectangleX2 As Long
Dim Shared RectangleY2 As Long


' #################################################################
' # DECLARATIONS RELATED TO SCRIPTS

Const COMMAND_WAIT = 0
Const COMMAND_WALK = 1
Const COMMAND_JUMP = 2
Const COMMAND_FACE = 3
Const COMMAND_SAY = 4
Const COMMAND_TALK = 5

Type ScriptCommand
    CommandType As Integer ' COMMAND_WAIT, etc
    Str1 As String
    Num1 As Long
    Num2 As Long
End Type

Type Script
    CharacterNumber As Long ' Index into Characters
    Start As Long ' Index into ScriptCommands
    Length As Long ' Number of commands in this script

    ' True if this script should loop around to the beginning, otherwise False
    ShouldLoop As Integer
End Type

Type ScriptState
    ScriptNumber As Long
    CommandNumber As Long ' Between 1 and Script.Length
    Frame As Long ' Frame of animation for current command
End Type

ReDim Shared ScriptCommands(0) As ScriptCommand
ReDim Shared Scripts(0) As Script
ReDim Shared ScriptStates(1) As ScriptState

' If we're currently talking to someone, Talking = True, and the state of
' the script for that conversation is ScriptStates(TALKING_SCRIPT_STATE).
Const TALKING_SCRIPT_STATE = 1
Dim Shared Talking As Integer
Dim Shared TalkingText As String
Talking = False


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

' The file to which we will save the map
Dim MapFilename As String

' Index into MapTilesets
Dim Shared MapTilesetNumber As Long

' Solidity of map tiles, see Solidity field of type MapTile
Const NOT_SOLID = 0
Const SOLID = 1
Const JUMP_UP = 2
Const JUMP_DOWN = 3
Const JUMP_LEFT = 4
Const JUMP_RIGHT = 5

' Map tiles aren't the same as the regular tiles stored in a Tileset.
' Each map tile is actually a 2x2 square of regular tiles, plus some
' information affecting whether you can walk onto it, etc.
Const MapTileWidth = TileWidth * 2
Const MapTileHeight = TileHeight * 2
Type MapTile
    Solidity As Integer ' See SOLID, JUMP_DOWN, etc
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

' Current number of maptiles, i.e. entries of MapTiles
Dim Shared NumMapTiles As Long

' Width and height of the map, in "map tiles" (see the MapTile type).
' The elements of Map are indices into MapTiles.
Dim Shared MapWidth As Long
Dim Shared MapHeight As Long
ReDim Shared Map(MapWidth - 1, MapHeight - 1) As Long

' The "anchor point" in the map editor
Dim Shared MapEditorAnchorX As Long
Dim Shared MapEditorAnchorY As Long


' #################################################################
' # DECLARATIONS RELATED TO CHARACTERS

Dim Shared CharacterTileset As Tileset
CharacterTileset.Image = CharactersImage
CharacterTileset.TileWidth = 16
CharacterTileset.TileHeight = 16
CharacterTileset.AddX = 17
CharacterTileset.AddY = 17
CharacterTileset.StartX = 9
CharacterTileset.StartY = 34
SetTilesetClearColor CharacterTileset

' When pokemon show up on the map, they use this tileset.
' See also the Character type's IsPokemon field, which just controls
' whether this tileset is used.
Dim Shared PokemonCharacterTileset As Tileset
PokemonCharacterTileset = CharacterTileset
PokemonCharacterTileset.StartY = 994

' For non-character things which appear on the map, like items, or
' your shadow when you jump.
Dim Shared MiscCharacterTileset As Tileset
MiscCharacterTileset = CharacterTileset
MiscCharacterTileset.StartY = 1087

' Directions a character can be facing
' NOTE: these numbers are chosen so that turning to the right means adding
' 1 (Mod 4).
Const FACING_UP = 0
Const FACING_RIGHT = 1
Const FACING_DOWN = 2
Const FACING_LEFT = 3

' States a character can be in, that is, things they can be doing
Const STATE_STANDING = 0
Const STATE_WALKING = 1
Const STATE_JUMPING = 2
Const STATE_SHORT_JUMPING = 3
Const STATE_GONE = 4 ' Don't render, collide with, etc this character

Type Character
    ' Character's name; should be unique within a given map
    Name As String

    ' If True, this character is invisible.
    ' Can be used to implement "triggers", i.e. map locations which do
    ' something when you step on them or "talk" to them.
    IsHidden As Integer

    ' Whether we should use PokemonCharacterTileset instead of
    ' CharacterTileset
    IsPokemon As Integer

    ' Character's position on the map, in map tiles
    X As Long
    Y As Long

    ' When drawing the character, add these extra pixel coordinates, so
    ' that they can smoothly move between map tiles
    ExtraX As Long
    ExtraY As Long

    ' Start of this character's tiles within CharacterTileset.
    ' This should basically never change: it determines what this character
    ' look like, for instance if it's the player character, or a boy with
    ' glasses, or a girl with a headband, or a bald man, etc.
    TileStartX As Long
    TileStartY As Long

    ' Offset from TileStartX, TileStartY to the tile character should use
    ' when rendered.
    ' This is updated each frame of animation according to the character's
    ' State and Frame. See HandleCharacterAnimation for details.
    TileAddX As Long
    TileAddY As Long

    ' Which direction the character is facing, for instance, FACING_UP
    Facing As Integer

    ' The character's state, that is, what they are doing.
    ' For instance, STATE_STANDING
    State As Integer

    ' Frame of animation, that is, how far along in the animation the
    ' character is for their State.
    ' For instance, if the character is walking (State = STATE_WALKING),
    ' their Frame starts at 0 and goes up by 1 until the walking animation
    ' is finished, and State goes back to STATE_STANDING.
    Frame As Long

    ' Value which alternates between 0 and 1.
    ' Determines whether left or right foot is used by the walking animation.
    OtherFoot As Integer

    ' Index into ScriptStates, or 0
    ScriptStateNumber As Long

    ' Index into Scripts
    TalkScriptsStart As Long

    ' Number of entries in Scripts
    TalkScriptsLength As Long

    ' Combined with TalkScriptsStart to indicate the talk script to be
    ' used.
    ' If 0, no talk script is used; otherwise, the script to be used
    ' is Scripts(TalkScriptsStart + TalkScriptNumber - 1)
    TalkScriptNumber As Integer

    ' Starting values for other character fields, set when the map is loaded,
    ' and used when the map is saved.
    StartX As Long
    StartY As Long
    StartFacing As Integer
End Type

ReDim Shared Characters(1) As Character


' #################################################################
' # DECLARATIONS RELATED TO THE MAP EDITOR

' The map tiles which can be put on the map by pressing the number
' keys 0-9
Dim Shared SelectedMapTiles(10) As Integer
SelectedMapTiles(1) = 0
SelectedMapTiles(2) = 1
SelectedMapTiles(3) = 2
SelectedMapTiles(4) = 3
SelectedMapTiles(5) = 4
SelectedMapTiles(6) = 5
SelectedMapTiles(7) = 6
SelectedMapTiles(8) = 7
SelectedMapTiles(9) = 8
SelectedMapTiles(10) = 9
Dim Shared SelectedMapTileNumber As Long

' The currently selected character, if any (used by CHARACTER_EDITOR_MODE)
Dim Shared SelectedCharacter As Long


' #################################################################
' # DECLARATIONS RELATED TO THE PLAYER

' Index into Characters
Const PLAYER = 1

' TODO: items


' ########################################################################
' # THAT'S THE END OF ALL THE DECLARATIONS!
' # NOW WE ACTUALLY CREATE A WINDOW AND START THE GAME!

' Load the map!..
' NOTE: MapFilename might change later, if the user wants to save the map
' to a different file.
MapFilename = "maps/test0.txt"
LoadMap MapFilename

' Set up the window/screen
Screen _NewImage(ScreenWidth, ScreenHeight, 32)
_Title "Gameboy"
_ScreenMove _Middle

' Enter the main loop!..
Do
    ' Clear the screen
    _Dest ScreenImage
    Cls

    Dim X As Long, Y As Long, I As Long
    Dim NewX As Long, NewY As Long
    Dim MoveDirection As Long

    If Mode = GAME_MODE And Talking Then
        If ScriptStateDone(TALKING_SCRIPT_STATE) Then Talking = False
    End If

    ' Mode-specific behaviour
    If Mode = GAME_MODE And Talking Then
        ' We're talking to another character, so pause the game while we
        ' run the talk script...

        If TalkingText <> "" And KeyPressed(Asc("z")) Then TalkingText = ""

        If TalkingText = "" Then
            UpdateScriptState TALKING_SCRIPT_STATE
            HandleCharacterAnimation _
                Scripts(TALKING_SCRIPT_STATE).CharacterNumber
        End If

        RenderMap
        For I = 1 To UBound(Characters)
            RenderCharacter I
        Next
        If TalkingText <> "" Then RenderTextBox TalkingText
    ElseIf Mode = GAME_MODE Then
        If KeyPressed(Asc("m")) Then
            Mode = MAP_EDITOR_MODE
            MapEditorAnchorX = -1 ' Anchor starts off unset
        End If

        ' Handle player's controls, that is, react to keys the player
        ' is pressing
        If Characters(PLAYER).State = STATE_STANDING Then
            ' Handle arrow keys
            MoveDirection = GetPlayerMoveDirection(True)
            If MoveDirection >= 0 Then
                ' A single arrow key was pressed!.. so, let's walk in
                ' that direction.
                Characters(PLAYER).Facing = MoveDirection
                NewX = PlayerX + FacingAddX(MoveDirection)
                NewY = PlayerY + FacingAddY(MoveDirection)
                Dim CanMove As Integer
                CanMove = CanMoveTo(PlayerX, PlayerY, NewX, NewY, _
                    MoveDirection, PLAYER)
                If ALWAYS_JUMP Then CanMove = 2 ' For debugging!
                If CanMove = 1 Then
                    ' We are ok to walk to the new map position
                    Characters(PLAYER).X = NewX
                    Characters(PLAYER).Y = NewY
                    Characters(PLAYER).State = STATE_WALKING
                ElseIf CanMove = 2 Then
                    ' We are ok to jump to the new map position
                    Characters(PLAYER).X = NewX + FacingAddX(MoveDirection)
                    Characters(PLAYER).Y = NewY + FacingAddY(MoveDirection)
                    Characters(PLAYER).State = STATE_JUMPING
                ElseIf CanMove = 3 Then
                    ' We are ok to short jump to the new map position
                    Characters(PLAYER).X = NewX
                    Characters(PLAYER).Y = NewY
                    Characters(PLAYER).State = STATE_SHORT_JUMPING
                End If
            End If

            ' Handle gameboy's "A" button
            If KeyPressed(Asc("z")) Then
                NewX = PlayerX + FacingAddX(Characters(Player).Facing)
                NewY = PlayerY + FacingAddY(Characters(Player).Facing)
                I = CollideCharacters(NewX, NewY, PLAYER)
                If I > 0 Then
                    If Characters(I).State = STATE_STANDING Then
                        ' We're talking to another character!..
                        ' Get them to face us, and run their talk script,
                        ' if any.
                        Characters(I).Facing = _
                            (Characters(PLAYER).Facing + 2) Mod 4
                        If Characters(I).TalkScriptNumber > 0 Then
                            Talking = True
                            SetScriptState TALKING_SCRIPT_STATE, _
                                Characters(I).TalkScriptsStart _
                                    + Characters(I).TalkScriptNumber - 1
                        End If
                    End If
                End If
            End If
        End If

        ' Update all characters
        For I = 1 To UBound(Characters)
            If Characters(I).ScriptStateNumber Then
                UpdateScriptState Characters(I).ScriptStateNumber
            End If
            HandleCharacterAnimation I
        Next

        ' Render the map onto the game boy's screen
        RenderMap

        ' Render all characters
        For I = 1 To UBound(Characters)
            RenderCharacter I
        Next
    ElseIf Mode = MAP_EDITOR_MODE Then
        ' Move the player with the arrow keys; in map editor mode, the
        ' player is invisible, and in their place is a box showing the
        ' current map location (that is, tile) to be edited.
        HandleEditorArrowKeys

        ' Set/unset the "anchor point"
        If KeyPressed(Asc("a")) Then
            If MapEditorAnchorX < 0 Then
                ' Set the anchor point to player's current location
                MapEditorAnchorX = PlayerX
                MapEditorAnchorY = PlayerY
            Else
                ' Unset the anchor point if it's already at player's current
                ' location
                MapEditorAnchorX = -1
            End If
        End If
        UpdateMapEditorAnchorRectangle

        ' Handle number keys
        For I = 0 To 9
            ' Edit the map if a number key was pressed.
            If KeyPressed(Asc("0") + I) Then
                For X = RectangleX1 To RectangleX2
                    For Y = RectangleY1 To RectangleY2
                        Map(X, Y) = SelectedMapTiles(I + 1)
                        RenderMapTile X, Y
                    Next
                Next
            EndIf

            ' Update a selected tile if shift + number key was pressed.
            If KeyPressed(Asc(Mid$(")!@#$%^&*(", I + 1, 1))) Then
                SelectedMapTiles(I + 1) = Map(PlayerX, PlayerY)
            EndIf
        Next

        ' Render the map onto the game boy's screen
        RenderMap

        ' Draw the map tiles currently selected for use with number keys
        ' 0-9 at the bottom of the screen
        RenderSelectedMapTiles

        ' Map saving/loading
        If KeyPressed(Asc("f")) Then
            _Dest 0
            Print "Current map filename: " + MapFilename
            Input "Change map filename: ", MapFilename
            If Not Instr(MapFilename, "/") Then _
                MapFilename = "maps/" + MapFilename
            If Not Instr(MapFilename, ".") Then _
                MapFilename = MapFilename + ".txt"
            ' The enter key was just pressed (because we used Input), so
            ' make sure we don't immediately exit the map editor because
            ' of that!..
            PrevKeyCode = EnterCode
        End If
        If KeyPressed(F5Code) Then
            SaveMap MapFilename
            ShowMessage "Map saved!"
        End If
        If KeyPressed(F7Code) Then
            LoadMap MapFilename
            ShowMessage "Map loaded!"
        End If

        ' Reload various images, files, etc
        If KeyPressed(Asc("l")) Then
            LoadMapTiles
            RenderMapImage
            ShowMessage "Map tiles reloaded!"
        End If

        ' Maybe switch to a different mode
        If KeyPressed(Asc("m")) Or KeyPressed(EnterCode) Then _
            Mode = GAME_MODE
        If KeyPressed(Asc("t")) Then Mode = TILE_SELECTOR_MODE
        If KeyPressed(Asc("s")) Then Mode = MAP_SCROLL_MODE
        If KeyPressed(Asc("r")) Then Mode = MAP_RESIZE_MODE
        If KeyPressed(Asc("c")) Then
            Mode = CHARACTER_EDITOR_MODE
            SelectedCharacter = 0
        End If
    ElseIf Mode = MAP_SCROLL_MODE Then
        RenderMap
        HandleMapScrollMode
        If KeyPressed(Asc("s")) Or KeyPressed(EnterCode) Then _
            Mode = MAP_EDITOR_MODE
    ElseIf Mode = MAP_RESIZE_MODE Then
        RenderMap
        HandleMapResizeMode
        If KeyPressed(Asc("r")) Or KeyPressed(EnterCode) Then _
            Mode = MAP_EDITOR_MODE
    ElseIf Mode = TILE_SELECTOR_MODE Then
        ' Change the currently selected map tile
        If KeyPressed(UpCode) Then SelectedMapTileNumber = _
            SelectedMapTileNumber - TileSelectorWidth
        If KeyPressed(DownCode) Then SelectedMapTileNumber = _
            SelectedMapTileNumber + TileSelectorWidth
        If KeyPressed(LeftCode) Then SelectedMapTileNumber = _
            SelectedMapTileNumber - 1
        If KeyPressed(RightCode) Then SelectedMapTileNumber = _
            SelectedMapTileNumber + 1
        If SelectedMapTileNumber < 0 Then SelectedMapTileNumber = 0
        If SelectedMapTileNumber >= NumMapTiles Then _
            SelectedMapTileNumber = NumMapTiles - 1

        ' Render all map tiles as a grid
        RenderMapTiles

        ' Draw the map tiles currently selected for use with number keys
        ' 0-9 at the bottom of the screen
        RenderSelectedMapTiles

        ' Select map tiles using the number keys
        For I = 0 To 9
            If KeyPressed(Asc("0") + I) Then
                SelectedMapTiles(I + 1) = SelectedMapTileNumber
            EndIf
        Next

        ' Change modes
        If KeyPressed(Asc("t")) Or KeyPressed(EnterCode) Then _
            Mode = MAP_EDITOR_MODE
    ElseIf Mode = CHARACTER_EDITOR_MODE Then
        If KeyPressed(Asc(" ")) Then
            If SelectedCharacter Then
                ' Unselect currently selected character
                SelectedCharacter = 0
            Else
                ' Attempt to select a character
                SelectedCharacter = CollideCharacters(PlayerX, PlayerY, PLAYER)
            End If
        End If
        If SelectedCharacter Then
            ' Move the selected character around with us, and make them face
            ' the direction of the arrow key we're pressing (if any)
            MoveDirection = GetPlayerMoveDirection(False)
            If MoveDirection >=0 _
                And Not Characters(SelectedCharacter).IsHidden _
                And Characters(SelectedCharacter).Facing <> MoveDirection _
            Then
                ' Just rotate the character
                Characters(SelectedCharacter).StartFacing = MoveDirection
                Characters(SelectedCharacter).Facing = MoveDirection
            Else
                ' Move the character
                If MoveDirection >= 0 Then PrevKeyCode = 0
                HandleEditorArrowKeys
                Characters(SelectedCharacter).StartX = PlayerX
                Characters(SelectedCharacter).StartY = PlayerY
                Characters(SelectedCharacter).X = PlayerX
                Characters(SelectedCharacter).Y = PlayerY
            End If
            ResetCharacterState SelectedCharacter
            ResetCharacterScripts SelectedCharacter
            HandleCharacterAnimation SelectedCharacter
        Else
            HandleEditorArrowKeys
        End If

        RenderMap
        For I = PLAYER + 1 To UBound(Characters)
            RenderCharacter I
        Next
        If KeyPressed(Asc("c")) Or KeyPressed(EnterCode) Then _
            Mode = MAP_EDITOR_MODE
    Else
        Die "Unknown mode: " + Mode
    End If

    ' Render game boy's screen to actual screen
    RenderScreen

    ' While the H key is being held down, show the "help" message
    If _KeyDown(Asc("h")) Then PrintHelp

    ' Display the current mode
    _Dest 0
    Locate 1, 1
    If Mode = CHARACTER_EDITOR_MODE And SelectedCharacter Then
        Print Mode; ": "; Characters(SelectedCharacter).Name
    ElseIf Mode = MAP_RESIZE_MODE Then
        Print Mode; ": "; MapWidth; " x "; MapHeight
    Else
        Print Mode
    End If

    ' Show whatever we've drawn on the screen
    _Display

    ' Make sure the animation doesn't go faster than our intended
    ' frames-per-second (FPS)
    _Limit FPS
Loop Until _KeyDown(EscapeCode) ' Quit if escape key is pressed

System ' Close the program without saying "Press any key..."


' ########################################################################
' # FUNCTION AND SUBROUTINE DEFINITIONS

Function Wrap(Value As Long, MaxValue As Long)
    ' This works so long as Value > -MaxValue
    Wrap = (Value + MaxValue) Mod MaxValue
End Function

Function KeyPressed(KeyCode AS Long)
    KeyPressed = False
    If _KeyDown(KeyCode) Then
        If KeyCode <> PrevKeyCode Then
            KeyPressed = True
            PrevKeyCode = KeyCode
        End If
    Else
        If KeyCode = PrevKeyCode Then
            PrevKeyCode = 0
        End If
    End If
End Function

Sub PrintHelp
    _Dest 0
    Locate 2, 2
    Print "Keyboard controls:"
    Print " H: show this help"
    If Mode = GAME_MODE Then
        Print " Arrow keys: move"
        Print " Z: gameboy's A button"
        Print " X: gameboy's B button"
        Print " C: gameboy's Select button"
        Print " Enter: gameboy's Start button"
        Print " M: enter map editor mode"
        Print " F: change map filename"
        Print " F5: save map"
        Print " F7: load map"
    ElseIf Mode = MAP_EDITOR_MODE Then
        Print " Arrow keys: move"
        Print " 0-9: place tile"
        Print " A: set/unset anchor point"
        Print " T: enter tile selection mode"
        Print " S: enter map scroll mode"
        Print " R: enter map resize mode"
        Print " C: enter character editor mode"
        Print " L: reload images, map tiles, etc"
        Print " M or Enter: exit map editor mode"
    ElseIf Mode = MAP_SCROLL_MODE Then
        Print " Arrow keys: scroll the map"
        Print " C or Enter: exit map scroll mode"
    ElseIf Mode = MAP_RESIZE_MODE Then
        Print " Arrow keys: resize the map"
        Print " R or Enter: exit map resize mode"
    ElseIf Mode = TILE_SELECTOR_MODE Then
        Print " Arrow keys: move"
        Print " 0-9: choose tile"
        Print " T or Enter: exit tile selection mode"
    ElseIf Mode = CHARACTER_EDITOR_MODE Then
        Print " Arrow keys: move"
        Print " Space: select/unselect character"
        Print " C or Enter: exit character editor mode"
    Else
        Die "Unknown mode: " + Mode
    End If
    Print " Escape: quit the program"
End Sub

Sub HandleMapScrollMode
    If KeyPressed(UpCode) Then ScrollMap 0, -1
    If KeyPressed(DownCode) Then ScrollMap 0, 1
    If KeyPressed(LeftCode) Then ScrollMap -1, 0
    If KeyPressed(RightCode) Then ScrollMap 1, 0
End Sub

Sub ScrollMap(AddX As Long, AddY As Long)
    Dim X As Long, Y As Long, X2 As Long, Y2 As Long
    Dim TempValue As Long

    ' The start and end values of the for-loops
    Dim StartX As Long, EndX As Long, StepX As Long
    Dim StartY As Long, EndY As Long, StepY As Long
    If AddX < 0 Then
        StartX = 0
        EndX = MapWidth - 1
        StepX = 1
    Else
        StartX = MapWidth - 1
        EndX = 0
        StepX = -1
    End If
    If AddY < 0 Then
        StartY = 0
        EndY = MapHeight - 1
        StepY = 1
    Else
        StartY = MapHeight - 1
        EndY = 0
        StepY = -1
    End If

    For Y = StartY To EndY + AddY Step StepY
        For X = StartX To EndX + AddX Step StepX
            X2 = Wrap(X + AddX, MapWidth)
            Y2 = Wrap(Y + AddY, MapHeight)
            TempValue = Map(X, Y)
            Map(X, Y) = Map(X2, Y2)
            Map(X2, Y2) = TempValue
        Next
    Next

    RenderMapImage
End Sub

Sub HandleMapResizeMode
    If KeyPressed(UpCode) And MapHeight > 1 Then ResizeMap 0, -1
    If KeyPressed(DownCode) Then ResizeMap 0, 1
    If KeyPressed(LeftCode) And MapWidth > 1 Then ResizeMap -1, 0
    If KeyPressed(RightCode) Then ResizeMap 1, 0
End Sub

Sub ResizeMap(AddX As Long, AddY As Long)
    Dim X As Long, Y As Long

    ' Create a copy of the map's data
    Dim OldMap(0 To MapWidth - 1, 0 To MapHeight - 1) As Long
    For X = 0 To MapWidth - 1
        For Y = 0 To MapHeight - 1
            OldMap(X, Y) = Map(X, Y)
        Next
    Next

    ' Prepare to copy the old data to the new map...
    Dim MinMapWidth As Long, MinMapHeight As Long
    MinMapWidth = MapWidth
    MinMapHeight = MapHeight
    If AddX < 0 Then MinMapWidth = MinMapWidth + AddX
    If AddY < 0 Then MinMapHeight = MinMapHeight + AddY

    ' Actually resize the map
    MapWidth = MapWidth + AddX
    MapHeight = MapHeight + AddY
    ReDim Map(0 To MapWidth - 1, 0 To MapHeight - 1) As Long

    ' Copy the old data to the new map
    For X = 0 To MinMapWidth - 1
        For Y = 0 To MinMapHeight - 1
            Map(X, Y) = OldMap(X, Y)
        Next
    Next

    ' Now create a fresh map image, and render it
    _FreeImage MapImage
    MapImage = _NewImage(MapWidth * MapTileWidth, _
        MapHeight * MapTileHeight, 32)
    RenderMapImage
End Sub

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

Sub WriteAt(X As Long, Y As Long, Width As Long)
    WriteX = X
    WriteY = Y
    WriteStartX = X
    WriteWidth = Width
End Sub

Sub RenderTile( _
    Tileset As Tileset, TileX As Long, TileY As Long, _
    X As Long, Y As Long, ExtraX As Long, ExtraY As Long _
)
    ' NOTE: this subroutine expects _Dest to already be set!..

    Dim TileW As Long, TileH As Long
    TileW = Tileset.TileWidth
    TileH = Tileset.TileHeight

    ' Source and destination (X, Y) coordinates
    Dim SrcX As Long, SrcY As Long
    Dim DstX As Long, DstY As Long

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
    Dim LineWidth As Long
    Dim I As Integer
    Dim Ch As Integer
    Dim Entry As XYPair

    _Dest ScreenImage ' Write to the game boy's screen

    ' Now loop over the characters in the text, and draw each character on
    ' the screen, using the tiles in FontTileset
    For I = 1 To Len(Text)
        Ch = Asc(Mid$(Text, I, 1)) ' Get the next character from Text
        If Ch = Asc("`") Then
            ' Backtick means newline!..
            LineWidth = 0
            WriteX = WriteStartX
            WriteY = WriteY + 1
            _Continue
        ElseIf WriteWidth > 0 And LineWidth >= WriteWidth Then
            LineWidth = 0
            WriteX = WriteStartX
            WriteY = WriteY + 1
        Else
            LineWidth = LineWidth + 1
        End If
        Entry = CharacterMapEntries(FontTileset.CharacterMap, Ch)
        RenderTile FontTileset, Entry.X, Entry.Y, WriteX, WriteY, 0, 0
        WriteX = WriteX + 1
    Next
End Sub

Sub RenderSelectionBox(X As Long, Y As Long, Width As Long, Height As Long)
    ' Draw a box representing the user's selection of something.

    ' White box, closer in
    Line (X - 1, Y - 1)-(X + Width, Y + Height), _RGB(255, 255, 255), B
    ' Black box, further out
    Line (X - 2, Y - 2)-(X + Width + 1, Y + Height + 1), _RGB(0, 0, 0), B
End Sub

Sub ShowMessage(Message As String)
    _Dest 0 ' Print to the screen
    Cls ' Clear the screen
    Locate 1, 1
    Print Message
    _Display ' Show the message
    Sleep ' Wait for a key to be pressed
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

    ' Strip spaces from the beginning of the line
    Dim I As Long
    I = 1
    While Mid$(ParseText, I, 1) = " "
        I = I + 1
    Wend
    ParseText = Mid$(ParseText, I)

    ' Parse the first token right away
    NextToken
End Sub

Sub NextToken
    Dim I As Long, J As Long
    I = Instr(ParseText, " ")
    J = I + 1
    While Mid$(ParseText, J, 1) = " "
        J = J + 1
    Wend
    If I > 0 Then
        Token = Left$(ParseText, I - 1)
        ParseText = Mid$(ParseText, J)
    Else
        Token = ParseText
        ParseText = ""
    End If
    '_Dest = 0: Print "Parsed token: [" + Token + "]"
End Sub

Sub ParseDie(Text As String)
    Die "Don't know what to do with line" + Str$(LineNumber) _
        + ": [" + Text + "]"
End Sub

Sub LoadMapTiles
    Dim Filename As String
    Dim File As Long
    Dim Text As String
    Dim ParsingBottom As Long

    ' E.g. "tilesets/0.txt"
    Filename = "tilesets/" + LTrim$(Str$(MapTilesetNumber)) + ".txt"

    ' Index into MapTiles
    Dim I As Long
    I = 0

    LineNumber = 0

    File = FreeFile
    Open Filename For Input As File
    Do Until Eof(File)
        Line Input #File, Text
        LineNumber = LineNumber + 1
        Dim FirstChar As String
        FirstChar = Left$(Text, 1)
        If Text = "" Or FirstChar = "#" Then
            ' Empty line or comment, ignore it!
        ElseIf Text = "solid" Then
            MapTiles(I).Solidity = SOLID
        ElseIf Text = "jumpup" Then
            MapTiles(I).Solidity = JUMP_UP
        ElseIf Text = "jumpdown" Then
            MapTiles(I).Solidity = JUMP_DOWN
        ElseIf Text = "jumpleft" Then
            MapTiles(I).Solidity = JUMP_LEFT
        ElseIf Text = "jumpright" Then
            MapTiles(I).Solidity = JUMP_RIGHT
        ElseIf Text = "pokemon" Then
            MapTiles(I).HasPokemon = True
        ElseIf ParsingBottom Then
            ' Parsing bottom two tiles of this map tile
            MapTiles(I).BL = Val("&H" + Left$(Text, 2))
            MapTiles(I).BR = Val("&H" + Mid$(Text, 4, 2))
            ParsingBottom = False
            I = I + 1
        ElseIf Instr("0123456789abcdef", FirstChar) Then
            ' Parsing top two tiles of this map tile
            MapTiles(I).TL = Val("&H" + Left$(Text, 2))
            MapTiles(I).TR = Val("&H" + Mid$(Text, 4, 2))
            ParsingBottom = True
        Else
            ParseDie Text
        End If
    Loop
    Close File

    If ParsingBottom Then
        Die "Hit end of file while still parsing map tile" + Str$(I)
    End If

    NumMapTiles = I
End Sub

Sub SaveMap(Filename As String)
    Dim I As Long, X As Long, Y As Long
    Dim File As Long
    File = FreeFile
    Open Filename For Output As File
        Print #File, "tileset "; MapTilesetNumber
        Print #File, ""

        Print #File, "tiles "; MapWidth; " "; MapHeight
        For Y = 0 To MapHeight - 1
            For X = 0 To MapWidth - 1
                If X > 0 Then Print #File, " ";
                Print #File, Hex$(Map(X, Y));
            Next
            Print #File, ""
        Next
        Print #File, ""

        Print #File, "facing "; FacingStr$(Characters(PLAYER).StartFacing)
        Print #File, "position "; _
            Characters(PLAYER).StartX; Characters(PLAYER).StartY
        Print #File, ""

        For I = PLAYER + 1 To UBound(Characters)
            WriteCharacter I, File
            Print #File, ""
        Next
    Close File
End Sub

Sub WriteCharacter(I As Long, File As Long)
    Dim J As Long
    Print #File, "character"
    If Len(Characters(I).Name) Then Print #File, "    name "; _
        Characters(I).Name
    If Characters(I).IsHidden Then Print #File, "    hidden"
    If Characters(I).IsPokemon Then Print #File, "    pokemon"
    Print #File, "    start_y "; Characters(I).TileStartY
    Print #File, "    facing "; FacingStr$(Characters(I).StartFacing)
    Print #File, "    position "; Characters(I).StartX; Characters(I).StartY
    If Characters(I).ScriptStateNumber > 0 Then _
        WriteScript File, _
            Scripts( _
                ScriptStates(Characters(I).ScriptStateNumber).ScriptNumber _
            ), _
            "script"
    For J = 1 To Characters(I).TalkScriptsLength
        WriteScript File, _
            Scripts(Characters(I).TalkScriptsStart + J - 1), _
            "talk"
    Next
    Print #File, "end"
End Sub

Function ParseFacing(Char As String)
    If Char = "u" Then
        ParseFacing = FACING_UP
    ElseIf Char = "d" Then
        ParseFacing = FACING_DOWN
    ElseIf Char = "l" Then
        ParseFacing = FACING_LEFT
    ElseIf Char = "r" Then
        ParseFacing = FACING_RIGHT
    Else
        Die "Can't parse as direction: [" + Char + "]"
    End If
End Function

Function FacingStr$(Facing As Long)
    If Facing = FACING_UP Then FacingStr$ = "u"
    If Facing = FACING_DOWN Then FacingStr$ = "d"
    If Facing = FACING_LEFT Then FacingStr$ = "l"
    If Facing = FACING_RIGHT Then FacingStr$ = "r"
End Function

Sub LoadMap(Filename As String)
    Dim File As Long
    Dim Text As String
    Dim X As Long
    Dim Y As Long

    LineNumber = 0
    MapTilesetNumber = 0

    ReDim _Preserve Characters(1) As Character
    InitializeCharacter PLAYER
    Characters(PLAYER).Name = "PLAYER"

    File = FreeFile
    Open Filename For Input As File
    Do Until Eof(File)
        Line Input #File, Text
        Parse Text
        LineNumber = LineNumber + 1
        If Text = "" Or Left$(Text, 1) = "#" Then
            ' Empty line or comment, ignore it!
        ElseIf Token = "tileset" Then
            NextToken
            MapTilesetNumber = Val(Token)
        ElseIf Token = "tiles" Then
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
                LineNumber = LineNumber + 1
                For X = 0 To MapWidth - 1
                    Map(X, Y) = Val("&H" + Token)
                    NextToken
                Next
            Next

            ' Reset the character: they are now standing in the middle of the map.
            Characters(PLAYER).X = MapWidth / 2
            Characters(PLAYER).Y = MapHeight / 2
        ElseIf Token = "position" Then
            NextToken
            Characters(PLAYER).X = Val(Token)
            NextToken
            Characters(PLAYER).Y = Val(Token)
        ElseIf Token = "facing" Then
            NextToken
            Characters(PLAYER).Facing = ParseFacing(Token)
        ElseIf Token = "character" Then
            ParseCharacter File
        Else
            ParseDie Text
        End If
    Loop
    Close File

    ' Set the "start" versions of various fields of the player character
    Characters(PLAYER).StartX = Characters(PLAYER).X
    Characters(PLAYER).StartY = Characters(PLAYER).Y
    Characters(PLAYER).StartFacing = Characters(PLAYER).Facing

    LoadMapTiles

    RenderMapImage
End Sub

Sub InitializeCharacter(I As Long)
    ' When you ReDim an array of a custom type, its memory has random garbage
    ' in it!.. maybe just when you ReDim Shared?..
    ' Anyway, we need to make sure to zero out all Character fields to each
    ' new element of the Characters array.
    Characters(I).Name = "NONAME"
    Characters(I).IsHidden = False
    Characters(I).IsPokemon = False
    Characters(I).X = 0
    Characters(I).Y = 0
    Characters(I).ExtraX = 0
    Characters(I).ExtraY = 0
    Characters(I).TileStartX = 0
    Characters(I).TileStartY = 0
    Characters(I).TileAddX = 0
    Characters(I).TileAddY = 0
    Characters(I).Facing = FACING_UP
    Characters(I).State = STATE_STANDING
    Characters(I).Frame = 0
    Characters(I).OtherFoot = 0
    Characters(I).ScriptStateNumber = 0
    Characters(I).TalkScriptsStart = 0
    Characters(I).TalkScriptsLength = 0
    Characters(I).TalkScriptNumber = 0
    Characters(I).StartX = 0
    Characters(I).StartY = 0
    Characters(I).StartFacing = FACING_UP
End Sub

Sub ParseCharacter(File As Long)
    Dim I As Long
    Dim Text As String
    I = UBound(Characters) + 1
    ReDim _Preserve Characters(I) As Character
    InitializeCharacter I
    Characters(I).TalkScriptsStart = UBound(Scripts) + 1
    Do
        Line Input #File, Text
        Parse Text
        LineNumber = LineNumber + 1
        If Text = "" Or Left$(Text, 1) = "#" Then
            ' Empty line or comment, ignore it!
        ElseIf Token = "name" Then
            NextToken
            Characters(I).Name = Token
        ElseIf Token = "hidden" Then
            Characters(I).IsHidden = True
        ElseIf Token = "pokemon" Then
            Characters(I).IsPokemon = True
        ElseIf Token = "start_y" Then
            NextToken
            Characters(I).TileStartY = Val(Token)
        ElseIf Token = "facing" Then
            NextToken
            Characters(I).Facing = ParseFacing(Token)
        ElseIf Token = "position" Then
            NextToken
            Characters(I).X = Val(Token)
            NextToken
            Characters(I).Y = Val(Token)
        ElseIf Token = "script" Then
            If Characters(I).TalkScriptsStart < UBound(Scripts) + 1 Then
                Die "Line " + Str$(LineNumber) + ": Can't have " + QUOTE + _
                    "talk" + QUOTE + " before " + QUOTE + "script" + QUOTE
            End If
            ParseScript File, I, True
            Characters(I).TalkScriptsStart = UBound(Scripts) + 1
            ReDim _Preserve ScriptStates(UBound(ScriptStates) + 1) _
                As ScriptState
            SetScriptState UBound(ScriptStates), UBound(Scripts)
            Characters(I).ScriptStateNumber = UBound(ScriptStates)
        ElseIf Token = "talk" Then
            ParseScript File, I, False
        ElseIf Token = "end" Then
            Exit Do
        End If
    Loop

    Characters(I).TalkScriptsLength = _
        UBound(Scripts) + 1 - Characters(I).TalkScriptsStart

    ResetCharacterScripts I

    ' Set the "start" versions of various fields of the player character
    Characters(I).StartX = Characters(I).X
    Characters(I).StartY = Characters(I).Y
    Characters(I).StartFacing = Characters(I).Facing
End Sub

Function WithinMap(X As Long, Y As Long)
    WithinMap = True
    If X < 0 Or X > UBound(Map, 1) Then WithinMap = False
    If Y < 0 Or Y > UBound(Map, 2) Then WithinMap = False
End Function

Function MapSolidityAt(X As Long, Y As Long)
    If Not WithinMap(X, Y) Then
        MapSolidityAt = SOLID
        Exit Function
    End If
    MapSolidityAt = MapTiles(Map(X, Y)).Solidity
End Function

' Whether a character facing the indicated direction can move to the
' indicated map location.
' Returns 1 if character can walk there, 2 if they can jump, 3 if they can
' "short jump", 0 otherwise.
Function CanMoveTo( _
    X As Long, Y As Long, NewX As Long, NewY As Long, _
    MoveDirection As Long, IgnoreCharacter As Long _
)
    Dim Solidity As Long, NewSolidity As Long
    Solidity = MapSolidityAt(X, Y)
    NewSolidity = MapSolidityAt(NewX, NewY)
    CanMoveTo = 0
    If Solidity = JUMP_UP And MoveDirection = FACING_UP Then
        If CollideCharacters(NewX, NewY, IgnoreCharacter) = 0 _
            Then CanMoveTo = 3 ' Can short jump there
    ElseIf _
        NewSolidity = NOT_SOLID Or _
        (NewSolidity = JUMP_UP And MoveDirection <> FACING_DOWN) _
    Then
        If CollideCharacters(NewX, NewY, IgnoreCharacter) = 0 Then _
            CanMoveTo = 1 ' Can walk there
    ElseIf NewSolidity = JUMP_DOWN Then
        NewSolidity = MapSolidityAt(NewX, NewY + 1)
        If MoveDirection = FACING_DOWN And _
            (NewSolidity = NOT_SOLID Or NewSolidity = JUMP_UP) And _
            CollideCharacters(NewX, NewY + 1, IgnoreCharacter) = 0 _
                Then CanMoveTo = 2 ' Can jump there
    ElseIf NewSolidity = JUMP_LEFT Then
        NewSolidity = MapSolidityAt(NewX - 1, NewY)
        If MoveDirection = FACING_LEFT And _
            (NewSolidity = NOT_SOLID Or NewSolidity = JUMP_UP) And _
            CollideCharacters(NewX - 1, NewY, IgnoreCharacter) = 0 _
                Then CanMoveTo = 2 ' Can jump there
    ElseIf NewSolidity = JUMP_RIGHT Then
        NewSolidity = MapSolidityAt(NewX + 1, NewY)
        If MoveDirection = FACING_RIGHT And _
            (NewSolidity = NOT_SOLID Or NewSolidity = JUMP_UP) And _
            CollideCharacters(NewX + 1, NewY, IgnoreCharacter) = 0 _
                Then CanMoveTo = 2 ' Can jump there
    End If
End Function

Function MapHasPokemonAt(X As Long, Y As Long)
    If WithinMap(X, Y) Then
        MapHasPokemonAt = True
        Exit Function
    End If
    MapHasPokemonAt = MapTiles(Map(X, Y)).HasPokemon
End Function

Sub RenderMapTile(X As Long, Y As Long)
    ' Render one of the map's tiles onto the MapImage
    _Dest MapImage
    RenderMapTileAt MapTiles(Map(X, Y)), X, Y, 0, 0
End Sub

' Render a map tile, that is, a 2x2 square of tiles, onto the indicated image
Sub RenderMapTileAt(MapTile As MapTile, X As Long, Y As Long, _
    ExtraX As Long, ExtraY As Long _
)
    ' Number encoding an (X, Y) coordinate into the map's tileset
    Dim XY As Long

    ' Top-left tile
    XY = MapTile.TL
    RenderTile MapTilesets(MapTilesetNumber), _
        XY Mod 16, Int(XY / 16), X * 2 + 0, Y * 2 + 0, _
        ExtraX, ExtraY

    ' Top-right tile
    XY = MapTile.TR
    RenderTile MapTilesets(MapTilesetNumber), _
        XY Mod 16, Int(XY / 16), X * 2 + 1, Y * 2 + 0, _
        ExtraX, ExtraY

    ' Bottom-left tile
    XY = MapTile.BL
    RenderTile MapTilesets(MapTilesetNumber), _
        XY Mod 16, Int(XY / 16), X * 2 + 0, Y * 2 + 1, _
        ExtraX, ExtraY

    ' Bottom-right tile
    XY = MapTile.BR
    RenderTile MapTilesets(MapTilesetNumber), _
        XY Mod 16, Int(XY / 16), X * 2 + 1, Y * 2 + 1, _
        ExtraX, ExtraY
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

Sub UpdateMapEditorAnchorRectangle
    Dim Temp As Long
    RectangleX1 = PlayerX
    RectangleY1 = PlayerY
    If MapEditorAnchorX < 0 Then
        ' Anchor is unset
        RectangleX2 = RectangleX1
        RectangleY2 = RectangleY1
    Else
        ' Anchor is set
        RectangleX2 = MapEditorAnchorX
        RectangleY2 = MapEditorAnchorY
        If RectangleX1 > RectangleX2 Then
            Temp = RectangleX2
            RectangleX2 = RectangleX1
            RectangleX1 = Temp
        End If
        If RectangleY1 > RectangleY2 Then
            Temp = RectangleY2
            RectangleY2 = RectangleY1
            RectangleY1 = Temp
        End If
    End If
End Sub

Sub RenderMap
    ' Copy the map onto the game boy's screen

    ' The location in pixels of the top-left corner of the map on the
    ' game boy's screen
    Dim MapScrollX As Long
    Dim MapScrollY As Long
    MapScrollX = TrueScreenWidth / 2 - PlayerX * MapTileWidth - PlayerExtraX - 8
    MapScrollY = TrueScreenHeight / 2 - PlayerY * MapTileHeight - PlayerExtraY - 8

    _PutImage _
        (MapScrollX, MapScrollY) - ( _
            MapScrollX + MapWidth * MapTileWidth - 1, _
            MapScrollY + MapHeight * MapTileHeight - 1 _
        ), _
        MapImage, ScreenImage

    If Mode = MAP_EDITOR_MODE Then
        ' Render a "selection box" around the tile at the current map location,
        ' or around a larger area selected using the "anchor point"
        _Dest ScreenImage
        UpdateMapEditorAnchorRectangle
        RenderSelectionBox _
            MapScrollX + RectangleX1 * MapTileWidth, _
            MapScrollY + RectangleY1 * MapTileHeight, _
            (RectangleX2 - RectangleX1 + 1) * MapTileWidth, _
            (RectangleY2 - RectangleY1 + 1) * MapTileHeight
    ElseIf Mode = CHARACTER_EDITOR_MODE Then
        ' Render a "selection box" around the tile at the current map location
        _Dest ScreenImage
        RenderSelectionBox _
            MapScrollX + PlayerX * MapTileWidth, _
            MapScrollY + PlayerY * MapTileHeight, _
            MapTileWidth, MapTileHeight
    End If
End Sub

Sub RenderMapTiles
    ' Called when Mode = TILE_SELECTOR_MODE.
    ' Render all map tiles as a grid for the user to select from.

    Dim X As Long, Y As Long, I As Long
    Const ExtraX = 16
    Const ExtraY = 16

    _Dest ScreenImage

    ' Draw all map tiles as a grid
    For I = 0 To NumMapTiles - 1
        X = I Mod TileSelectorWidth
        Y = Int(I / TileSelectorWidth)
        RenderMapTileAt MapTiles(I), X, Y, ExtraX, ExtraY
    Next

    ' Draw a box around the selected map tile
    X = SelectedMapTileNumber Mod TileSelectorWidth
    Y = Int(SelectedMapTileNumber / TileSelectorWidth)
    RenderSelectionBox X * MapTileWidth + ExtraX, Y * MapTileHeight + ExtraY, _
        MapTileWidth, MapTileHeight
End Sub

Sub RenderSelectedMapTiles
    ' Draw the map tiles currently selected for use with number keys 0-9
    ' at the bottom of the screen
    Dim X As Long, Y As Long, I As Long
    Const ExtraX = 0
    Const ExtraY = TrueScreenHeight - MapTileHeight - 1
    _Dest ScreenImage
    For I = 1 To 10
        X = (I - 2 + 10) Mod 10 ' Causes 1 to be at left, 0 at right
        Y = 0
        RenderMapTileAt MapTiles(SelectedMapTiles(I)), _
            X, Y, ExtraX, ExtraY
    Next
    RenderSelectionBox ExtraX, ExtraY, TrueScreenWidth, MapTileHeight
End Sub

Sub RenderTileset
    ' Copy the current tileset onto the game boy's screen

    Dim Tileset As Tileset
    Tileset = MapTilesets(MapTilesetNumber)

    Dim X As Long, Y As Long, Width As Long, Height As Long
    X = 16
    Y = 16
    Width = 16 * TileWidth
    Height = 6 * TileHeight

    _PutImage (X, Y), Tileset.Image, ScreenImage, _
        (Tileset.StartX, Tileset.StartY) - ( _
            Tileset.StartX + Width - 1, _
            Tileset.StartY + Height - 1 _
        )
End Sub

Sub RenderScreen
    ' Draw the game boy's screen (that is, ScreenImage) on the actual screen
    ' (that is, the program's window)
    _PutImage (0, 0)-(ScreenWidth - 1, ScreenHeight - 1), ScreenImage, 0
End Sub

Sub HandleEditorArrowKeys
    If KeyPressed(UpCode) And PlayerY > 0 Then _
        Characters(PLAYER).Y = PlayerY - 1
    If KeyPressed(DownCode) And PlayerY < MapHeight - 1 Then _
        Characters(PLAYER).Y = PlayerY + 1
    If KeyPressed(LeftCode) And PlayerX > 0 Then _
        Characters(PLAYER).X = PlayerX - 1
    If KeyPressed(RightCode) And PlayerX < MapWidth - 1 Then _
        Characters(PLAYER).X = PlayerX + 1
End Sub

Function GetPlayerMoveDirection(Smooth As Integer)
    Dim Up As Long, Down As Long, Left As Long, Right As Long
    If Smooth Then
        Up = _KeyDown(UpCode)
        Down = _KeyDown(DownCode)
        Left = _KeyDown(LeftCode)
        Right = _KeyDown(RightCode)
    Else
        Up = KeyPressed(UpCode)
        Down = KeyPressed(DownCode)
        Left = KeyPressed(LeftCode)
        Right = KeyPressed(RightCode)
    End If
    
    If Up + Down + Left + Right = True Then
        ' Exactly one of the 4 arrow keys is being pressed
        If Up Then GetPlayerMoveDirection = FACING_UP
        If Down Then GetPlayerMoveDirection = FACING_DOWN
        If Left Then GetPlayerMoveDirection = FACING_LEFT
        If Right Then GetPlayerMoveDirection = FACING_RIGHT
    Else
        ' Either no arrow keys were pressed, or more than one were.
        ' In either case, we don't have a specific direction.
        GetPlayerMoveDirection = -1
    End If
End Function

Function FacingAddX(Facing As Long)
    If Facing = FACING_LEFT Then FacingAddX = -1
    If Facing = FACING_RIGHT Then FacingAddX = 1
End Function

Function FacingAddY(Facing As Long)
    If Facing = FACING_UP Then FacingAddY = -1
    If Facing = FACING_DOWN Then FacingAddY = 1
End Function

Sub ResetCharacterState(I As Long)
    Characters(I).State = STATE_STANDING
    Characters(I).Frame = 0
    Characters(I).ExtraX = 0
    Characters(I).ExtraY = 0
End Sub

Sub ResetCharacterScripts(I As Long)
    ' If we have at least one talk script, then the first time someone talks
    ' to us, use the first talk script.
    ' Otherwise (that is, if we have no talk scripts), don't use one!
    If Characters(I).TalkScriptsLength > 0 Then
        Characters(I).TalkScriptNumber = 1
    Else
        Characters(I).TalkScriptNumber = 0
    End If

    If Characters(I).ScriptStateNumber Then
        ResetScriptState Characters(I).ScriptStateNumber
    End If
End Sub

Sub ResetScriptState(I As Long)
    ScriptStates(I).CommandNumber = 1
    ScriptStates(I).Frame = 0
End Sub

Sub HandleCharacterAnimation(I As Long)
    If Characters(I).State = STATE_GONE Then Exit Sub

    Dim State As Long, Facing As Long
    State = Characters(I).State
    Facing = Characters(I).Facing
    If State = STATE_STANDING Then
        If Facing = FACING_DOWN Then Characters(I).TileAddX = 1
        If Facing = FACING_UP Then Characters(I).TileAddX = 4
        If Facing = FACING_LEFT Then Characters(I).TileAddX = 6
        If Facing = FACING_RIGHT Then Characters(I).TileAddX = 8
    ElseIf _
        State = STATE_WALKING Or _
        State = STATE_JUMPING Or _
        State = STATE_SHORT_JUMPING _
    Then
        Dim Frame As Long
        Frame = Characters(I).Frame

        ' Jumping takes twice as long as walking or short jumping
        Dim Multiplier As Long
        Multiplier = 1
        If State = STATE_JUMPING Then Multiplier = 2

        ' Update character's tile of animation... like, which picture
        ' should we draw?.. character facing up with left foot forward?..
        ' that kind of thing.
        Dim Anim As Long
        Anim = Int(Frame / 4 / Multiplier)
        If Facing = FACING_DOWN Then
            Anim = (Anim + Characters(I).OtherFoot * 2) Mod 4
            If Anim = 3 Then Anim = 1
            Characters(I).TileAddX = Anim
        ElseIf Facing = FACING_UP Then
            Anim = (Anim + Characters(I).OtherFoot * 2) Mod 4
            If Anim = 3 Then Anim = 1
            Characters(I).TileAddX = 3 + Anim
        ElseIf Facing = FACING_LEFT Then
            Anim = (Anim + 1) Mod 2
            Characters(I).TileAddX = 6 + Anim
        ElseIf Facing = FACING_RIGHT Then
            Anim = (Anim + 1) Mod 2
            Characters(I).TileAddX = 8 + Anim
        End If

        ' Update character's ExtraX and ExtraY, that is, smoothly move
        ' them from their old map tile towards the one they're walking
        ' onto.
        Const PixelsPerFrame = 2
        Characters(I).ExtraX = _
            (-16 * Multiplier + PixelsPerFrame * (Frame + 1)) _
            * FacingAddX(Facing)
        Characters(I).ExtraY = _
            (-16 * Multiplier + PixelsPerFrame * (Frame + 1)) _
            * FacingAddY(Facing)

        If Frame >= 8 * Multiplier - 1 Then
            ' Done the walking/jumping animation
            ResetCharacterState I
            Characters(I).OtherFoot = (Characters(I).OtherFoot + 1) Mod 2
        Else
            Characters(I).Frame = Frame + 1
        End If
    End If
End Sub

Sub RenderCharacter(I As Long)
    ' Draw the character Characters(I) onto the game boy's screen

    Dim Character As Character
    Character = Characters(I)

    If Character.State = STATE_GONE Then Exit Sub
    If Character.IsHidden And Mode <> CHARACTER_EDITOR_MODE Then Exit Sub

    _Dest ScreenImage

    ' The location in pixels to render the character at
    Dim X As Long
    Dim Y As Long

    ' The location in pixels of the top-left corner of the map on the
    ' game boy's screen
    X = TrueScreenWidth / 2 - PlayerX * MapTileWidth - PlayerExtraX - 8
    Y = TrueScreenHeight / 2 - PlayerY * MapTileHeight - PlayerExtraY - 8

    If Character.IsHidden Then
        ' In "character editor" mode, show an exclamation mark over
        ' hidden characters
        RenderTile MiscCharacterTileset, 2, 1, _
            Character.X, Character.Y, X, Y - 4
        Exit Sub
    End If

    Dim TileX As Long, TileY As Long
    TileX = Character.TileStartX + Character.TileAddX
    TileY = Character.TileStartY + Character.TileAddY

    ' HACK: the bird pokemon character's left and right walking frames are
    ' in the wrong order in the tilesheet, we fix them here...
    If Character.IsPokemon And Character.TileStartY = 0 Then
        If Character.TileAddX >= 6 Then
            TileX = TileX - ((TileX Mod 2) * 2 - 1)
        End If
    End If

    If _
        Character.State = STATE_JUMPING Or _
        Character.State = STATE_SHORT_JUMPING _
    Then
        ' When a character is jumping, we need to render their shadow
        RenderTile MiscCharacterTileset, 9, 0, Character.X, Character.Y, _
            X + Character.ExtraX, Y + Character.ExtraY

        Dim Multiplier As Long
        Multiplier = 1
        If Character.State = STATE_JUMPING Then Multiplier = 2

        ' Character's sprite moves up and down as they jump
        Dim Frame As Long
        Frame = Character.Frame
        Y = Y - (5 * Multiplier - Abs(Frame - 4 * Multiplier))
    End If

    ' Pokemon use a different character tileset
    Dim Tileset As Tileset
    If Character.IsPokemon Then
        Tileset = PokemonCharacterTileset
    Else
        Tileset = CharacterTileset
    End If

    ' Characters are rendered a little bit above the map tile they're
    ' at, which results in a slightly 3d effect, like they're in front
    ' of any walls "behind" them (i.e. the tile directly above them)
    Y = Y - 4

    ' Actually render the character onto the game boy's screen
    RenderTile Tileset, TileX, TileY, Character.X, Character.Y, _
        X + Character.ExtraX, Y + Character.ExtraY
End Sub

Function PlayerX
    PlayerX = Characters(PLAYER).X
End Function

Function PlayerY
    PlayerY = Characters(PLAYER).Y
End Function

Function PlayerExtraX
    PlayerExtraX = Characters(PLAYER).ExtraX
End Function

Function PlayerExtraY
    PlayerExtraY = Characters(PLAYER).ExtraY
End Function

Function AddScriptCommand(CommandType As Integer)
    ReDim _Preserve ScriptCommands(UBound(ScriptCommands) + 1) _
        As ScriptCommand
    ScriptCommands(UBound(ScriptCommands)).CommandType = CommandType
    AddScriptCommand = UBound(ScriptCommands)
End Function

Sub ParseScript(File As Long, CharacterNumber As Long, ShouldLoop As Integer)
    Dim I As Long
    Dim Start As Long
    Dim Script As Script
    Dim Text As String
    Start = UBound(ScriptCommands) + 1
    Do
        Line Input #File, Text
        Parse Text
        LineNumber = LineNumber + 1
        If Text = "" Or Left$(Text, 1) = "#" Then
            ' Empty line or comment, ignore it!
        ElseIf Token = "wait" Then
            I = AddScriptCommand(COMMAND_WAIT)
            NextToken
            ScriptCommands(I).Num1 = Val(Token)
        ElseIf Token = "walk" Then
            I = AddScriptCommand(COMMAND_WALK)
            NextToken
            ScriptCommands(I).Num1 = ParseFacing(Token)
            NextToken
            ScriptCommands(I).Num2 = Val(Token)
        ElseIf Token = "jump" Then
            I = AddScriptCommand(COMMAND_JUMP)
            NextToken
            ScriptCommands(I).Num1 = ParseFacing(Token)
            ScriptCommands(I).Num2 = 1 ' Jump once per command
        ElseIf Token = "face" Then
            I = AddScriptCommand(COMMAND_FACE)
            NextToken
            ScriptCommands(I).Num1 = ParseFacing(Token)
        ElseIf Token = "say" Then
            I = AddScriptCommand(COMMAND_SAY)
            ScriptCommands(I).Str1 = ParseText
        ElseIf Token = "talk" Then
            I = AddScriptCommand(COMMAND_TALK)
            NextToken
            ScriptCommands(I).Num1 = Val(Token)
        ElseIf Token = "end" Then
            Exit Do
        End If
    Loop

    ' Set script fields
    Script.CharacterNumber = CharacterNumber
    Script.Start = Start
    Script.Length = UBound(ScriptCommands) - (Start - 1)
    Script.ShouldLoop = ShouldLoop

    ' Append the new script to the end of the Scripts array
    ReDim _Preserve Scripts(UBound(Scripts) + 1) As Script
    Scripts(UBound(Scripts)) = Script
End Sub

Sub WriteScript(File As Long, Script As Script, ScriptType As String)
    Print #File, "    "; ScriptType
    Dim I As Long
    For I = Script.Start To Script.Start + Script.Length - 1
        Dim Command As ScriptCommand
        Command = ScriptCommands(I)
        If Command.CommandType = COMMAND_WAIT Then
            Print #File, "        wait "; Command.Num1
        ElseIf Command.CommandType = COMMAND_WALK Then
            Print #File, "        walk "; FacingStr$(Command.Num1); Command.Num2
        ElseIf Command.CommandType = COMMAND_JUMP Then
            Print #File, "        jump "; FacingStr$(Command.Num1)
        ElseIf Command.CommandType = COMMAND_FACE Then
            Print #File, "        face "; FacingStr$(Command.Num1)
        ElseIf Command.CommandType = COMMAND_SAY Then
            Print #File, "        say "; Command.Str1
        ElseIf Command.CommandType = COMMAND_TALK Then
            Print #File, "        talk "; Command.Num1
        Else
            Die "Unknown command type: " + Str$(Command.CommandType)
        End If
    Next
    Print #File, "    end"
End Sub

Function CollideCharacters(X As Long, Y As Long, IgnoreCharacter As Long)
    CollideCharacters = False
    Dim I As Long
    For I = 1 To UBound(Characters)
        If I = IgnoreCharacter Then _Continue
        If Characters(I).State = STATE_GONE Then _Continue
        If Characters(I).X = X And Characters(I).Y = Y Then
            CollideCharacters = I
            Exit Function
        End If
    Next
End Function

Sub SetScriptState(StateNumber As Long, ScriptNumber As Long)
    ScriptStates(StateNumber).ScriptNumber = ScriptNumber
    ScriptStates(StateNumber).CommandNumber = 1
    ScriptStates(StateNumber).Frame = 0
End Sub

Function ScriptStateDone(StateNumber As Long)
    Dim State As ScriptState
    State = ScriptStates(StateNumber)
    ScriptStateDone = State.CommandNumber > Scripts(State.ScriptNumber).Length
End Function

Sub UpdateScriptState(StateNumber As Long)
    ' NOTE: caller guarantees Not ScriptStateDone(StateNumber)!

    Dim I As Long
    Dim Script As Script
    Dim State As ScriptState
    Dim Command As ScriptCommand
    State = ScriptStates(StateNumber)
    Script = Scripts(State.ScriptNumber)

    Do
        TalkingText = ""

        If State.CommandNumber > Script.Length Then
            If Script.ShouldLoop Then
                State.CommandNumber = 0
            Else
                Exit Do
            End If
        End If

        ' Wait for character's current animation to complete, before doing
        ' any script actions
        If Characters(Script.CharacterNumber).State <> STATE_STANDING _
            Then Exit Do

        Command = ScriptCommands(Script.Start + State.CommandNumber - 1)
        If Command.CommandType = COMMAND_WAIT Then
            If State.Frame >= Command.Num1 Then
                ' Done waiting!..
                State.CommandNumber = State.CommandNumber + 1
                State.Frame = 0
            Else
                ' Wait some more...
                State.Frame = State.Frame + 1
                Exit Do
            End If
        ElseIf _
            Command.CommandType = COMMAND_WALK Or _
            Command.CommandType = COMMAND_JUMP _
        Then
            If State.Frame >= Command.Num2 Then
                ' Done walking/jumping!..
                State.CommandNumber = State.CommandNumber + 1
                State.Frame = 0
            Else
                ' Walk/jump some more...
                I = Script.CharacterNumber
                Characters(I).Facing = Command.Num1
                Dim Multiplier As Long
                If Command.CommandType = COMMAND_JUMP Then
                    Multiplier = 2
                    Characters(I).State = STATE_JUMPING
                Else
                    Multiplier = 1
                    Characters(I).State = STATE_WALKING
                End If
                Characters(I).X = Characters(I).X + _
                    FacingAddX(Command.Num1) * Multiplier
                Characters(I).Y = Characters(I).Y + _
                    FacingAddY(Command.Num1) * Multiplier
                State.Frame = State.Frame + 1
                Exit Do
            End If
        ElseIf Command.CommandType = COMMAND_FACE Then
            Characters(Script.CharacterNumber).Facing = Command.Num1
            State.CommandNumber = State.CommandNumber + 1
        ElseIf Command.CommandType = COMMAND_SAY Then
            If State.Frame = 1 Then
                ' Done talking
                State.CommandNumber = State.CommandNumber + 1
                State.Frame = 0
            Else
                ' Talk some more...
                TalkingText = Command.Str1
                State.Frame = State.Frame + 1
                Exit Do
            End If
        ElseIf Command.CommandType = COMMAND_TALK Then
            Characters(Script.CharacterNumber).TalkScriptNumber = Command.Num1
            State.CommandNumber = State.CommandNumber + 1
        Else
            Die "Unknown command type: " + Str$(Command.CommandType)
        End If
    Loop

    ' Copy any state updates we've made back into the array of states
    ScriptStates(StateNumber) = State
End Sub

Sub RenderTextBox(Text As String)
    _Dest ScreenImage ' Draw onto the game boy's screen

    Const TextBoxX = 0
    Const TextBoxY = 13
    Const TextBoxWidth = 18
    Const TextBoxHeight = 3
    Dim X As Long, Y As Long

    ' Top-left corner
    RenderTile FontTileset, 9, 9, _
        TextBoxX, TextBoxY, 0, 0

    ' Top-right corner
    RenderTile FontTileset, 11, 9, _
        TextBoxX + TextBoxWidth + 1, TextBoxY, 0, 0

    ' Bottom-left corner
    RenderTile FontTileset, 13, 9, _
        TextBoxX, TextBoxY + TextBoxHeight + 1, 0, 0

    ' Bottom-right corner
    RenderTile FontTileset, 14, 9, _
        TextBoxX + TextBoxWidth + 1, TextBoxY + TextBoxHeight + 1, 0, 0

    For X = 1 To TextBoxWidth
        ' Top edge
        RenderTile FontTileset, 10, 9, _
            TextBoxX + X, TextBoxY, 0, 0
        ' Bottom edge
        RenderTile FontTileset, 10, 9, _
            TextBoxX + X, TextBoxY + TextBoxHeight + 1, 0, 0
    Next

    For Y = 1 To TextBoxHeight
        ' Left edge
        RenderTile FontTileset, 12, 9, _
            TextBoxX, TextBoxY + Y, 0, 0
        ' Bottom edge
        RenderTile FontTileset, 12, 9, _
            TextBoxX + TextBoxWidth + 1, TextBoxY + Y, 0, 0
        ' Fill the middle with emptiness
        For X = 1 To TextBoxWidth
            RenderTile FontTileset, 0, 4, _
                TextBoxX + X, TextBoxY + Y, 0, 0
        Next
    Next

    WriteAt TextBoxX + 1, TextBoxY + 1, TextBoxWidth
    WriteText Text
End Sub
