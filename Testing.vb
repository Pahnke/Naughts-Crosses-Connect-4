'Commented out as not needed when programme is run normally

Module Testing
    'Public Sub Finding_Console_Size()
    '    'When this sub is ran, it will be placed in the playing module but when it is not being used, it is placed here, the testing module, as it does not relate to the main purpose of the  module "playing module"
    '    'This is used to find what size the console needs to be depended on the grid size and game type
    '    'Grid, Width, Height
    '    '1, 54, 16
    '    '2, 54,18
    '    'width constant, +2 on height per row
    '    'Width increase at 13, +4 per column

    '    Dim Width, Height As Integer
    '    Dim Grid_Width, Grid_Height As Integer
    '    Dim Board As New GridInfo
    '    Height = 0
    '    Width = 0
    '    Board.Set_Board_Size(Grid_Width, Grid_Height)
    '    Do
    '        Dim Key_Pressed As ConsoleKey
    '        Key_Pressed = Console.ReadKey(True).Key
    '        Select Case Key_Pressed
    '            Case ConsoleKey.UpArrow
    '                Height -= 1
    '            Case ConsoleKey.DownArrow
    '                Height += 1
    '            Case ConsoleKey.LeftArrow
    '                Width -= 1
    '            Case ConsoleKey.RightArrow
    '                Width += 1
    '            Case ConsoleKey.Enter
    '                Grid_Width += 1
    '                Grid_Height += 1
    '                Board.Set_Board_Size(Grid_Width, Grid_Height)
    '        End Select
    '        Try
    '            Console.SetWindowSize(Width, Height)
    '            Console.SetBufferSize(Width, Height)
    '        Catch
    '            Width += 1
    '            Height += 1
    '        End Try
    '        Console.Clear()
    '        Console.WriteLine()
    '        Board.Draw_Grid()
    '        Playing_Module.Display_Whose_Turn(True, False)
    '        Console.WriteLine()
    '        Playing_Module.Display_How_To_Play(True)
    '        Console.WriteLine("Invalid Option")
    '        Console.WriteLine("The box is already taken")
    '    Loop
    'End Sub
    'Public Sub Set_Up_AI_Test()
    '    'Each difficulty plays the other with difficulties for the default and custom set ups for Noughts and Crosses and Connect Four
    '    'Each difficulty will go first and second against the other difficultys
    '    For Player_One_Diff = 1 To 3
    '        For Player_Two_Diff = 1 To 3
    '            TestAI(2, 2, 2, False, Player_One_Diff, Player_Two_Diff, 100)
    '            TestAI(8, 8, 2, False, Player_One_Diff, Player_Two_Diff, 100)
    '            TestAI(6, 5, 3, True, Player_One_Diff, Player_Two_Diff, 100)
    '            TestAI(6, 7, 4, True, Player_One_Diff, Player_Two_Diff, 100)
    '        Next
    '    Next
    '    'This is to alert me when the testing has been completed
    '    Console.Beep()
    '    Console.WriteLine("Done")
    '    Console.Beep()
    '    Console.ReadLine()
    'End Sub
    'Private Sub TestAI(ByVal Grid_Width As Integer, ByVal Grid_Height As Integer, ByVal No_Together As Integer, ByVal Connect As Boolean, ByVal Diff_One As Integer, ByVal Diff_Two As Integer, ByVal No_Wanted_Games As Integer)
    '    'The AI plays itself and records the number of wins
    '    Dim No_Won(1), No_Games As Integer
    '    Dim Who_Won, No_Moves As Integer
    '    Dim Player_Turn, Game_Over As Boolean
    '    Dim Max_Depth(1) As Integer
    '    Dim Start_Info As New GameInfo
    '    Dim Position As Coordinate
    '    Start_Info.Board = New GridInfo
    '    Start_Info.Set_Connect(Connect)
    '    Start_Info.Set_No_Together(No_Together)
    '    If Diff_One > 1 Then
    '        Max_Depth(0) = Playing_Module.Calculate_Max_Depth(Grid_Width + 1, Grid_Height + 1, Diff_One)
    '    End If
    '    If Diff_Two > 1 Then
    '        Max_Depth(1) = Playing_Module.Calculate_Max_Depth(Grid_Width + 1, Grid_Height + 1, Diff_Two)
    '    End If
    '    Do
    '        No_Moves = 0
    '        Who_Won = 0
    '        Player_Turn = True
    '        Game_Over = False
    '        Start_Info.Board.Set_Board_Size(Grid_Width, Grid_Height)
    '        Do
    '            No_Moves += 1
    '            If Player_Turn = False Then
    '                Start_Info.Set_Difficulty(Diff_Two)
    '                Position = Computers_Move.Computer_Makes_Move(Start_Info, No_Moves, Max_Depth(1))
    '            Else
    '                Start_Info.Set_Difficulty(Diff_One)
    '                Position = Computers_Move.Computer_Makes_Move(Start_Info, No_Moves, Max_Depth(0))
    '            End If
    '            If Connect = True Then
    '                Position.Y = Find_Connect_Y(Grid_Height, Start_Info.Board.Get_Grid, Position.X)
    '            End If
    '            Start_Info.Board.Make_Move_On_Board(Position, Set_Char_Being_Placed(Connect, Player_Turn))
    '            Game_Over = Start_Info.Board.End_Game(Position, No_Together)
    '            If Game_Over = True Then
    '                If Player_Turn = True Then
    '                    Who_Won = 1
    '                Else
    '                    Who_Won = 2
    '                End If
    '            End If
    '            Player_Turn = Not (Player_Turn)
    '        Loop Until Game_Over = True Or No_Moves = ((Grid_Height + 1) * (Grid_Width + 1))
    '        If Who_Won = 1 Then
    '            No_Won(0) += 1
    '        ElseIf Who_Won = 2 Then
    '            No_Won(1) += 1
    '        End If
    '        No_Games += 1
    '    Loop Until No_Games = No_Wanted_Games

    '    'Storing the results in a database
    '    Write_To_Database(Diff_One, Diff_Two, Grid_Width, Grid_Height, Connect, No_Won, No_Wanted_Games, No_Together)
    'End Sub
    'Private Function Find_Connect_Y(ByVal Grid_Height As Integer, ByVal Grid(,) As Char, ByVal X As Integer)
    '    Dim Y As Integer
    '    Dim Empty_Below As Boolean
    '    Do
    '        If Y = Grid_Height Then
    '            Empty_Below = False
    '        ElseIf Grid(X, Y + 1) = vbNullChar Then
    '            Empty_Below = True
    '            Grid(X, Y) = vbNullChar
    '            Y += 1
    '        Else
    '            Empty_Below = False
    '        End If
    '    Loop While Empty_Below = True
    '    Return Y
    'End Function
    'Private Sub Write_To_Database(ByVal Diff_One As Integer, ByVal Diff_Two As Integer, ByVal Grid_Width As Integer, ByVal Grid_Height As Integer, ByVal Connect As Boolean, ByVal No_Won() As Integer, ByVal No_Wanted_Games As Integer, ByVal No_Together As Integer)
    '    Dim New_Connection = New OleDb.OleDbConnection
    '    Dim cmd As OleDb.OleDbCommand
    '    Dim Command_Str As String
    '    Dim Field_Names As String = "Player_One_Difficulty, Player_Two_Difficulty, Grid_Width, Grid_Height, No_Together, Connect_Mode, Player_One_Wins, Player_Two_Wins, Total_Number_Games"
    '    Dim Value_Str As String = Diff_One & ", " & Diff_Two & ", " & Grid_Width & ", " & Grid_Height & ", " & No_Together & ", " & Connect & ", " & No_Won(0) & ", " & No_Won(1) & ", " & No_Wanted_Games
    '    New_Connection.ConnectionString = "Provider = Microsoft.Ace.oledb.12.0; Data Source = AI_Results.accdb"
    '    Try
    '        New_Connection.Open()
    '        Command_Str = "INSERT INTO Results_Table (" & Field_Names & ") VALUES (" & Value_Str & ");"
    '        cmd = New OleDb.OleDbCommand(Command_Str, New_Connection)
    '        cmd.ExecuteNonQuery()
    '    Catch ex As Exception
    '        New_Connection.Close()
    '    End Try
    '    New_Connection.Close()
    'End Sub

End Module
'Private Sub Pretty_Thing(ByVal Grid_Width As Integer, ByVal Grid_Height As Integer)
'    'This is a hidden function which was developed while testing a failed idea but I decided to keep it as it creates interesting patterns
'    Dim Board As New GridInfo
'    Dim Loops As Boolean = True
'    Dim TempInt As Integer
'    Board.Set_Board_Size(2, 2)
'    Change_Window_Size(0, 0)
'    Randomize()
'    TempInt = Math.Floor(Rnd() * 6) + 1
'    Select Case TempInt
'        Case 1
'            Board.Grid = {{"/", "|", "\"}, {"-", "o", "-"}, {"\", "|", "/"}}
'        Case 2
'            Board.Grid = {{"#", "£", "?"}, {"$", "@", "$"}, {"?", "£", "#"}}
'        Case 3
'            Board.Grid = {{"R", "X", "Y"}, {"O", "*", "O"}, {"Y", "X", "R"}}
'        Case 4
'            Board.Grid = {{"R", "R", "R"}, {"Y", "Y", "Y"}, {"R", "R", "R"}}
'        Case 5
'            Board.Grid = {{"1", "4", "7"}, {"2", "5", "8"}, {"3", "6", "9"}}
'        Case 6
'            Board.Grid = {{"Y", "R", "R"}, {"R", "Y", "R"}, {"R", "R", "Y"}}
'    End Select
'    Do
'        Console.Clear()
'        Board.Draw_Grid()
'        System.Threading.Thread.Sleep(350)
'        Board.Grid = RotateGrid(Board)
'        If My.Computer.Keyboard.AltKeyDown = True Then
'            Loops = False
'        End If
'    Loop While Loops = True
'    Change_Window_Size(Grid_Width, Grid_Height)
'End Sub
'Private Function RotateGrid(ByVal Board As GridInfo) As Char(,)
'    Dim Temp(Board.Get_Grid_Width, Board.Get_Grid_Height) As Char
'    For i = 0 To Board.Get_Grid_Width
'        For j = 0 To Board.Get_Grid_Height
'            Temp(i, j) = Board.Grid(Board.Grid_Width - j, i)
'        Next
'    Next
'    Return Temp
'End Function

'Previous Attempt at Minmax algorithm

'Class Branch
'    Private Tree_X, Tree_Y As Integer
'    Private Branch_Score As Long
'    Private ContinueSearch As Boolean
'    Private Tree_Grid(,) As Char
'    Public Lower_Branch As List(Of Branch)
'    Public Sub New(ByVal _Grid_Width As Integer, ByVal _Grid_Height As Integer, ByVal _Grid(,) As Char, Optional ByVal _X As Integer = -1, Optional ByVal _Y As Integer = -1, Optional ByVal CharPlaced As Char = "E")
'        ReDim Tree_Grid(_Grid_Width, _Grid_Height)
'        Tree_Grid = _Grid
'        ContinueSearch = True
'        Tree_X = _X
'        Tree_Y = _Y
'        If Tree_X < _Grid_Width + 1 And Tree_X > -1 Then
'            If Tree_Y < _Grid_Height + 1 And Tree_Y > -1 Then
'                Tree_Grid(Tree_X, Tree_Y) = CharPlaced
'            End If
'        End If
'    End Sub
'    Public Sub ResetTreeValues(ByVal _Grid_Width As Integer, ByVal _Grid_Height As Integer, ByVal _Grid(,) As Char)
'        ReDim Tree_Grid(_Grid_Width, _Grid_Height)
'        Tree_Grid = _Grid
'        ContinueSearch = True
'        Lower_Branch.Clear()
'    End Sub
'    Public Sub Add_BranchScore(ByVal TempNo As Long)
'        Branch_Score += TempNo
'    End Sub
'    Public Function Get_X() As Integer
'        Return Tree_X
'    End Function
'    Public Function Get_Y() As Integer
'        Return Tree_Y
'    End Function
'    Public Function Get_Grid() As Char(,)
'        Return Tree_Grid
'    End Function
'    Public Sub Increase_BranchScore()
'        Branch_Score += 1
'    End Sub
'    Public Sub Decrease_BranchScore()
'        Branch_Score -= 1
'    End Sub
'    Public Function Get_BranchScore() As Long
'        Return Branch_Score
'    End Function
'    Public Function Get_ContinueSearch() As Boolean
'        Return ContinueSearch
'    End Function
'    Public Sub End_Search()
'        ContinueSearch = False
'    End Sub
'End Class
'Sub Proper_Move(ByVal Grid(,) As Char, ByVal Grid_Width As Integer, ByVal Grid_Height As Integer, ByVal Connect As Boolean, ByVal No_Together As Integer, ByRef X As Integer, ByRef Y As Integer, ByVal No_Moves As Integer)
'    Dim CharPlace As Char
'    Dim Tree As New Branch(Grid_Width, Grid_Height, CopyCharArray(Grid))
'    CharPlace = FindCharPlace(Connect)
'    Work_On_This(Grid, Grid_Width, Grid_Height, Connect, No_Together, True, Tree, CharPlace, No_Moves)
'    SearchTree(Tree, True)
'    Find_Best_XY(X, Y, Tree, Connect)
'End Sub
'Function FindCharPlace(ByVal Connect As Boolean) As Char
'    If Connect = False Then
'        Return Nought
'    Else
'        Return Yellow
'    End If
'End Function
'Sub Work_On_This(ByVal Grid(,) As Char, ByVal Grid_Width As Integer, ByVal Grid_Height As Integer, ByVal Connect As Boolean, ByVal No_Together As Integer, ByVal Comp_Turn As Boolean, ByRef _Tree As Branch, ByVal CharPlaced As Char, ByVal Depth As Integer)
'    Dim TempX, TempY, i As Integer
'    Depth += 1
'    _Tree.Lower_Branch = New List(Of Branch)
'    Find_Next_Untaken_Box(Grid, Grid_Width, Grid_Height, TempX, TempY, Connect, True, True) 'Reset i
'    Do While Find_Next_Untaken_Box(Grid, Grid_Width, Grid_Height, TempX, TempY, Connect, True) = True
'        _Tree.Lower_Branch.Add(New Branch(Grid_Width, Grid_Height, CopyCharArray(_Tree.Get_Grid), TempX, TempY, CharPlaced))
'        If CheckRotation(_Tree, _Tree.Lower_Branch(i).Get_Grid, Grid_Width, Grid_Height, Connect, Depth) = False Then
'            If End_Game(_Tree.Lower_Branch(i).Get_Grid, Grid_Width, Grid_Height, No_Together, TempX, TempY) = True Then
'                _Tree.Lower_Branch(i).End_Search()
'                If Comp_Turn = True Then
'                    _Tree.Lower_Branch(i).Increase_BranchScore()
'                Else
'                    _Tree.Lower_Branch(i).Decrease_BranchScore()
'                End If
'            End If
'            i += 1
'        Else
'            _Tree.Lower_Branch.RemoveAt(i)
'        End If
'    Loop
'    CharPlaced = Update_Char_Place(CharPlaced)
'    If (Depth < 4 And Connect = True) Or (Connect = False And Depth < Grid_Width * Grid_Height) Then
'        For i = 0 To _Tree.Lower_Branch.Count - 1
'            If _Tree.Lower_Branch(i).Get_ContinueSearch = True Then
'                Work_On_This(CopyCharArray(_Tree.Lower_Branch(i).Get_Grid), Grid_Width, Grid_Height, Connect, No_Together, Not (Comp_Turn), _Tree.Lower_Branch(i), CharPlaced, Depth)
'            End If
'        Next
'    End If
'End Sub
'Function CopyCharArray(ByVal array(,) As Char) As Char(,)
'    Dim arrayCopy(0 To array.GetLength(0) - 1, 0 To array.GetLength(1) - 1) As Char
'    For i = 0 To array.GetLength(0) - 1
'        For j = 0 To array.GetLength(1) - 1
'            arrayCopy(i, j) = array(i, j)
'        Next
'    Next
'    Return arrayCopy
'End Function
'Function CheckRotation(ByVal Tree As Branch, ByVal TempGrid(,) As Char, ByVal Grid_Width As Integer, ByVal Grid_Height As Integer, ByVal Connect As Boolean, ByVal Depth As Integer) As Boolean
'    If Connect = False And Depth < 3 And Grid_Height = Grid_Width Then
'        Dim AlreadyIn As Boolean = False
'        Dim Temp_Tree_Grid(Grid_Width, Grid_Height) As Char
'        Dim i, Temp_Branch_No As Integer
'        Do
'            TempGrid = RotateGrid(TempGrid, Grid_Width, Grid_Height)
'            Temp_Branch_No = 0
'            Do While Temp_Branch_No < Tree.Lower_Branch.Count And AlreadyIn = False
'                Temp_Tree_Grid = Tree.Lower_Branch(Temp_Branch_No).Get_Grid
'                If GridMatch(TempGrid, Temp_Tree_Grid, Grid_Width, Grid_Height) = True Then
'                    AlreadyIn = True
'                End If
'                Temp_Branch_No += 1
'            Loop
'            i += 1
'        Loop Until AlreadyIn = True Or i = 3
'        Return AlreadyIn
'    Else
'        Return False
'    End If
'End Function
'Function GridMatch(ByVal Grid1(,) As Char, ByVal Grid2(,) As Char, ByVal Grid_Width As Integer, ByVal Grid_Height As Integer) As Boolean
'    Dim Match As Boolean
'    Dim i, x, y As Integer
'    Match = True
'    Do
'        x = i Mod Grid_Width
'        y = i \ Grid_Width
'        If Not Grid1(x, y) = Grid2(x, y) Then
'            Match = False
'        End If
'        i += 1
'    Loop Until Match = False Or i = Grid_Height * Grid_Width
'    Return Match
'End Function
'Sub SearchTree(ByRef Tree As Branch, ByVal CompTurn As Boolean)
'    Dim BestBranchNo As Integer
'    Try
'        If Not Tree.Lower_Branch Is Nothing Then
'            For i = 0 To Tree.Lower_Branch.Count - 1
'                SearchTree(Tree.Lower_Branch(i), Not (CompTurn))
'            Next
'        End If
'    Catch
'    End Try
'    Try
'        Dim BestScore As Integer
'        If CompTurn = True Then
'            BestScore = -Infinity
'            For i = 0 To Tree.Lower_Branch.Count - 1
'                If Tree.Lower_Branch(i).Get_BranchScore > BestScore Then
'                    BestScore = Tree.Lower_Branch(i).Get_BranchScore
'                    BestBranchNo = i
'                End If
'            Next
'        Else
'            BestScore = Infinity
'            For i = 0 To Tree.Lower_Branch.Count - 1
'                If Tree.Lower_Branch(i).Get_BranchScore < BestScore Then
'                    BestScore = Tree.Lower_Branch(i).Get_BranchScore
'                    BestBranchNo = i
'                End If
'            Next
'        End If
'        Tree.Add_BranchScore(Tree.Lower_Branch(BestBranchNo).Get_BranchScore)
'    Catch
'    End Try
'End Sub
'Sub Find_Best_XY(ByRef X As Integer, ByRef Y As Integer, ByVal Tree As Branch, ByVal Connect As Boolean)
'    Dim BestBranchScore, BestBranchNo As Integer
'    BestBranchScore = -Infinity
'    For i = 0 To Tree.Lower_Branch.Count - 1
'        If Tree.Lower_Branch(i).Get_BranchScore > BestBranchScore Then
'            BestBranchNo = i
'            BestBranchScore = Tree.Lower_Branch(i).Get_BranchScore
'        End If
'    Next
'    X = Tree.Lower_Branch(BestBranchNo).Get_X
'    If Connect = False Then
'        Y = Tree.Lower_Branch(BestBranchNo).Get_Y
'    Else
'        Y = 0
'    End If
'End Sub
