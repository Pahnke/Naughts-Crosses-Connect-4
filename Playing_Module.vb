Module Playing_Module
    Public Const Cross As Char = "X"
    Public Const Nought As Char = "O"
    Public Const Red As Char = "R"
    Public Const Yellow As Char = "Y"
    <Serializable> Class GridInfo
        Private Grid(,) As Char
        Private Grid_Height As Integer
        Private Grid_Width As Integer
        Public Function Check_Taken(ByVal Position As Coordinate) As Boolean
            'If that position has been taken then it returns true
            'If it's empty then it returns false
            Try
                If Grid(Position.X, Position.Y) = vbNullChar Then
                    Return False
                Else
                    Return True
                End If
            Catch
                Return False
            End Try
        End Function
        Public Function Get_Board_Size() As Integer
            Return ((Grid_Width + 1) * (Grid_Height + 1))
        End Function
        Public Function Get_Grid_Width() As Integer
            Return Grid_Width
        End Function
        Public Function Get_Grid_Height() As Integer
            Return Grid_Height
        End Function
        Public Function Get_Grid() As Char(,)
            'When VB copies arrays, it doesn't copy all of the values, it just sets the array which had values copied to as a pointer to the array which was copied
            'This means that if you want to send a separate array, you have to create a new array and manually copy the values across
            Dim New_Grid(Grid_Width, Grid_Height) As Char
            For x = 0 To Grid_Width
                For y = 0 To Grid_Height
                    New_Grid(x, y) = Grid(x, y)
                Next
            Next
            Return New_Grid
        End Function
        Public Sub Set_Entire_Grid(ByVal New_Grid(,) As Char)
            If New_Grid.GetLength(0) - 1 = Grid_Width And New_Grid.GetLength(1) - 1 = Grid_Height Then
                For x = 0 To Grid_Width
                    For y = 0 To Grid_Height
                        Grid(x, y) = New_Grid(x, y)
                    Next
                Next
            End If
        End Sub
        Public Sub Make_Move_On_Board(ByVal Position As Coordinate, ByVal Char_Being_Placed As Char)
            Grid(Position.X, Position.Y) = Char_Being_Placed
        End Sub
        Public Sub Set_Board_Size(ByVal Temp_Grid_Width As Integer, ByVal Temp_Grid_Height As Integer)
            Grid_Width = Temp_Grid_Width
            Grid_Height = Temp_Grid_Height
            ReDim Grid(Grid_Width, Grid_Height)
        End Sub
        Public Function Get_Move_From_Board(ByVal Position As Coordinate) As Char
            Return Grid(Position.X, Position.Y)
        End Function
        Public Function End_Game(ByVal Position As Coordinate, ByVal No_Together As Integer) As Boolean
            'Instead of checking the entire grid, it only checks the last move
            'It checks the row, column, diagonal down to the right, diagonal up to the left of the last move
            If Check_Row(Position, No_Together) = True Then
                Return True
            ElseIf Check_Column(Position, No_Together) = True Then
                Return True
            ElseIf Check_y_x(Position, No_Together) = True Then
                Return True
            ElseIf Check_yx(Position, No_Together) = True Then
                Return True
            End If
            Return False
        End Function
        Private Function Check_Row(ByVal Position As Coordinate, ByVal No_Together As Integer) As Boolean
            Dim Temp_No_Nxt As Integer
            'It goes from left to right in the row
            For i = 0 To Grid_Width - 1
                If Grid(i, Position.Y) = Grid(i + 1, Position.Y) And Grid(i, Position.Y) <> vbNullChar Then
                    Temp_No_Nxt += 1
                    If Temp_No_Nxt = No_Together Then
                        Return True
                    End If
                Else
                    Temp_No_Nxt = 0
                End If
            Next
            Return False
        End Function
        Private Function Check_Column(ByVal Position As Coordinate, ByVal No_Together As Integer) As Boolean
            Dim Temp_No_Nxt As Integer
            'Top of the column to the bottom
            For i = 0 To Grid_Height - 1
                If Grid(Position.X, i) = Grid(Position.X, i + 1) And Grid(Position.X, i) <> vbNullChar Then
                    Temp_No_Nxt += 1
                    If Temp_No_Nxt = No_Together Then
                        Return True
                    End If
                Else
                    Temp_No_Nxt = 0
                End If
            Next
            Return False
        End Function
        Private Function Check_y_x(ByVal Position As Coordinate, ByVal No_Together As Integer) As Boolean
            Dim Temp_No_Nxt, Column, Row As Integer
            'It gets the value most to the left on the diagonal line going up to the left
            'It then scans diagonally down that line
            If Position.X = Position.Y Then
                Row = 0
                Column = 0
            ElseIf Position.X < Position.Y Then
                Column = 0
                Row = Position.Y - Position.X
            ElseIf Position.X > Position.Y Then
                Column = Position.X - Position.Y
                Row = 0
            End If
            Do Until Column = Grid_Width Or Row = Grid_Height
                If Grid(Column, Row) = Grid(Column + 1, Row + 1) And Grid(Column, Row) <> vbNullChar Then
                    Temp_No_Nxt += 1
                    If Temp_No_Nxt = No_Together Then
                        Return True
                    End If
                Else
                    Temp_No_Nxt = 0
                End If
                Row += 1
                Column += 1
            Loop
            Return False
        End Function
        Private Function Check_yx(ByVal Position As Coordinate, ByVal No_Together As Integer) As Boolean
            'It gets the value most to the left on the diagonal line going down to the left
            'It then scans diagonally up that line
            Dim Temp_No_Nxt, Column, Row As Integer
            If (Position.Y + Position.X) > Grid_Width Then
                Row = Position.Y + Position.X - Grid_Width
                Column = Grid_Width
            Else
                Column = Position.X + Position.Y
                Row = 0
            End If
            Do Until Column = 0 Or Row = Grid_Height
                If Grid(Column, Row) = Grid(Column - 1, Row + 1) And Grid(Column, Row) <> vbNullChar Then
                    Temp_No_Nxt += 1
                    If Temp_No_Nxt = No_Together Then
                        Return True
                    End If
                Else
                    Temp_No_Nxt = 0
                End If
                Row += 1
                Column -= 1
            Loop
            Return False
        End Function
        Public Sub Draw_Grid()
            'Top parts of the square
            Display_Board_Border()
            For Row = 0 To Grid_Height
                'The row with the values in
                Console.WriteLine()
                Console.Write("|")
                For Column = 0 To Grid_Width
                    'As the grid is part of this class, the grid is still available in the sub display char as that's also part of the class
                    'But it makes the code look neater if just the char is sent through 
                    'so you don't have a 2d array with two variables for the location in that array sent through and used
                    Display_Char(Grid(Column, Row))
                Next
                Console.WriteLine()
                'The bottom part of each square
                Display_Board_Border()
            Next
            Console.WriteLine()
        End Sub
        Private Sub Display_Board_Border()
            For Border = 0 To Grid_Width
                Console.Write("----")
            Next
            Console.Write("-")
        End Sub
        Private Sub Display_Char(ByVal Temp_Char As Char)
            If Temp_Char = vbNullChar Then
                Console.Write("   ")
            Else
                'Sets Red and Yellow moves to actually be Red and Yellow colour
                If Temp_Char = Red Then
                    Console.ForegroundColor = ConsoleColor.Red
                ElseIf Temp_Char = Yellow Then
                    Console.ForegroundColor = ConsoleColor.Yellow
                End If
                Console.Write(" " & Temp_Char & " ")
                Console.ResetColor()
            End If
            Console.Write("|")
        End Sub
    End Class
    <Serializable> Class GameInfo
        Public Board As GridInfo
        Private Connect As Boolean
        Private No_Together As Integer
        Private One_Player As Boolean
        Private Difficulty As Integer
        Private Who_Went_First As Boolean
        Public Sub Set_Connect(ByVal Temp_Connect As Boolean)
            Connect = Temp_Connect
        End Sub
        Public Function Get_Connect() As Boolean
            Return Connect
        End Function
        Public Sub Set_No_Together(ByVal Temp_No_Together As Integer)
            No_Together = Temp_No_Together
        End Sub
        Public Function Get_No_Together() As Integer
            Return No_Together
        End Function
        Public Sub Set_One_Player(ByVal Temp_One_Player As Boolean)
            One_Player = Temp_One_Player
        End Sub
        Public Function Get_One_Player() As Boolean
            Return One_Player
        End Function
        Public Sub Set_Difficulty(ByVal Temp_Difficulty As Integer)
            Difficulty = Temp_Difficulty
        End Sub
        Public Function Get_Difficulty() As Integer
            Return Difficulty
        End Function
        Public Sub Set_Who_Went_First(ByVal Temp_Who_Went_First As Boolean)
            Who_Went_First = Temp_Who_Went_First
        End Sub
        Public Function Get_Who_Went_First() As Boolean
            Return Who_Went_First
        End Function
    End Class
    <Serializable> Structure Coordinate
        Dim X As Integer
        Dim Y As Integer
    End Structure
    <Serializable> Class PlayingInfo
        Public Set_Up As GameInfo
        Private No_Moves As Integer
        Private Player_Turn As Boolean
        Public Sub Set_Player_Turn(ByVal Temp_Player_Turn As Boolean)
            Player_Turn = Temp_Player_Turn
        End Sub
        Public Function Get_Player_Turn() As Boolean
            Return Player_Turn
        End Function
        Public Sub Update_Player_Turn()
            Player_Turn = Not (Player_Turn)
        End Sub
        Public Function Get_No_Moves() As Integer
            Return No_Moves
        End Function
        Public Sub Increment_No_Moves()
            No_Moves += 1
        End Sub
        Public Sub Decrement_No_Moves()
            No_Moves -= 1
        End Sub
        Public Sub Reset_No_Moves()
            No_Moves = 0
        End Sub
    End Class
    Public Sub Play(ByVal The_Game As PlayingInfo, ByVal Load_Game As Boolean)
        'The same sub is used when playing either game as they are similar games
        'The same logic is used for 1/2 players as the actual game is the same         
        Change_Title(The_Game.Set_Up.Get_Connect)
        Dim Save_Game_Bool, Someone_Won As Boolean
        Dim Who_Won As Integer
        Dim Max_Depth As Integer
        Dim Position As Coordinate
        If Load_Game = False And The_Game.Set_Up.Get_One_Player = True Then
            Dim Temp_Player_Turn As Boolean
            'This is to decide the difficulty and who goes first
            The_Game.Set_Up.Set_Difficulty(Set_Up_One_Player(Temp_Player_Turn, The_Game.Set_Up.Get_Connect))
            The_Game.Set_Player_Turn(Temp_Player_Turn)
            'Who's going first is stored so the same person goes first in the next game is the game is replayed
            The_Game.Set_Up.Set_Who_Went_First(The_Game.Get_Player_Turn)
        ElseIf The_Game.Set_Up.Get_One_Player = True And Load_Game = True Then
            'If it's a load_game so one from a file or someone plays again, then the difficulty and who's goes first was already selected so the game loads it
            The_Game.Set_Player_Turn(The_Game.Get_Player_Turn)
        ElseIf Load_Game = False Then
            'This is for 2 player so player 1 always goes first and the 2 users choose between them who's first 
            The_Game.Set_Player_Turn(True)
        End If
        'If it's 2 player, the difficulty is 0, if it's easy, the difficulty is 1 so this is only calculate for medium and hard which are the 2 difficulties which it's needed
        If The_Game.Set_Up.Get_Difficulty = 2 Or The_Game.Set_Up.Get_Difficulty = 3 Then
            Max_Depth = Calculate_Max_Depth(The_Game.Set_Up.Board.Get_Grid_Width + 1, The_Game.Set_Up.Board.Get_Grid_Height + 1, The_Game.Set_Up.Get_Difficulty)
        End If
        'This is to set the original value of the char else it would be placing a vbNUllChar
        Change_Console_Size(The_Game.Set_Up.Board.Get_Grid_Width, The_Game.Set_Up.Board.Get_Grid_Height, The_Game.Set_Up.Get_Connect)
        'Who_Won starts as 0 which is a draw
        Who_Won = 0
        Do
            The_Game.Increment_No_Moves()
            If The_Game.Get_Player_Turn = False Then
                If The_Game.Set_Up.Get_One_Player = True Then
                    'As the computer can take a couple of seconds to make a move, this shows the user what is happening and not that the game has frozen
                    Display_Calculating_Message(The_Game.Set_Up.Board)
                    Position = Computers_Move.Computer_Makes_Move(The_Game.Set_Up, The_Game.Get_No_Moves, Max_Depth)
                Else
                    'The user chooses if they want to save the game when inputing their move by pressing 's'
                    Save_Game_Bool = Recieve_Coord_From_User(The_Game.Set_Up.Board, Position, The_Game.Set_Up.Get_Connect, False)
                End If
            Else
                Save_Game_Bool = Recieve_Coord_From_User(The_Game.Set_Up.Board, Position, The_Game.Set_Up.Get_Connect, True)
            End If
            If Save_Game_Bool = True Then
                'As a move wasn't made this turn, the number of moves needs to be decreased so when they continue playing, the number of moves is at the correct value
                The_Game.Decrement_No_Moves()
                Saving.Save_Game(The_Game)
            End If
            'The move which is inputted is then played and it checks if the game has ended
            Make_Move(The_Game.Set_Up.Board, Position, The_Game.Set_Up.Get_Connect, True, Computers_Move.Set_Char_Being_Placed(The_Game.Set_Up.Get_Connect, The_Game.Get_Player_Turn))
            Someone_Won = The_Game.Set_Up.Board.End_Game(Position, The_Game.Set_Up.Get_No_Together)
            If Someone_Won = True Then
                If The_Game.Get_Player_Turn = True Then
                    Who_Won = 1
                Else
                    Who_Won = 2
                End If
            End If
            'This then makes it the next persons turn
            The_Game.Update_Player_Turn()
        Loop Until Someone_Won = True Or The_Game.Get_No_Moves = The_Game.Set_Up.Board.Get_Board_Size
        'This displays who won and allows the user to decide what they want to do next, eg. replay or leave
        After_Game(The_Game, Who_Won)
    End Sub
    Private Sub Display_Calculating_Message(ByVal Board As GridInfo)
        Console.Clear()
        Board.Draw_Grid()
        Console.WriteLine()
        Console.WriteLine("Computer is Calculating")
    End Sub
    Private Sub After_Game(ByVal The_Game As PlayingInfo, ByVal Who_Won As Integer)
        Dim Key_Pressed As ConsoleKey
        Do
            Display_Winner(The_Game.Set_Up.Board, Who_Won, The_Game.Set_Up.Get_One_Player, The_Game.Set_Up.Get_Connect)
            Console.WriteLine("Press P to play again")
            Console.WriteLine("Press Escape to leave")
            Key_Pressed = Console.ReadKey(True).Key
            If Key_Pressed = ConsoleKey.P Then
                The_Game = Reset_Game(The_Game)
                Play(The_Game, True)
            ElseIf Key_Pressed = ConsoleKey.Escape Then
                Module1.Main()
            End If
        Loop
    End Sub
    Private Sub Change_Title(ByVal Connect As Boolean)
        If Connect = True Then
            Console.Title = "Connect 4"
        Else
            Console.Title = "XvO"
        End If
    End Sub
    Private Sub Change_Console_Size(ByVal Grid_Width As Integer, ByVal Grid_Height As Integer, ByVal Connect As Boolean)
        'It changes the window size depended on the grid size for aesthetics and practicality as a small grid on a large console looks bad and a large grid in a small console isn't properly displayed 
        Dim Width, Height As Integer
        If Grid_Width < 13 Then
            Width = 54
        Else
            Width = 54 + (Grid_Width - 12) * 4
        End If
        If Connect = True Then
            Height = 13 + 2 * Grid_Height
        Else
            Height = 11 + 2 * Grid_Height
        End If
        Console.SetWindowSize(Width, Height)
        Try
            Console.SetBufferSize(Width, Height)
        Catch
            Console.SetBufferSize(Width + 1, Height + 1)
        End Try
    End Sub
    Private Function Set_Up_One_Player(ByRef Player_Turn As Boolean, ByVal Connect As Boolean)
        Dim Loop_Back As Boolean
        Dim Difficulty As Integer
        'The if statement and the loop is to allow the return function to work while in the menus
        Do
            Difficulty = Choose_Difficulty()
            If Difficulty = 4 Then
                Module1.Game_Mode_Menu(Connect)
            End If
            Player_Turn = User_Chooses_Who_Goes_First(Loop_Back)
        Loop While Loop_Back = True
        Return Difficulty
    End Function
    Private Function Choose_Difficulty() As Integer
        Dim Difficulty As Integer
        'It chooses the difficulty when one player
        'It works the same as the menus earlier on
        Do
            Console.Clear()
            Console.WriteLine("   Choose Difficulty")
            If Difficulty = 0 Then
                Console.WriteLine("     <  Easy  >")
            Else
                Console.WriteLine("     -  Easy  -")
            End If
            If Difficulty = 1 Then
                Console.WriteLine("     < Medium >")
            Else
                Console.WriteLine("     - Medium -")
            End If
            If Difficulty = 2 Then
                Console.WriteLine("     <  Hard  >")
            Else
                Console.WriteLine("     -  Hard  -")
            End If
            If Difficulty = 3 Then
                Console.WriteLine("     < Return >")
            Else
                Console.WriteLine("     - Return -")
            End If
        Loop While Update_Menu_Choice(Difficulty, 3) = False
        Return Difficulty + 1
    End Function
    Private Function User_Chooses_Who_Goes_First(ByRef Continue_Loop As Boolean) As Boolean
        'It's used in One player to decide if the player goes first or second
        'Continue Loop is sent through to allow the return function to work
        Dim Choice As Integer
        Continue_Loop = False
        Choice = 0
        Do
            Console.Clear()
            Console.WriteLine("   Do you want to go    first or second?")
            Console.WriteLine()
            If Choice = 0 Then
                Console.WriteLine("     < First  >")
            Else
                Console.WriteLine("     - First  -")
            End If
            If Choice = 1 Then
                Console.WriteLine("     < Second >")
            Else
                Console.WriteLine("     - Second -")
            End If
            If Choice = 2 Then
                Console.WriteLine("     < Return >")
            Else
                Console.WriteLine("     - Return -")
            End If
        Loop While Update_Menu_Choice(Choice, 2) = False
        If Choice = 0 Then
            Return True
        ElseIf Choice = 1 Then
            Return False
        ElseIf Choice = 2 Then
            Continue_Loop = True
        End If
        Return True
    End Function
    Private Function Calculate_Max_Depth(ByVal Grid_Width As Integer, ByVal Grid_Height As Integer, ByVal Difficulty As Integer) As Integer
        Dim Max_Depth As Long
        Try
            Max_Depth = Math.Ceiling((Grid_Width * Grid_Height) - (Math.Floor(((((Factorial(Grid_Width)) ^ (Grid_Height)) / (Grid_Height) ^ (Grid_Width))) / 1670000000000000)))
        Catch
            'If the grid size is really big, it will cause an overflow error
            'And if the grid size is really big then its depth can only be 2 else it would take too long calculate each move
            Max_Depth = 2
        End Try
        'Sometimes the depth is negative but it only happens with larger grid sizes so it is set to 2
        If Max_Depth < 2 Then
            Max_Depth = 2
        End If
        'Medium difficulty will be "half" as hard as hard difficulty
        If Difficulty = 2 Then
            Max_Depth = Math.Floor(0.5 * (Max_Depth))
        End If
        Return Max_Depth
    End Function
    Private Function Factorial(ByVal Fact_No As Integer) As Integer
        Dim Temp_No As Long = 1
        For i = 1 To Fact_No
            Temp_No = Temp_No * i
        Next
        Return Temp_No
    End Function
    Private Function Recieve_Coord_From_User(ByVal Board As GridInfo, ByRef Position As Coordinate, ByVal Connect As Boolean, ByVal P1_Turn As Boolean) As Boolean
        Dim Old_Position As Coordinate
        Dim Temp_Char As Char
        Dim Key_Pressed As ConsoleKey
        'The question mark will appear in the next available box, the function was originally designed for the minimax algorithm but it helps improve the user experience here
        Computers_Move.Find_Next_Untaken_Box(Board, Position, Connect, False)
        'The question mark may be moved into a box which a move has already been made
        'The Old_Position is where the question mark was and store what was there so when the question mark moves
        'The previous move can be placed back in the box 
        If Connect = False Then
            Old_Position.X = Position.X
            Old_Position.Y = Position.Y
            Temp_Char = Board.Get_Move_From_Board(Position)
        End If
        Do
            Console.Clear()
            If Connect = False Then
                'The postion of the question mark is recorded so once it moves, the value which was there is returned
                Board.Make_Move_On_Board(Old_Position, Temp_Char)
                Temp_Char = Board.Get_Move_From_Board(Position)
                Board.Make_Move_On_Board(Position, "?")
                Old_Position.X = Position.X
                Old_Position.Y = Position.Y
            Else
                Display_Current_Choice_Connect(Position.X, Board.Get_Grid_Width)
            End If
            Board.Draw_Grid()
            Display_Whose_Turn(P1_Turn, Connect)
            Console.WriteLine()
            Display_How_To_Play(Connect)
            Key_Pressed = Console.ReadKey(True).Key
            'It records the key pressed so it knows the direction to move the question mark or to confirm the move with enter or go to the menu if escape is pressed
            Select Case Key_Pressed
                'This is used to leave the game
                Case ConsoleKey.Escape
                    Main()
                    'The up/down/left/right moves the question mark in that direction 
                    'up/down doesn't work in connect 4 as you can only go left or right when choosing a move
                Case ConsoleKey.UpArrow
                    If Connect = False Then
                        Position.Y -= 1
                        If Position.Y < 0 Then
                            Position.Y = Board.Get_Grid_Height
                        End If
                    End If
                Case ConsoleKey.DownArrow
                    If Connect = False Then
                        Position.Y += 1
                        If Position.Y > Board.Get_Grid_Height Then
                            Position.Y = 0
                        End If
                    End If
                Case ConsoleKey.LeftArrow
                    Position.X -= 1
                    If Position.X < 0 Then
                        Position.X = Board.Get_Grid_Width
                    End If
                Case ConsoleKey.RightArrow
                    Position.X += 1
                    If Position.X > Board.Get_Grid_Width Then
                        Position.X = 0
                    End If
                Case ConsoleKey.Enter
                    If Connect = False Then
                        Board.Make_Move_On_Board(Old_Position, Temp_Char)
                    End If
                    If Board.Get_Move_From_Board(Position) <> vbNullChar Then
                        Console.WriteLine()
                        Console.WriteLine("Invalid Option - The box is already taken")
                        System.Threading.Thread.Sleep(1500)
                    End If
                Case ConsoleKey.S
                    If Connect = False Then
                        Board.Make_Move_On_Board(Old_Position, Temp_Char)
                    End If
                    'True is to say that it will save the game
                    Return True
            End Select
        Loop Until Key_Pressed = ConsoleKey.Enter And Board.Get_Move_From_Board(Position) = vbNullChar
        Return False
    End Function
    Private Sub Display_Winner(ByVal Board As GridInfo, ByVal Who_Won As Integer, ByVal One_Player As Boolean, ByVal Connect As Boolean)
        'It displays the board, who won and the piece which the winner was using eg "R" or "X"
        Console.Clear()
        Board.Draw_Grid()
        Select Case Who_Won
            Case 0
                Console.WriteLine("Draw!")
            Case 1
                Display_Whose_Turn(True, Connect)
                Console.Write(" - Won")
            Case 2
                If One_Player = True Then
                    Display_Whose_Turn(False, Connect, True)
                Else
                    Display_Whose_Turn(False, Connect)
                End If
                Console.Write(" - Won")
        End Select
        Console.WriteLine()
    End Sub
    Private Sub Display_Current_Choice_Connect(ByVal X As Integer, ByVal Grid_Width As Integer)
        'It displays the question mark over the column which is currently selected
        For i = 0 To Grid_Width
            If i = X Then
                Console.Write("  ? ")
            Else
                Console.Write("    ")
            End If
        Next
        Console.WriteLine()
    End Sub
    Private Sub Display_Whose_Turn(ByVal P1_Turn As Boolean, ByVal Connect As Boolean, Optional ByVal Computer_Won As Boolean = False)
        'It displays whose turn it is
        'It is also used to show the piece the computer was using (player 2's piece) at the end of the game is the computer has won
        If P1_Turn = True Then
            Console.Write("Player 1 - ")
            If Connect = True Then
                Console.ForegroundColor = ConsoleColor.Red
                Console.Write(Red)
                Console.ResetColor()
            Else
                Console.Write(Cross)
            End If
        Else
            If Computer_Won = False Then
                Console.Write("Player 2 - ")
            Else
                Console.Write("Computer - ")
            End If
            If Connect = True Then
                Console.ForegroundColor = ConsoleColor.Yellow
                Console.Write(Yellow)
                Console.ResetColor()
            Else
                Console.Write(Nought)
            End If
        End If
    End Sub
    Private Sub Display_How_To_Play(ByVal Connect As Boolean)
        'Displays the controls used to play
        Console.WriteLine("Use the arrows keys to move the question mark")
        Console.WriteLine("Press enter to confirm your selection")
        Console.WriteLine("If you want to save the game then press 'S'")
        Console.WriteLine("If you would like to leave the game then press escape")
        If Connect = True Then
            If My.Computer.Keyboard.CapsLock = False Then
                Console.WriteLine("Turn Caps Lock ON to turn OFF the dropping animation")
            Else
                Console.WriteLine("Turn Caps Lock OFF to turn ON the dropping animation")
            End If
        End If
    End Sub
    Private Sub Make_Move(ByRef Board As GridInfo, ByRef Position As Coordinate, ByVal Connect As Boolean, ByVal P1_Turn As Boolean, ByVal Char_Being_Placed As Char)
        'In Noughts and Crosses, the piece is just placed in that position
        'In connect 4, it is "dropped" into position so it is placed into a box and if the box below it is empty then it is placed in the box below it
        'If the connect animation is on then it drops it box by box, else it goes straight to the bottom
        If Connect = False Then
            Board.Make_Move_On_Board(Position, Char_Being_Placed)
        Else
            Dim Empty_Below As Boolean = True
            Dim Connect_Animation As Boolean = True
            Do
                Connect_Animation = Not (My.Computer.Keyboard.CapsLock)
                If Connect_Animation = True Then
                    Board.Make_Move_On_Board(Position, Char_Being_Placed)
                    Console.Clear()
                    Console.WriteLine()
                    Board.Draw_Grid()
                    System.Threading.Thread.Sleep(175)
                End If
                If Position.Y = Board.Get_Grid_Height Then
                    Position.Y += 1
                    Empty_Below = False
                Else
                    Position.Y += 1
                    If Board.Get_Move_From_Board(Position) = vbNullChar Then
                        Empty_Below = True
                        Position.Y -= 1
                        Board.Make_Move_On_Board(Position, vbNullChar)
                        Position.Y += 1
                    Else
                        Empty_Below = False
                    End If
                End If
            Loop While Empty_Below = True
            Position.Y -= 1
            Board.Make_Move_On_Board(Position, Char_Being_Placed)
        End If
    End Sub
    Private Function Reset_Game(ByVal The_Game As PlayingInfo) As PlayingInfo
        'Everything is set to their original value
        Dim New_Game As New PlayingInfo
        New_Game.Set_Up = New GameInfo
        New_Game.Set_Up.Board = New GridInfo
        With New_Game
            .Set_Up = The_Game.Set_Up
            .Reset_No_Moves()
            .Set_Up.Board.Set_Board_Size(The_Game.Set_Up.Board.Get_Grid_Width, The_Game.Set_Up.Board.Get_Grid_Height)
        End With
        If The_Game.Set_Up.Get_One_Player = True Then
            New_Game.Set_Player_Turn(The_Game.Set_Up.Get_Who_Went_First)
        Else
            New_Game.Set_Player_Turn(True)
        End If
        Return New_Game
    End Function
End Module