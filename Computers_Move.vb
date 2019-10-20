Module Computers_Move
    Private Const Infinity As Integer = 141421 'It's just a large value which is greater than the possible value of a leaf node
    Private Class Move
        Private Value_of_Move As Integer
        Private Prune As AlphaBeta
        Public Next_Move As List(Of Move)
        Public Sub New(ByVal Temp_Prune As AlphaBeta)
            Prune.Alpha_Value = Temp_Prune.Alpha_Value
            Prune.Beta_Value = Temp_Prune.Beta_Value
        End Sub
        Public Function Check_Prune(ByVal Comp_Move As Boolean) As Boolean
            'Value of move is currently the best move avaliable to the branch
            'If it's the computers turn, then the move is the highest score but if it's the players turn, then it's the lowest score
            If Comp_Move = True Then
                Prune.Alpha_Value = Value_of_Move
            Else
                Prune.Beta_Value = Value_of_Move
            End If
            'If the value of the maximiser is greater than the minimiser value, then it can be pruned            
            If Prune.Beta_Value < Prune.Alpha_Value Then
                'True means that it can be pruned
                Return True
            End If
            'False means that it can't be pruned
            Return False
        End Function
        Public Function Get_Prune()
            Return Prune
        End Function
        Public Sub Set_Value_of_Move(ByVal Temp_Value As Integer)
            Value_of_Move = Temp_Value
        End Sub
        Public Function Get_Branch_Value()
            Return Value_of_Move
        End Function
    End Class
    Structure Calc_Move_Struc
        Dim Board As GridInfo
        Dim No_Together As Integer
        Dim Connect As Boolean
        Dim No_Moves As Integer
        Dim Max_Depth As Integer
    End Structure
    Structure AI_Struc
        Dim Calc_Move As Calc_Move_Struc
        Dim Depth As Integer
        Dim No_Remaining_Moves As Integer
        Dim Comp_Move As Boolean
    End Structure
    Structure AlphaBeta
        Dim Alpha_Value As Integer 'Maximiser's Value
        Dim Beta_Value As Integer 'Minimizer's Value
    End Structure
    Public Function Computer_Makes_Move(ByVal Start_Info As GameInfo, ByVal No_Moves As Integer, ByVal Max_Depth As Integer) As Coordinate
        Dim Position As Coordinate
        Select Case Start_Info.Get_Difficulty
            Case 1   'Easy
                'Easy is just random
                Position = Rand_Comp_Move(Start_Info.Board, Start_Info.Get_Connect)
            Case 2, 3  'Medium, Hard
                'Medium and hard are calculated but to different depths
                'Its depth is how many moves ahead it looks
                Dim Info_Required_From_Game As Calc_Move_Struc
                Info_Required_From_Game.Board = New GridInfo
                Info_Required_From_Game.Board = Start_Info.Board
                Info_Required_From_Game.Connect = Start_Info.Get_Connect
                Info_Required_From_Game.Max_Depth = Max_Depth
                Info_Required_From_Game.No_Moves = No_Moves
                Info_Required_From_Game.No_Together = Start_Info.Get_No_Together
                Position = Calc_Comp_Move(Info_Required_From_Game)
        End Select
        Return Position
    End Function
    Private Function Rand_Comp_Move(ByVal Board As GridInfo, ByVal Connect As Boolean)
        Dim Position As Coordinate
        'It loops until it finds a valid move
        Do
            Randomize()
            Position.X = Math.Floor(Rnd() * (Board.Get_Grid_Width + 1))
            If Connect = False Then
                Randomize()
                Position.Y = Math.Floor(Rnd() * (Board.Get_Grid_Height + 1))
            Else
                Position.Y = 0
            End If
        Loop While Board.Check_Taken(Position) = True
        Return Position
    End Function
    Private Function Calc_Comp_Move(ByVal Info_Required_From_Game As Calc_Move_Struc) As Coordinate
        Dim Info_Needed_Minimax As AI_Struc
        Dim Prune As AlphaBeta
        Dim Position As Coordinate
        'With 2d arrays, VB does not send them byval, it just sends a pointer to the array
        'Board error is set up so if the minimax algorithm crashes mid search, the board can be reset to the value which was sent through so the board in the game doesn't become the board from the minimax algorithm
        Dim Board_Error(Info_Required_From_Game.Board.Get_Grid_Width, Info_Required_From_Game.Board.Get_Grid_Height) As Char
        Prune.Alpha_Value = -Infinity
        Prune.Beta_Value = Infinity
        'This represents the last move the player made and is the root of tree
        Dim Players_Last_Move As New Move(Prune)
        'This sets up the parameter which is passed into the minimax algorithm to begin with
        Info_Needed_Minimax.Calc_Move.Board = New GridInfo
        Info_Needed_Minimax.Calc_Move = Info_Required_From_Game
        Info_Needed_Minimax.Comp_Move = True
        Info_Needed_Minimax.Depth = 0
        'This sets the number of remaining moves
        If Info_Required_From_Game.Connect = False Then
            Info_Needed_Minimax.No_Remaining_Moves = (Info_Required_From_Game.Board.Get_Grid_Width + 1) * (Info_Required_From_Game.Board.Get_Grid_Height + 1) - Info_Required_From_Game.No_Moves + 1
        Else
            'For connect 4, it goes column by column and checks if the top row is empty, if it is then it counts as an avaliable move
            Info_Needed_Minimax.No_Remaining_Moves = Remaining_Connect_Moves(Info_Required_From_Game.Board)
        End If
        'If the programme crashes during the minimax search, the programme should continue playing which means that it will just generate a random move
        Try
            'It will calculate values for the next moves
            Board_Error = Info_Required_From_Game.Board.Get_Grid
            MiniMax(Players_Last_Move, Info_Needed_Minimax)
            'It will then look at those values and decide which one it wants
            Find_Best_Move(Players_Last_Move.Next_Move, Position, Info_Required_From_Game.Board, Info_Required_From_Game.Connect)
            'In connect four, piece is dropped so the value is set to 0 so it can be dropped into position
        Catch
            'This is only here incase the programme crashes so the user can still continue playing the game
            'Hopefully the minimax algorithm won't crash though so this function won't be required here
            'The board is reset to the value of the board from the game and not from the minimax algorithm
            Info_Required_From_Game.Board.Set_Entire_Grid(Board_Error)
            Return Rand_Comp_Move(Info_Required_From_Game.Board, Info_Required_From_Game.Connect)
        End Try
        If Info_Required_From_Game.Connect = True Then
            Position.Y = 0
        End If
        Return Position
    End Function
    Private Function Remaining_Connect_Moves(ByVal Board As GridInfo) As Integer
        'It connect 4 the no of moves is the grid width unless columns are full so this checks if a column is full and if it isn't then it means that a move can be played there
        Dim No_Remaining_Moves As Integer
        Dim Temp_Position As Coordinate
        Temp_Position.Y = 0
        For Temp_Position.X = 0 To Board.Get_Grid_Width
            If Board.Get_Move_From_Board(Temp_Position) = vbNullChar Then
                No_Remaining_Moves += 1
            End If
        Next
        Return No_Remaining_Moves
    End Function
    Private Sub MiniMax(ByRef Previous_Move As Move, ByVal Info_Needed_Recursion As AI_Struc)
        Dim Temp_Value_To_Be_Separated_To_XY As Integer
        Dim Temp_Position As Coordinate
        Info_Needed_Recursion.Depth += 1
        Previous_Move.Next_Move = New List(Of Move)
        'Initalise the value depending on comp move to +- infinity
        Previous_Move = Set_Defualt_MinMax_Value(Previous_Move, Not (Info_Needed_Recursion.Comp_Move))
        'Goes through the next possible moves and see if they end the game
        For i = 1 To Info_Needed_Recursion.No_Remaining_Moves
            'It gets the next available move and adds it to the list of moves
            Find_Next_Untaken_Box(Info_Needed_Recursion.Calc_Move.Board, Temp_Position, Info_Needed_Recursion.Calc_Move.Connect, True, Temp_Value_To_Be_Separated_To_XY)
            Previous_Move.Next_Move.Add(New Move(Previous_Move.Get_Prune))
            'It then makes the newly found move
            Info_Needed_Recursion.Calc_Move.Board.Make_Move_On_Board(Temp_Position, Set_Char_Being_Placed(Info_Needed_Recursion.Calc_Move.Connect, Not (Info_Needed_Recursion.Comp_Move)))
            'It checks if the game has ended, so if the move was a winning move
            If Info_Needed_Recursion.Calc_Move.Board.End_Game(Temp_Position, Info_Needed_Recursion.Calc_Move.No_Together) = True Then
                'If the game has ended then the value of the move is set
                If Info_Needed_Recursion.Comp_Move = True Then
                    'Computer wins then it's Grid size+1 - depth as depth can't be greater than grid size +1 in means it's always positive and the less moves to get there means it's a higher score
                    Previous_Move.Next_Move(i - 1).Set_Value_of_Move(1 + Info_Needed_Recursion.Calc_Move.Board.Get_Board_Size - Info_Needed_Recursion.Depth)
                Else
                    'Computer loses then it's depth - grid size -1 so it's always negative as gridsize + 1 has to be greater than depth and it means that the more moves it takes to lose is less negative so a better score
                    Previous_Move.Next_Move(i - 1).Set_Value_of_Move(Info_Needed_Recursion.Depth - Info_Needed_Recursion.Calc_Move.Board.Get_Board_Size - 1)
                End If
                'It only continues if the depth hasn't reached the limit yet
            ElseIf Info_Needed_Recursion.Depth < Info_Needed_Recursion.Calc_Move.Max_Depth Then
                'For noughts and crosses, the number of remaining moves is reduced by 1
                'For connect 4, the top row is counted to see if that column can be played in
                If Info_Needed_Recursion.Calc_Move.Connect = False Then
                    Info_Needed_Recursion.No_Remaining_Moves -= 1
                Else
                    Info_Needed_Recursion.No_Remaining_Moves = Remaining_Connect_Moves(Info_Needed_Recursion.Calc_Move.Board)
                End If
                'If there is a move to play, then the algorithm continues to go deeper 
                If Info_Needed_Recursion.No_Remaining_Moves > 0 Then
                    'The next turn will be played by the opposite player so the piece being played and the boolean for who's turn it is needs updating
                    Info_Needed_Recursion.Comp_Move = Not (Info_Needed_Recursion.Comp_Move)
                    MiniMax(Previous_Move.Next_Move(i - 1), Info_Needed_Recursion)
                    'It needs to be changed back so the correct moves can be made when considering the sibling branches of the branch which was just expanded
                    Info_Needed_Recursion.Comp_Move = Not (Info_Needed_Recursion.Comp_Move)
                End If
                If Info_Needed_Recursion.Calc_Move.Connect = False Then
                    Info_Needed_Recursion.No_Remaining_Moves += 1
                Else
                    Info_Needed_Recursion.No_Remaining_Moves = Remaining_Connect_Moves(Info_Needed_Recursion.Calc_Move.Board)
                End If
            End If
            Info_Needed_Recursion.Calc_Move.Board.Make_Move_On_Board(Temp_Position, vbNullChar)
            'Above is the undoing of making the move to allow the search to continue for sibling branches and parent branches sibling branches
            'Below is the setting of the branch value, so if the child branch has a new score which is better than the current score of the branch then it equals that
            If Info_Needed_Recursion.Comp_Move = True Then
                'Computers move means that it wants the highest score
                If Previous_Move.Next_Move(i - 1).Get_Branch_Value > Previous_Move.Get_Branch_Value Then
                    Previous_Move.Set_Value_of_Move(Previous_Move.Next_Move(i - 1).Get_Branch_Value)
                End If
            Else
                'Players turn so it wants the lowest score
                If Previous_Move.Next_Move(i - 1).Get_Branch_Value < Previous_Move.Get_Branch_Value Then
                    Previous_Move.Set_Value_of_Move(Previous_Move.Next_Move(i - 1).Get_Branch_Value)
                End If
            End If
            'It then checks if any branches can be pruned so if it doesn't have to explore any branches and leave the loop 
            If Previous_Move.Check_Prune(Info_Needed_Recursion.Comp_Move) = True Then
                Exit For
            End If
        Next
    End Sub
    Private Sub Find_Best_Move(ByVal Possibile_Moves As List(Of Move), ByRef Position As Coordinate, ByVal Board As GridInfo, ByVal Connect As Boolean)
        Dim Temp_Best_Score As Integer = -Infinity
        Dim Temp_Best_Move_No As New List(Of Integer)
        Dim Chosen_Move_No, Temp_I As Integer
        'It creates a list of moves which have the highest scores
        For Current_Move = 0 To Possibile_Moves.Count - 1
            If Possibile_Moves(Current_Move).Get_Branch_Value > Temp_Best_Score Then
                Temp_Best_Score = Possibile_Moves(Current_Move).Get_Branch_Value
                Temp_Best_Move_No.Clear()
                Temp_Best_Move_No.Add(Current_Move)
            ElseIf Possibile_Moves(Current_Move).Get_Branch_Value = Temp_Best_Score Then
                Temp_Best_Move_No.Add(Current_Move)
            End If
        Next
        'If there's only one move in the list, it chooses that move else it randomly chooses from the list of the highest scoring moves
        If Temp_Best_Move_No.Count = 1 Then
            Chosen_Move_No = Temp_Best_Move_No(0)
        Else
            Randomize()
            Chosen_Move_No = Temp_Best_Move_No(Math.Floor(Temp_Best_Move_No.Count * Rnd()))
        End If
        'It then finds the coordinates of that available move as it finds the next move in a chronological order so running it n times is the position of the nth branch
        'By recalculating it here, it means that the position doesn't have to be stored during the calculation
        For No_Moves_Made = 0 To Chosen_Move_No
            Find_Next_Untaken_Box(Board, Position, Connect, True, Temp_I)
        Next
    End Sub
    Private Function Set_Defualt_MinMax_Value(ByVal Temp_Move As Move, ByVal Comp_Move As Boolean) As Move
        'Comp_Move refers to move that's about to be played so the opposite to what Temp_Move is
        'So if Comp_Move is True then Temp_Move is a player's move so will be trying to minimize
        'As it's trying to miminimze then it wants the lowest score possibile so to begin with 
        'it 's given the highest score possible to work towards a lower score
        'The opposite goes for when Comp_move is False so it's maximising and is after the highest possible move
        'so it starts with lowest possible score and then tries to get a higher score
        If Comp_Move = False Then
            Temp_Move.Set_Value_of_Move(-Infinity)
        Else
            Temp_Move.Set_Value_of_Move(Infinity)
        End If
        Return Temp_Move
    End Function
    Public Function Find_Next_Untaken_Box(ByVal Board As GridInfo, ByRef Position As Coordinate, ByVal Connect As Boolean, ByVal AI_Mode As Boolean, Optional ByRef Value_To_Be_Separated_Into_XY As Integer = 0) As Boolean
        Dim Empty_Below, Found_Box As Boolean
        Found_Box = True
        If Connect = False Then
            'It goes from the top left to the bottom right row by row until it finds an empty box
            Do
                Position.X = Value_To_Be_Separated_Into_XY Mod (Board.Get_Grid_Width + 1)
                Position.Y = Value_To_Be_Separated_Into_XY \ (Board.Get_Grid_Width + 1)
                Value_To_Be_Separated_Into_XY += 1
            Loop While Board.Check_Taken(Position) = True And Not Value_To_Be_Separated_Into_XY > ((Board.Get_Grid_Width + 1) * (Board.Get_Grid_Height + 1))
            If Value_To_Be_Separated_Into_XY > ((Board.Get_Grid_Width + 1) * (Board.Get_Grid_Height + 1)) Then
                Found_Box = False
            End If
        Else
            'It just scans the top box as it will drop into position if it is an available move
            Position.Y = 0
            Do
                Position.X = Value_To_Be_Separated_Into_XY Mod (Board.Get_Grid_Width + 1)
                Value_To_Be_Separated_Into_XY += 1
            Loop While Board.Check_Taken(Position) = True And Not Value_To_Be_Separated_Into_XY > Board.Get_Grid_Width + 1
            'When the computer makes a move in needs the x,y values where the piece will end up so this is used to calculate it
            'When this is used to show the next available move for the user, it only needs to be the top value as it will drop down when that column is selected
            If AI_Mode = True Then
                Position.Y = -1
                Do
                    Position.Y += 1
                    If Position.Y = Board.Get_Grid_Height Then
                        Empty_Below = False
                    Else
                        Position.Y += 1
                        Empty_Below = Not (Board.Check_Taken(Position))
                        Position.Y -= 1
                    End If
                Loop While Empty_Below = True
            End If
            If Value_To_Be_Separated_Into_XY > Board.Get_Grid_Width + 1 Then
                Found_Box = False
            End If
        End If
        Return Found_Box
    End Function
    Public Function Set_Char_Being_Placed(ByVal Connect As Boolean, ByVal Player_Turn As Boolean) As Char
        'This is used to set the original char being placed
        If Connect = True Then
            If Player_Turn = True Then
                Return Red
            Else
                Return Yellow
            End If
        Else
            If Player_Turn = True Then
                Return Cross
            Else
                Return Nought
            End If
        End If
    End Function
End Module