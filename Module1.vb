Module Module1
    'This contains the menus for deciding which game to play and on what board size
    Public Sub Main()
        Console.Title = "Game"
        Dim Choice As Integer
        Try
            Console.SetWindowSize(21, 8)
            Console.SetBufferSize(21, 8)
        Catch
            Console.SetBufferSize(22, 9)
        End Try
        'Chooses which game
        Do
            Display_Main_Menu_Options(Choice)
        Loop While Update_Menu_Choice(Choice, 3) = False
        Select Case Choice
            Case 0
                Game_Mode_Menu(False)
            Case 1
                Game_Mode_Menu(True)
            Case 2
                Loading.File_Selection()
            Case 3
                End
        End Select
    End Sub
    Private Sub Display_Main_Menu_Options(ByVal Choice As Integer)
        Console.Clear()
        Console.WriteLine("    CHOOSE A GAME")
        Console.WriteLine()
        If Choice = 0 Then
            Console.WriteLine("< Noughts & Crosses>")
        Else
            Console.WriteLine("- Noughts & Crosses-")
        End If
        If Choice = 1 Then
            Console.WriteLine("<    Connect 4     >")
        Else
            Console.WriteLine("-    Connect 4     -")
        End If
        If Choice = 2 Then
            Console.WriteLine("<     Load File    >")
        Else
            Console.WriteLine("-     Load File    -")
        End If
        If Choice = 3 Then
            Console.WriteLine("<       Exit       >")
        Else
            Console.WriteLine("-       Exit       -")
        End If
    End Sub
    Public Function Update_Menu_Choice(ByRef Choice As Integer, ByVal Limit As Integer) As Boolean
        Dim KeyPressed As ConsoleKey
        KeyPressed = Console.ReadKey(True).Key
        'Records the key pressed and updates the menu accordingly 
        'It can loop over the top
        'As the bottom option is the return option, by pressing escape, the bottom option is selected so it acts as a return function
        Select Case KeyPressed
            Case ConsoleKey.UpArrow
                Choice -= 1
                If Choice < 0 Then
                    Choice = Limit
                End If
            Case ConsoleKey.DownArrow
                Choice += 1
                If Choice > Limit Then
                    Choice = 0
                End If
            Case ConsoleKey.Enter
                Return True
            Case ConsoleKey.Escape
                Choice = Limit
                Return True
        End Select
        Return False
    End Function
    Public Sub Game_Mode_Menu(ByVal Connect As Boolean)
        Dim Choice As Integer
        Choice = 0
        'Chooses 1/2 player or Custom board
        Do
            Console.Clear()
            Console.WriteLine("       Game Mode")
            Display_One_Two_Player_Options(Choice)
            If Choice = 2 Then
                Console.WriteLine("    <   Custom   >")
            Else
                Console.WriteLine("    -   Custom   -")
            End If
            If Choice = 3 Then
                Console.WriteLine("    <   Return   >")
            Else
                Console.WriteLine("    -   Return   -")
            End If
        Loop While Update_Menu_Choice(Choice, 3) = False
        Process_Menu_Choice(Choice, Connect)
    End Sub
    Private Sub Display_One_Two_Player_Options(ByVal Choice As Integer)
        If Choice = 0 Then
            Console.WriteLine("    < One Player >")
        Else
            Console.WriteLine("    - One Player -")
        End If
        If Choice = 1 Then
            Console.WriteLine("    < Two Player >")
        Else
            Console.WriteLine("    - Two Player -")
        End If
    End Sub
    Private Sub Process_Menu_Choice(ByVal Decision As Integer, ByVal Connect As Boolean)
        Dim Start_Info As New PlayingInfo
        'This sub decides which subs to call depending on the users inputs
        Start_Info.Set_Up = New GameInfo
        Start_Info.Set_Up.Board = New GridInfo
        'The width and height of the grid and the number in a row required for a win is 1 less than their normal value as the programme counts from 0 not 1
        If Connect = True Then
            Start_Info.Set_Up.Board.Set_Board_Size(6, 5)
            Start_Info.Set_Up.Set_No_Together(3)
        Else
            Start_Info.Set_Up.Board.Set_Board_Size(2, 2)
            Start_Info.Set_Up.Set_No_Together(2)
        End If
        Start_Info.Set_Up.Set_Connect(Connect)
        Select Case Decision
            Case 0
                Start_Info.Set_Up.Set_One_Player(True)
                Playing_Module.Play(Start_Info, False)
            Case 1
                Start_Info.Set_Up.Set_One_Player(False)
                Playing_Module.Play(Start_Info, False)
            Case 2
                Custom_Set_Up(Connect)
            Case 3
                Main()
            Case Else
                ErrorMessage()
                Main()
        End Select
    End Sub
    Private Sub Custom_Set_Up(ByVal Connect As Boolean)
        Dim Max_Grid_Width, Max_Grid_Height As Integer
        Dim Choice As Integer
        Dim Start_Info As New PlayingInfo
        Dim Temp_Width, Temp_Height As Integer
        Start_Info.Set_Up = New GameInfo
        Start_Info.Set_Up.Board = New GridInfo
        Find_Max_Grid_Size(Max_Grid_Width, Max_Grid_Height, Connect)
        'The grid width & height are inputted and the number in a row required
        'Each of them have expection handling for invalid data types inputted (not a number) and for being in the right range so a positive grid size
        'Both the grid width and grid height have to be greater than 2
        'The No_Together has to be greater than 1 or the game would end after the first turn 
        Temp_Width = Input_Width(Max_Grid_Width)
        Temp_Height = Input_Height(Max_Grid_Height, Temp_Width)
        Start_Info.Set_Up.Set_No_Together(Input_No_Together(Temp_Width, Temp_Height) - 1)
        Choice = Input_One_Player(Temp_Width, Temp_Height, Start_Info.Set_Up.Get_No_Together + 1)
        Temp_Width -= 1
        Temp_Height -= 1
        Start_Info.Set_Up.Set_Connect(Connect)
        Start_Info.Set_Up.Board.Set_Board_Size(Temp_Width, Temp_Height)
        If Choice = 0 Then
            Start_Info.Set_Up.Set_One_Player(True)
        ElseIf Choice = 1 Then
            Start_Info.Set_Up.Set_One_Player(False)
        Else
            Game_Mode_Menu(Connect)
        End If
        Playing_Module.Play(Start_Info, False)
    End Sub
    Private Sub Find_Max_Grid_Size(ByRef Width As Integer, ByRef Height As Integer, ByVal Connect As Boolean)
        'So the grid isn't larger than the screen
        Width = Math.Floor(((Console.LargestWindowWidth - 54) / 4) + 12)
        If Connect = True Then
            Height = Math.Floor((Console.LargestWindowHeight - 13) / 2)
        Else
            Height = Math.Floor((Console.LargestWindowHeight - 11) / 2)
        End If
    End Sub
    Private Function Input_Width(ByVal Max_Grid_Width As Integer) As Integer
        Dim Grid_Width As Integer
        Dim Valid_Input As Boolean
        Do
            Console.Clear()
            Try
                Console.WriteLine("Max: " & Max_Grid_Width)
                Console.WriteLine("Min: 2")
                Console.Write("Grid Width: ")
                Grid_Width = Console.ReadLine
                Valid_Input = True
                If Grid_Width <= 1 Or Grid_Width > Max_Grid_Width Then
                    ErrorMessage()
                    Valid_Input = False
                End If
            Catch
                ErrorMessage()
                Valid_Input = False
            End Try
        Loop While Valid_Input = False
        Return Grid_Width
    End Function
    Private Function Input_Height(ByVal Max_Grid_Height As Integer, ByVal Grid_Width As Integer) As Integer
        Dim Grid_Height As Integer
        Dim Valid_Input As Boolean
        Do
            Console.Clear()
            Console.WriteLine("Grid Width: " & Grid_Width)
            Console.WriteLine("Min: 2")
            Console.WriteLine("Max: " & Max_Grid_Height)
            Try
                Console.Write("Grid Height: ")
                Grid_Height = Console.ReadLine
                Valid_Input = True
                If Grid_Height <= 1 Or Grid_Height > Max_Grid_Height Then
                    ErrorMessage()
                    Valid_Input = False
                End If
            Catch ex As Exception
                ErrorMessage()
                Valid_Input = False
            End Try
        Loop While Valid_Input = False
        Return Grid_Height
    End Function
    Private Function Input_No_Together(ByVal Grid_Width As Integer, ByVal Grid_Height As Integer) As Integer
        Dim No_Together As Integer
        Dim Valid_Input As Boolean
        Do
            Console.Clear()
            Console.WriteLine("Grid Width: " & Grid_Width)
            Console.WriteLine("Grid Height: " & Grid_Height)
            Console.WriteLine("Min: 2")
            If Grid_Height > Grid_Width Then
                Console.WriteLine("Max " & Grid_Height)
            Else
                Console.WriteLine("Max " & Grid_Width)
            End If
            Try
                Console.Write("No. in a row required to win: ")
                No_Together = Console.ReadLine
                Valid_Input = True
                'The No. Together can't be 1 as the game just ends after the first move
                'The No. Together also can't be larger than both the grid width and the grid height else it would be impossible to achieve
                If (No_Together > Grid_Width And No_Together > Grid_Height) Or No_Together < 2 Then
                    ErrorMessage()
                    Valid_Input = False
                End If
            Catch
                ErrorMessage()
                Valid_Input = False
            End Try
        Loop While Valid_Input = False
        Return No_Together
    End Function
    Private Function Input_One_Player(ByVal Grid_Width As Integer, ByVal Grid_Height As Integer, ByVal No_Together As Integer) As Integer
        Dim Choice As Integer
        Do
            Console.Clear()
            Console.WriteLine("Grid Width: " & Grid_Width)
            Console.WriteLine("Grid Height: " & Grid_Height)
            Console.WriteLine("No. Together: " & No_Together)
            Console.WriteLine()
            Display_One_Two_Player_Options(Choice)
            If Choice = 2 Then
                Console.WriteLine("    <   Return   >")
            Else
                Console.WriteLine("    -   Return   -")
            End If
        Loop While Update_Menu_Choice(Choice, 2) = False
        Return Choice
    End Function
    Private Sub ErrorMessage()
        'Error
        Console.WriteLine("Invalid Option - Please Try Again")
        System.Threading.Thread.Sleep(1500)
    End Sub
End Module
