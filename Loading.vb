Module Loading
    Public Sub File_Selection()
        Dim File_Names() As String
        Dim Selected_File, Decision As Integer
        Dim Return_Loop As Boolean
        Dim FullPath As String
        Dim Temp_File_Name As String = "Temp_File314.txt"
        Console.Title = "File Select"
        Console.Clear()
        'It gets the path of the files
        FullPath = System.IO.Path.GetFullPath(Temp_File_Name)
        FullPath = Left(FullPath, Len(FullPath) - 17)
        'Then it gets the files
        File_Names = Get_Files(FullPath)
        'Changes Window Size
        Console.SetWindowSize(Max_File_Name + 4, 14)
        If File_Names.Count < 11 Then
            Try
                Console.SetBufferSize(Max_File_Name + 4, 14)
            Catch
                Console.SetBufferSize(Max_File_Name + 5, 15)
            End Try
        Else
            'If there is a list of names, it allows the user to scroll through the list
            Console.SetBufferSize(Max_File_Name + 4, File_Names.Count + 3)
        End If
        If File_Names.Count <> 0 Then
            Do
                Return_Loop = False
                Selected_File = Choose_File_Name(File_Names)
                If Selected_File = File_Names.Count Then
                    Main()
                End If
                Decision = File_Decision_Menu(File_Names(Selected_File))
                File_Names(Selected_File) = FullPath & "\" & File_Names(Selected_File) & ".bin"
                Return_Loop = Process_File_Selected(Decision, File_Names(Selected_File))
                File_Names = Get_Files(FullPath)
                If File_Names.Count = 0 Then
                    Console.Clear()
                    Display_No_Files()
                End If
            Loop While Return_Loop = True
        Else
            Display_No_Files()
        End If
    End Sub
    Private Function Get_Files(ByVal FullPath As String) As String()
        Dim File_Names() As String
        'It gets all of the files in that file path
        File_Names = System.IO.Directory.GetFiles(FullPath, "*bin") '*.bin is for a binary file
        'Removes the path and file type and then sorts the names
        File_Names = Remove_FilePath_FileType(File_Names, FullPath)
        File_Names = Sort_Names(File_Names)
        Return File_Names
    End Function
    Private Sub Display_No_Files()
        Console.WriteLine("There are no saved games")
        Console.WriteLine("Please press enter to continue")
        Console.ReadLine()
        Module1.Main()
    End Sub
    Private Sub Load_Game(ByVal File_Name As String)
        Dim Binary_Form_Var As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
        Dim New_Game As PlayingInfo
        Dim IO_Stream_Var As New IO.FileStream(File_Name, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.None)
        Try
            'The game attempts to load the game from the binary file
            New_Game = DirectCast(Binary_Form_Var.Deserialize(IO_Stream_Var), Object)
            Playing_Module.Play(New_Game, True)
        Catch
            Console.WriteLine()
            Console.WriteLine("There was an error loading the game")
            System.Threading.Thread.Sleep(1250)
            File_Selection()
        Finally
            'It needs to be closed, even if there was an error in loading the file
            IO_Stream_Var.Close()
        End Try
    End Sub
    Private Function Remove_Selected_File(ByVal File_List() As String, ByVal File_No As Integer) As String()
        Dim New_File(File_List.Count - 2) As String
        Dim File_Deleted As Boolean
        For i = 0 To File_List.Count - 1
            If i <> File_No And File_Deleted = False Then
                New_File(i) = File_List(i)
            ElseIf i <> File_No And File_Deleted = True Then
                New_File(i - 1) = File_List(i)
            ElseIf i = File_No Then
                File_Deleted = True
            End If
        Next
        Return New_File
    End Function
    Private Function Process_File_Selected(ByVal Decision As Integer, ByRef File_Name As String) As Boolean
        Select Case Decision
            Case 0 'Play
                Load_Game(File_Name)
            Case 1 'Delete
                My.Computer.FileSystem.DeleteFile(File_Name,
                Microsoft.VisualBasic.FileIO.UIOption.AllDialogs,
        Microsoft.VisualBasic.FileIO.RecycleOption.DeletePermanently,
        Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)
                Return True
            Case 2
                Return True
        End Select
        Return False
    End Function
    Private Function File_Decision_Menu(ByVal File_Name As String) As Integer
        Dim Choice As Integer
        Do
            Console.Clear()
            Display_Extra_Space(((Max_File_Name - Len(File_Name)) / 2) + 1)
            Console.WriteLine(File_Name)
            Display_Extra_Space(((Max_File_Name - Len(File_Name)) / 2) + 1)
            Display_File_Decision_Menu_Options(Choice)
        Loop While Update_Menu_Choice(Choice, 2) = False
        Return Choice
    End Function
    Private Sub Display_File_Decision_Menu_Options(ByVal Choice As Integer)
        Console.WriteLine()
        Display_File_Option("Play Game", Choice, 0)
        Console.WriteLine()
        Display_File_Option("Delete File", Choice, 1)
        Console.WriteLine()
        Display_File_Option("Return", Choice, 2)
    End Sub
    Private Function Sort_Names(ByVal Unsorted() As String) As String()
        Dim Temp_String As String
        Dim Sorted_Bool As Boolean
        If Unsorted.Length > 0 Then
            'Bubble sort
            Do
                Sorted_Bool = True
                For i = 0 To Unsorted.Count - 2
                    If Unsorted(i) > Unsorted(i + 1) Then
                        Sorted_Bool = False
                        Temp_String = Unsorted(i)
                        Unsorted(i) = Unsorted(i + 1)
                        Unsorted(i + 1) = Temp_String
                    End If
                Next
            Loop Until Sorted_Bool = True
        End If
        Return Unsorted
    End Function
    Private Function Remove_FilePath_FileType(ByVal File_Names() As String, ByVal File_Path As String) As String()
        For i = 0 To File_Names.Count - 1
            File_Names(i) = Mid(File_Names(i), Len(File_Path) + 2)
            File_Names(i) = Left(File_Names(i), Len(File_Names(i)) - 4)
        Next
        Return File_Names
    End Function
    Private Function Choose_File_Name(ByVal File_Names() As String)
        Dim Choice As Integer
        Do
            'This is to display the file names which could be any length in the same format as the other menus
            Console.Clear()
            Console.WriteLine()
            Display_Extra_Space(((Max_File_Name - 14) / 2) + 1)
            Console.Write("FILE SELECTION")
            Display_Extra_Space(((Max_File_Name - 14) / 2) + 1)
            Console.WriteLine()
            For File_No = 0 To File_Names.Count - 1
                Console.WriteLine()
                Display_File_Option(File_Names(File_No), Choice, File_No)
            Next
            Console.WriteLine()
            Display_File_Option("Return", Choice, File_Names.Count)
        Loop While Update_Menu_Choice(Choice, File_Names.Count) = False
        Return Choice
    End Function
    Private Sub Display_Extra_Space(ByVal Extra_Space As Integer)
        For Blank_Char = 1 To Extra_Space
            Console.Write(" ")
        Next
    End Sub
    Private Sub Display_File_Option(ByVal File_Name As String, ByVal Choice As Integer, ByVal File_No As Integer)
        'Displays a file name with either "-" or "<" around it depending if it is the chosen file or not
        Dim Temp_Length, Extra_Char As Integer
        Temp_Length = Len(File_Name)
        Extra_Char = (Max_File_Name - Temp_Length) \ 2
        Display_Chosen(Choice, File_No, True)
        Display_Extra_Space(Extra_Char)
        Console.Write(File_Name)
        Display_Extra_Space(Extra_Char)
        If Temp_Length Mod 2 <> 0 Then
            Console.Write(" ")
        End If
        Display_Chosen(Choice, File_No, False)
    End Sub
    Private Sub Display_Chosen(ByVal Choice As Integer, ByVal File_No As Integer, ByVal LeftSide As Boolean)
        If Choice = File_No Then
            If LeftSide = True Then
                Console.Write("<")
            Else
                Console.Write(">")
            End If

        Else
            Console.Write("-")
        End If
    End Sub
End Module
