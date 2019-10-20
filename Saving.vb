Module Saving
    Public Const Max_File_Name As Integer = 48

    Public Sub Save_Game(ByRef The_Game As PlayingInfo)
        Dim Decision As Integer
        Dim File_Name As String
        'Decides if the user wants to save it and then what happens after they save it
        Decision = Verify_Saving()
        Select Case Decision
            Case 0, 1
                File_Name = Get_File_Name()
                Console.Clear()
                Console.WriteLine("Game is being saved...")
                Serialise_Game(The_Game, File_Name)
                If Decision = 0 Then
                    Module1.Main()
                Else
                    Playing_Module.Play(The_Game, True)
                End If
            Case 2
                Playing_Module.Play(The_Game, True)
        End Select
    End Sub
    Private Function Get_File_Name() As String
        Dim File_Name As String
        Dim Valid_Name As Boolean
        Console.SetWindowSize(52, 14)
        Try
            Console.SetBufferSize(52, 14)
        Catch
            Console.SetBufferSize(53, 15)
        End Try
        Do
            Console.Clear()
            'Tries to get the file name and it validates the name
            Try
                Console.WriteLine("Please enter the file name (Max Length: {0})", (Max_File_Name - 4))
                Console.WriteLine("No special characters are allowed:")
                File_Name = Console.ReadLine
                Valid_Name = Validate_File_Name(File_Name)
            Catch
                Valid_Name = False
            End Try
        Loop While Valid_Name = False
        File_Name = File_Name & ".bin"
        Return File_Name
    End Function
    Private Function Validate_File_Name(ByVal File_Name As String) As Boolean
        Dim Reg_Exp As New System.Text.RegularExpressions.Regex("^[A-Za-z0-9 ]+$")
        'Checks it's not too long
        If Len(File_Name) > Max_File_Name - 4 Then
            Console.WriteLine("Name is too long")
            System.Threading.Thread.Sleep(1500)
            Return False
        End If
        'Checks it against the regular expression
        If Reg_Exp.IsMatch(File_Name) = False Then
            If File_Name = "" Then
                Console.WriteLine("Name can't be nothing")
                System.Threading.Thread.Sleep(1500)
            Else
                Console.WriteLine("No special characters are allowed in the file name")
                System.Threading.Thread.Sleep(2000)
            End If
            Return False
        End If
        Return True
    End Function
    Private Function Verify_Saving() As Integer
        Dim Choice As Integer
        Try
            Console.SetWindowSize(19, 7)
            Console.SetBufferSize(19, 7)
        Catch
            Console.SetBufferSize(20, 8)
        End Try
        Do
            Console.Clear()
            Console.WriteLine("     SAVE GAME  ")
            Console.WriteLine()
            If Choice = 0 Then
                Console.WriteLine("  < Save & Exit >")
            Else
                Console.WriteLine("  - Save & Exit -")
            End If
            If Choice = 1 Then
                Console.WriteLine("  < Save & Play >")
            Else
                Console.WriteLine("  - Save & Play -")
            End If
            If Choice = 2 Then
                Console.WriteLine("  <   Return    >")
            Else
                Console.WriteLine("  -   Return    -")
            End If
        Loop While Module1.Update_Menu_Choice(Choice, 2) = False
        Return Choice
    End Function
    Private Sub Serialise_Game(ByVal The_Game As PlayingInfo, ByVal File_Name As String)
        Dim Var_File_Stream As New System.IO.FileStream(File_Name, System.IO.FileMode.Create)
        Dim Binary_Form_Var As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
        Try
            'The game attempts to write the game to a binary file
            Binary_Form_Var.Serialize(Var_File_Stream, The_Game)
        Catch
            Console.WriteLine()
            Console.WriteLine("There was an error saving the game")
            System.Threading.Thread.Sleep(1250)
            Save_Game(The_Game)
        Finally
            'It needs to be closed if the file was successfully saved or not
            Var_File_Stream.Close()
        End Try
    End Sub
End Module