Sub Export_Date()

    Sheets("Settings").Select

    Application.DisplayAlerts = False

    ThisWorkbook.Sheets("Settings").Copy
    ActiveWorkbook.SaveAs Filename:="U:\Risk_Model\Risk_Data.csv", FileFormat:=xlCSV, CreateBackup:=True
    ActiveWorkbook.Close

    Application.DisplayAlerts = True
    
    Sheets("Sheet1").Select

End Sub
Sub R_Exec()
    Dim cmd As Object
    Dim rCommand As String, rBin As String, rScript As String
    Dim errorCode As Integer
    Dim waitTillComplete As Boolean: waitTillComplete = True
    Dim debugging As Boolean: debugging = True
    Set cmd = VBA.CreateObject("WScript.Shell")
    
    Dim WS As Worksheet
    Range("B5").Value = Date & " " & Time
    Range("B5").NumberFormat = "dd-mm-yy hh:mm:ss AM/PM"
    Dim ComputerName, UserName As String
    
         
    'Getting user name
    UserName = Environ("username")

    'Assigning value to cell B4
    Range("B4").Value = UserName

    Call Export_Date
    
    Sheets("Settings").Select
    
    Sheets("Sheet1").Select
    
    Set WS = ThisWorkbook.Worksheets("Settings")

    rBin = WS.Range("E2")
    
    rScript = """" & WS.Range("E9") & """"
    
    'Check if the shell has to keep CMD window or not
    If debugging Then
        rCommand = "cmd.exe ""/k"" " & rBin & " " & rScript
    Else
        rCommand = rBin & " " & rScript
    End If

    Debug.Print rCommand 'Print the command for debug

    'Runs R Script and Arguments into process, returning errorCode
    errorCode = cmd.Run(rCommand, vbNormalFocus, waitTillComplete)
    
End Sub
