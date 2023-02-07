Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_NORMAL = 1

Sub RealizarEnvioMasivo()

    Dim rng As Range
    Dim X
    Dim mensaje As String
    
    Application.ScreenUpdating = False
    
    
    For Each rng In Hoja6.Range("Table_1[Nombre]")
        
        mensaje = VBA.Replace("whatsapp://send?phone=" & "505" & rng.Offset(0, 2).Value & "&text=" & "Estimado(a) Sr(a) " & rng.Value & " " & rng.Offset(0, 3).Value & " " & rng.Offset(0, 4).Value & " ", " ", "%20")
        
        Hoja7.Select
        
        
        ActiveSheet.Shapes.Range(Array("Picture 1")).Select
        
        Selection.Copy
        
        X = ShellExecute(hwnd, "Open", mensaje, &O0, &O0, SW_NORMAL)
        
        Call SendKeys("~", True)
        Application.Wait Now + TimeValue("00:00:06")
        Call SendKeys("^v", True)
        
        Application.Wait Now + TimeValue("00:00:06")
        Call SendKeys("~", True)
        
        Application.CutCopyMode = False
        
        Windows(ThisWorkbook2.Name).Activate
        Application.Wait Now + TimeValue("00:00:06")
        
        
    Next rng
    
    Hoja6.Select
    
    Application.ScreenUpdating = True
    MsgBox "Mensajes enviados con exito", vbInformation


End Sub