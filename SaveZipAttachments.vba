Public Sub SaveZipAttachments()
    On Error GoTo ErrorHandler
    
    Dim olApp As Object
    Dim olNamespace As Object
    Dim olAccount As Object
    Dim olRootFolder As Object
    Dim olInboxFolder As Object
    Dim olFolder As Object
    Dim olItem As Object
    Dim olAttachment As Object
    Dim saveFolder As String
    Dim fileName As String
    Dim fileExtension As String
    Dim attachmentsSaved As Integer
    Dim startDate As Date
    Dim endDate As Date
    
      ' Solicitar fechas al usuario
    startDate = InputBox("Ingrese la fecha de inicio (dd/mm/yyyy):", "Fecha de inicio", Format(Date - 4, "dd/mm/yyyy"))
    If startDate = 0 Then Exit Sub ' El usuario canceló
    
    endDate = InputBox("Ingrese la fecha de fin (dd/mm/yyyy):", "Fecha de fin", Format(Date, "dd/mm/yyyy"))
    If endDate = 0 Then Exit Sub ' El usuario canceló
    
      ' Ajustar la fecha de fin para incluir todo el día
    endDate = DateAdd("d", 1, endDate) - TimeSerial(0, 0, 1)
    
    ' Validar que las fechas sean correctas
    If startDate > endDate Then
        MsgBox "La fecha de inicio debe ser anterior o igual a la fecha de fin.", vbExclamation
        Exit Sub
    End If
    
    ' Establecer la carpeta donde se guardarán los adjuntos
    saveFolder = "C:\Users\I01142\Desktop\2. FE_NO_BORRAR\"
    
    
    ' Crear la carpeta si no existe
    If Dir(saveFolder, vbDirectory) = "" Then
        MkDir saveFolder
    End If
    
    ' Obtener la aplicación de Outlook
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then
        Set olApp = CreateObject("Outlook.Application")
    End If
    
    ' Obtener el espacio de nombres MAPI
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    ' Obtener la segunda cuenta (ajusta el índice si es necesario)
    Set olAccount = olNamespace.Accounts.Item(1)
    
    ' Obtener la carpeta raíz de la segunda cuenta
    Set olRootFolder = olNamespace.Folders(olAccount.DeliveryStore.DisplayName)
    
    ' Obtener la carpeta de entrada
    Set olInboxFolder = olRootFolder.Folders("Bandeja de entrada")
    
    ' Obtener la subcarpeta "PRUEBA"
    On Error Resume Next
    Set olFolder = olInboxFolder.Folders("FE_ZIP")
    On Error GoTo ErrorHandler
    
    If olFolder Is Nothing Then
        MsgBox "La carpeta 'FE_ZIP' no existe en la bandeja de entrada de la segunda cuenta.", vbExclamation
        GoTo ExitSub
    End If
    
    attachmentsSaved = 0
    
 ' Recorrer todos los elementos en la carpeta
    For Each olItem In olFolder.Items
        ' Verificar si el elemento es un MailItem, tiene adjuntos y está dentro del rango de fechas
        If TypeName(olItem) = "MailItem" Then
            If olItem.ReceivedTime >= startDate And olItem.ReceivedTime <= endDate Then
                If olItem.Attachments.Count > 0 Then
                    For Each olAttachment In olItem.Attachments
                        ' Obtener la extensión del archivo
                        fileExtension = LCase(Right(olAttachment.fileName, 4))
                        
                        ' Verificar si es un archivo .zip
                        If fileExtension = ".zip" Then
                            fileName = saveFolder & olAttachment.fileName
                            ' Evitar sobrescribir archivos existentes
                            If Dir(fileName) <> "" Then
                                fileName = saveFolder & "Copy_" & olAttachment.fileName
                            End If
                            olAttachment.SaveAsFile fileName
                            attachmentsSaved = attachmentsSaved + 1
                        End If
                    Next olAttachment
                End If
            End If
        End If
    Next olItem
    
    If attachmentsSaved > 0 Then
        MsgBox attachmentsSaved & " archivos .zip guardados exitosamente en: " & saveFolder, vbInformation
    Else
        MsgBox "No se encontraron archivos .zip para guardar.", vbInformation
    End If
    
ExitSub:
    Set olAttachment = Nothing
    Set olItem = Nothing
    Set olFolder = Nothing
    Set olInboxFolder = Nothing
    Set olRootFolder = Nothing
    Set olAccount = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume ExitSub
End Sub
