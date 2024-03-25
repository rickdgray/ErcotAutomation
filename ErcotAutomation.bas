Attribute VB_Name = "ErcotAutomation"
'@Folder("ErcotAutomation")
Option Explicit

Sub UpdatePrices()
    Dim documents As New Collection
    Set documents = FetchAllErcotDocs("12301")
    
    Dim mostRecent As Date
    Dim currDate As Date
    
    Dim doc As ErcotDocument
    For Each doc In documents
        currDate = doc.PublishDate
        If currDate > mostRecent Then
            mostRecent = currDate
        End If
    Next doc
    
    Dim startDate As Date
    startDate = DateAdd("d", -2, mostRecent)
 
    For Each doc In documents
        If doc.PublishDate > startDate Then
            DownloadZip doc.DocId, doc.ConstructedName
        End If
    Next doc
    
    UnzipAllFilesInFolder Environ$("AppData") & "\ErcotDocumentCache\"
    
    ClearErcotDocumentCache
End Sub

Private Function FetchAllErcotDocs(ByVal reportTypeId As String) As Collection
    Dim Client As Object
    
    Set Client = CreateObject("WinHttp.WinHttpRequest.5.1")
    Client.Open "GET", "https://www.ercot.com/misapp/servlets/IceDocListJsonWS?reportTypeId=" & reportTypeId, False
    Client.setRequestHeader "Content-Type", "application/json"
    Client.setRequestHeader "Accept", "application/json"
    Client.send
    
    Dim Json As Object
    Set Json = JsonConverter.ParseJson(Client.ResponseText)
    
    Dim documents As New Collection
    Dim document As New ErcotDocument
    
    Dim doc As Dictionary
    For Each doc In Json("ListDocsByRptTypeRes")("DocumentList")
        If GetFileExtension(doc("Document")("FriendlyName")) = "csv" Then
            Set document = New ErcotDocument
            document.DocId = doc("Document")("DocID")
            document.ConstructedName = doc("Document")("ConstructedName")
            document.PublishDate = ConvertIsoTimestamp(doc("Document")("PublishDate"))
            documents.Add document
        End If
    Next doc
    
    Set FetchAllErcotDocs = documents
End Function

Private Sub DownloadZip(ByVal DocId As String, ByVal filename As String)
    Dim file As Object
    Set file = CreateObject("Scripting.FileSystemObject")
    
    If Not (file.FolderExists(Environ$("AppData") & "\ErcotDocumentCache\")) Then
        MkDir Environ$("AppData") & "\ErcotDocumentCache\"
    End If
    
    Dim Client As Object
    Set Client = CreateObject("WinHttp.WinHttpRequest.5.1")
    Client.Open "GET", "https://www.ercot.com/misdownload/servlets/mirDownload?doclookupId=" & DocId, False
    Client.send
    
    Dim stream As Object
    If Client.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        With stream
            .Open
            .Type = 1
            .Write Client.ResponseBody
            .SaveToFile Environ$("AppData") & "\ErcotDocumentCache\" & filename, 2
            .Close
        End With
    Else
        MsgBox "Download for document ID " & DocId & " failed!", vbExclamation, "Status " & Client.Status
    End If
End Sub

Private Sub UnzipAllFilesInFolder(ByVal path As String)
    Dim winShell As Shell
    Dim zipArchives As FolderItems
    
    Set winShell = CreateObject("Shell.Application")
    Set zipArchives = winShell.Namespace(path).Items
    
    Dim i As Long
    Dim zipItem As FolderItem
    For i = 0 To zipArchives.Count - 1
        Set zipItem = zipArchives.Item(i)
        'we assume only 1 file per zip
        winShell.Namespace(path).CopyHere winShell.Namespace(zipItem).Items.Item(0)
    Next i
End Sub

Private Sub ClearErcotDocumentCache()
    Dim file As Object
    Set file = CreateObject("Scripting.FileSystemObject")
    
    Dim f As file
    If file.FolderExists((Environ$("AppData") & "\ErcotDocumentCache\")) Then
        For Each f In file.GetFolder((Environ$("AppData") & "\ErcotDocumentCache\")).Files
            f.Delete True
        Next
    End If
End Sub

Private Function ConvertIsoTimestamp(ByVal value As String) As Date
    'Wrapper function to handle pass by value
    ConvertIsoTimestamp = UtcConverter.ParseISOTimeStampToUTC(value)
End Function

Private Function GetFileExtension(ByVal Name As String) As String
    GetFileExtension = Mid$(Name, Len(Name) - 2, 3)
End Function

