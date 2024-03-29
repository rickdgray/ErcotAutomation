Attribute VB_Name = "ErcotAutomation"
'@IgnoreModule IndexedDefaultMemberAccess, IndexedUnboundDefaultMemberAccess
'@Folder("ErcotAutomation")
Option Explicit

Public Sub UpdatePrices()
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
    
    ParseCsvPriceTables
    
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
    Dim file As Scripting.FileSystemObject
    Set file = New Scripting.FileSystemObject
    
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

Private Function ParseCsvPriceTables() As Object
    Dim filename As String
    filename = "cdr.00012301.0000000000000000.20240329.010211.SPPHLZNP6905_20240329_0100.csv"
    
    Dim file As Scripting.FileSystemObject
    Set file = New Scripting.FileSystemObject
    
    Dim csvStream As TextStream
    Set csvStream = file.OpenTextFile(Environ$("AppData") & "\ErcotDocumentCache\" & filename, ForReading)
    
    Dim csv As String
    csv = csvStream.ReadAll
    
    Dim csvd As Object
    Set csvd = CSVUtils.ParseCSVToDictionary(csv, 4)
    
    Dim priceTables As New Collection
    
    Dim priceTable As New ErcotPriceTable
    Set priceTable = New ErcotPriceTable
    priceTable.DeliveryDate = CDate(csvd("AMOCO_PUN1")(1))
    priceTable.DeliveryHour = csvd("AMOCO_PUN1")(2)
    priceTable.DeliveryInternal = csvd("AMOCO_PUN1")(3)
    priceTable.SettlementPointName = csvd("AMOCO_PUN1")(4)
    priceTable.SettlementPointType = csvd("AMOCO_PUN1")(5)
    priceTable.SettlementPointPrice = CDec(csvd("AMOCO_PUN1")(6))
    priceTable.DSTFlag = StrComp(csvd("AMOCO_PUN1")(7), "Y", vbTextCompare) = 0
    
    priceTables.Add priceTable
End Function

Private Sub ClearErcotDocumentCache()
    Dim file As Scripting.FileSystemObject
    Set file = New Scripting.FileSystemObject
    
    Dim currentFile As file
    If file.FolderExists((Environ$("AppData") & "\ErcotDocumentCache\")) Then
        For Each currentFile In file.GetFolder((Environ$("AppData") & "\ErcotDocumentCache\")).Files
            currentFile.Delete True
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

