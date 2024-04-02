Attribute VB_Name = "ErcotAutomation"
'@IgnoreModule IndexedDefaultMemberAccess, IndexedUnboundDefaultMemberAccess
'@Folder("ErcotAutomation")
Option Explicit

Public Sub UpdatePrices()
    ClearErcotDocumentCache
    
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
    
    Dim cvcCc1PriceTable As Collection
    Set cvcCc1PriceTable = ParseCsvFilesToPriceTable("CVC_CC1")
    
    Dim hbHoustonPriceTable As Collection
    Set hbHoustonPriceTable = ParseCsvFilesToPriceTable("HB_HOUSTON")
    
    Dim lhmCvcG4PriceTable As Collection
    Set lhmCvcG4PriceTable = ParseCsvFilesToPriceTable("LHM_CVC_G4")
    
    Dim today As String
    today = Format$(Date, "yyyy-mm-dd")
    Dim mostRecentHour As Long
    Dim mostRecentHourKey As String
    mostRecentHourKey = vbNullString
    
    Dim cvcCc1AccumulatedPrices As New Dictionary
    Dim cvcCc1PriceRecord As ErcotPriceRecord
    For Each cvcCc1PriceRecord In cvcCc1PriceTable
        'Need to track the most recent hour on the way for later
        If cvcCc1PriceRecord.DeliveryDate = Date And cvcCc1PriceRecord.DeliveryHour > mostRecentHour Then
            mostRecentHour = cvcCc1PriceRecord.DeliveryHour
            mostRecentHourKey = cvcCc1PriceRecord.DeliveryDate & cvcCc1PriceRecord.DeliveryHour
        End If
        
        AccumulateOnDict cvcCc1PriceRecord.DeliveryDate & cvcCc1PriceRecord.DeliveryHour, cvcCc1PriceRecord.SettlementPointPrice, cvcCc1AccumulatedPrices
    Next cvcCc1PriceRecord
    
    Dim cvcCc1AveragePrices As New Dictionary
    Dim values As Collection
    Dim key As Variant
    For Each key In cvcCc1AccumulatedPrices.Keys
        'Make sure we got all 4 intervals
        If cvcCc1AccumulatedPrices(key).Count = 4 Then
            cvcCc1AveragePrices(key) = CollectionSum(cvcCc1AccumulatedPrices(key)) / 4
        End If
        
        'Also get the intervals right now, as this hour hasn't completed
        If key = mostRecentHourKey Then
            cvcCc1AveragePrices(key) = CollectionSum(cvcCc1AccumulatedPrices(key)) / cvcCc1AccumulatedPrices(key).Count
        End If
    Next
    
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

Public Function ParseCsvFilesToPriceTable(ByVal name As String) As Collection
    Dim file As Scripting.FileSystemObject
    Set file = New Scripting.FileSystemObject
    
    Dim priceTable As New Collection
    
    Dim currentFile As file
    If file.FolderExists((Environ$("AppData") & "\ErcotDocumentCache\")) Then
        For Each currentFile In file.GetFolder((Environ$("AppData") & "\ErcotDocumentCache\")).Files
            If GetFileExtension(currentFile.path) = "csv" Then
                Dim csvStream As TextStream
                Set csvStream = file.OpenTextFile(currentFile.path, ForReading)
                
                Dim csv As String
                csv = csvStream.ReadAll
                
                Dim csvd As Object
                Set csvd = CSVUtils.ParseCSVToDictionary(csv, 4)
                
                Dim priceRecord As New ErcotPriceRecord
                Set priceRecord = New ErcotPriceRecord
                priceRecord.DeliveryDate = CDate(csvd(name)(1))
                priceRecord.DeliveryHour = csvd(name)(2)
                priceRecord.DeliveryInterval = csvd(name)(3)
                priceRecord.SettlementPointName = csvd(name)(4)
                priceRecord.SettlementPointPrice = CDec(csvd(name)(6))
                
                priceTable.Add priceRecord
            End If
        Next
    End If
    
    Set ParseCsvFilesToPriceTable = priceTable
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

Private Function CollectionSum(ByRef col As Collection) As Variant
    CollectionSum = 0
    
    Dim val As Variant
    For Each val In col
        CollectionSum = CollectionSum + val
    Next
End Function

Private Sub AccumulateOnDict(ByVal key As String, ByVal value As Variant, ByRef dict As Dictionary)
    If dict.Exists(key) = False Then
        Set dict(key) = New Collection
    End If
    
    dict(key).Add value
End Sub

Private Function ConvertIsoTimestamp(ByVal value As String) As Date
    'Wrapper function to handle pass by value
    ConvertIsoTimestamp = UtcConverter.ParseISOTimeStampToUTC(value)
End Function

Private Function GetFileExtension(ByVal name As String) As String
    GetFileExtension = Mid$(name, Len(name) - 2, 3)
End Function

Private Function StringEquals(ByVal a As String, ByVal b As String) As Boolean
    StringEquals = StrComp(a, b, vbTextCompare) = 0
End Function
