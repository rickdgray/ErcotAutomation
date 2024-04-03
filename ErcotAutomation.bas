Attribute VB_Name = "ErcotAutomation"
'@IgnoreModule UseMeaningfulName, SelfAssignedDeclaration, IndexedDefaultMemberAccess, IndexedUnboundDefaultMemberAccess
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
    
    Dim key As Variant
    Dim sheet As Worksheet
    Set sheet = ActiveWorkbook.Sheets("Sheet1")
    Dim cell As Range
    Dim row As Long
    row = 2
    
    Dim cvcCc1AveragePrices As Dictionary
    Set cvcCc1AveragePrices = AccumulateAndAverageAllPricesByHour(cvcCc1PriceTable)
    
    Dim hbHoustonAveragePrices As Dictionary
    Set hbHoustonAveragePrices = AccumulateAndAverageAllPricesByHour(hbHoustonPriceTable)
    
    Dim lhmCvcG4AveragePrices As Dictionary
    Set lhmCvcG4AveragePrices = AccumulateAndAverageAllPricesByHour(lhmCvcG4PriceTable)
    
    Dim hoursList As Dictionary
    Set hoursList = BuildHoursList(startDate)
    
    For Each key In hoursList.Keys
        If cvcCc1AveragePrices.Exists(key) = True Then
            Set cell = sheet.Range("A" & row)
            cell.value = hoursList(key)
            Set cell = sheet.Range("B" & row)
            cell.value = cvcCc1AveragePrices(key)
            Set cell = sheet.Range("C" & row)
            cell.value = hbHoustonAveragePrices(key)
            Set cell = sheet.Range("D" & row)
            cell.value = lhmCvcG4AveragePrices(key)
            row = row + 1
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
                priceRecord.SettlementPointName = csvd(name)(4)
                priceRecord.SettlementPointPrice = CDec(csvd(name)(6))
                
                priceTable.Add priceRecord
            End If
        Next
    End If
    
    Set ParseCsvFilesToPriceTable = priceTable
End Function

Private Function AccumulateAndAverageAllPricesByHour(ByVal priceTable As Collection) As Dictionary
    Dim mostRecentHour As Long
    Dim mostRecentHourKey As String
    mostRecentHourKey = vbNullString
    
    Dim accumulatedPrices As New Dictionary
    Dim priceRecord As ErcotPriceRecord
    For Each priceRecord In priceTable
        'Need to track the most recent hour on the way for later
        If priceRecord.DeliveryDate = Date And priceRecord.DeliveryHour > mostRecentHour Then
            mostRecentHour = priceRecord.DeliveryHour
            mostRecentHourKey = priceRecord.DeliveryDate & priceRecord.DeliveryHour
        End If
        
        AccumulateOnDict priceRecord.DeliveryDate & priceRecord.DeliveryHour, priceRecord.SettlementPointPrice, accumulatedPrices
    Next priceRecord
    
    Dim averagePrices As New Dictionary
    Dim key As Variant
    For Each key In accumulatedPrices.Keys
        'Make sure we got all 4 intervals and also get the intervals right now,
        'as this hour hasn't completed
        If accumulatedPrices(key).Count = 4 Or key = mostRecentHourKey Then
            averagePrices(key) = CollectionSum(accumulatedPrices(key)) / accumulatedPrices(key).Count
        End If
    Next
    
    Set AccumulateAndAverageAllPricesByHour = averagePrices
End Function

Private Function BuildHoursList(ByVal startingDate As Date) As Dictionary
    Dim currentDate As Date
    currentDate = startingDate
    Dim hoursList As New Dictionary
    Dim i As Long
    
    For i = 1 To 24
        hoursList(Format$(currentDate, "yyyy-mm-dd") & i) = Format$(currentDate, "yyyy-mm-dd") & " Hour: " & i
    Next i
    currentDate = DateAdd("d", 1, currentDate)
    
    For i = 1 To 24
        hoursList(Format$(currentDate, "yyyy-mm-dd") & i) = Format$(currentDate, "yyyy-mm-dd") & " Hour: " & i
    Next i
    
    '@Ignore AssignmentNotUsed
    'false positive
    currentDate = DateAdd("d", 1, currentDate)
    
    For i = 1 To 24
        hoursList(Format$(currentDate, "yyyy-mm-dd") & i) = Format$(currentDate, "yyyy-mm-dd") & " Hour: " & i
    Next i
    
    Set BuildHoursList = hoursList
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

