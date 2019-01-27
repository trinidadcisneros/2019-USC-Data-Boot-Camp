Attribute VB_Name = "cisneros_HW2_vba_sln"
'I use excel on a daily basis at work, and I intend to learn more VBA scripting to help automate the workload.
'To this end, my learning objectives were to
    'a) Prepare a macro that loops through all the worksheets in a workbook and converts data into table objects.
    'b) Prepare a macro that sorts data
    'c) Create a dictionary (since I struggle with this concept a bit), and use this strategy to calculate some value for a specific key
    'd) Use a macro that would iterate through a workbook and apply the dictionary script
    'e) Explore and use advance range referecing scripts, such as offset and resize
    'General Comments: I had a really hard time with objective d), it turns out that the code below is a key part of this strategy, and
    'it took me a long time to realize I could use the wellsfargo activity to get the worksheetname, and I had to google other trageties
    'to get the listObject parameter, which is the table name, and as you can see I had to get creative.
    'Set tables = Worksheets(WorksheetName).ListObjects(ActiveSheet.ListObjects(ActiveSheet.Index).Name)
    'As a result of the time I spent to constructs the macros here (which I will use at work) I was unable to complete the other tasks.
    

Sub format_cal_total_volume_with_dict():

    'Iterate through entire workbook
    For Each ws In Worksheets
        
        'Variable to hold the names of each worksheet, to be used for the dictionary
        Dim WorksheetName As String
        
        'Gets the worksheet name, this will be used for the dictionary
        WorksheetName = ws.Name
        
        ' Alphabetizes the ticker column--------------|
        ws.Range("A1").CurrentRegion.Sort key1:=ws.Range("A1"), order1:=xlAscending, Header:=xlGuess
        
        ' Converts all worksheets into table objects--|
        Dim src As Range
        Set src = ws.Range("A1").CurrentRegion
        ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=src, _
        xllistobjecthasHeaders:=xlYes, tablestyleName:="TableStyleLight18").Name = "Table"
        
        
        'USE DICTIONARY TO CALCULATE THE TOTAL VOLUME PER TICKER-----------|
        'Create the dictionary---------------------------------------------|
        'Sub ticker_dictionary():
        
        'Setup the variables that will be needed for the dictionary
        Dim dictData As Object
        Dim bItemExists As Boolean
        Dim tables As ListObject
        Dim arrData, arrDictValues, arrDictKeys
        Dim i_dict As Long
        Dim rng As Range
        
        
        'create the dictionary object
        Set dictData = CreateObject("Scripting.Dictionary")
        
        'creates a table object that gets the sheet name, and also the table name from sheet
        Set tables = Worksheets(WorksheetName).ListObjects(ActiveSheet.ListObjects(ActiveSheet.Index).Name)
        
        'put the data into an array for faster processing
        arrData = tables.DataBodyRange
        
        'looop through the array
        For i_dict = 1 To UBound(arrData)
            'if the key exist, add to it
            If dictData.Exists(arrData(i_dict, 1)) Then
                dictData.Item(arrData(i_dict, 1)) = dictData.Item(arrData(i_dict, 1)) + _
                    arrData(i_dict, 7)
            Else
                dictData.Add arrData(i_dict, 1), arrData(i_dict, 7)
            End If
            'else create and add to it
        Next i_dict
        
        'the range will be 2 columns to the right of the table
        Set rng = tables.Range.Offset(1, tables.Range.Columns.Count).Resize(1, 1)
        
        'put the dictionary keys into the array arrHeaders
        arrDictKeys = dictData.Keys
        
        'put all the keys one column from the end of the starting table
        rng.Resize(dictData.Count, 1).Value = Application.Transpose(arrDictKeys)
        
        'put all the dictionary values into the array arrReport
        arrDictValues = dictData.Items
        'Put all items one column over the range that the length of the table, and 1 column
        rng.Offset(, 1).Resize(dictData.Count, 1).Value = _
            Application.Transpose(arrDictValues)
        Set dictData = Nothing
        Set tables = Nothing
        Set rng = Nothing
        
        'Label Ticker and Total Stock Volume columns-----------------------|
        Range("H1") = "Ticker"
        Range("I1") = "Total Stock Volume"
        
  Next ws
  
End Sub


