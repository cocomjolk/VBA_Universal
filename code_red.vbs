
Function Logic()
    Dim ResultsDict As Dictionary
    Dim CurrentResultDict As Dictionary
    Dim cel As Range

    Set ResultsDict = New Dictionary


    'Cycle through column B and match with column H,I,J
    lastRow = Range("B2").End(xlDown).Row
    For Each cel In Range("B2:B" & lastRow).Cells
        Set CurrentResultDict = New Dictionary
        wkDate = cel.Value
        fiscalVer = cel.Offset(0, 6).Value
        fiscalMW = cel.Offset(0, 7).Value
        hyperionVer = cel.Offset(0, 8).Value

        CurrentResultDict.Add "address", fiscalVer
        CurrentResultDict.Add "name", fiscalMW
        CurrentResultDict.Add "range", hyperionVer


        ResultsDict.Add wkDate, CurrentResultDict
    Next

    Set getCalendarDict = ResultsDict

End Function

Sub LoopThroughResults(ResultsDict)

  For result in ResultsDict
      address = result.Item("address")

      Range(address).Copy

  Next result


End Sub
