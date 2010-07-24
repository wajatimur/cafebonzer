Attribute VB_Name = "MdlStringMathDate"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MdlStringMathDate
'    Project    : CafeBonzer
'
'    Description:
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
'==================================================================
' Aplication codename : CafeBonzer
' Programmer          : Azri Jamil a.k.a wajatimur
' Module Name         : String Control
' Description         :
'==================================================================


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
' Nilai yang di roundup (depend on setting)
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
Function GetRoundUpVal(dblValue As Double) As Double
    If SetGetDb("FinPriceRound", 1) = 1 Then
        GetRoundUpVal = Round(dblValue, 1)
    Else
        GetRoundUpVal = dblValue
    End If
End Function

Public Function GetMonthString(MonthNumber)
    GetMonthString = Choose(MonthNumber, "Januari", "Februari", "Mac", "April", "Mei", "Jun", "Julai", "Ogos", "September", "Oktober", "November", "Disember")
End Function
