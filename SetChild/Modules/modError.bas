Attribute VB_Name = "modError"
'=======================================================================
'Title              :modError
'System             :SetChild
'=======================================================================
'Copyright          :© Albion Software
'Date               :01/09/2004
'Author             :© BombDrop
'Technical Reviewer :
'Purpose            :This is a Generic Error logger for the whole of the
'                   :project.
'=======================================================================

Option Explicit

'=======================================================================
'Procedure :LogError (Sub)
'Date      :  02/07/2002
'Returns   :
'Author    : © BombDrop
'Purpose   :To be an Generic Error logger
'=======================================================================
Public Sub LogError(ByVal strLocation As String, _
    ByVal strErrorDescription As String, ByVal lngErrorNum As Long, _
    Optional ByVal intLine As Integer)

    Dim lngFileNum          As Long
    Dim strErrorMessage     As String
    Dim strErrorLogPath     As String

    On Error GoTo LogError_Error

    strErrorLogPath = App.Path & "\" & App.EXEName & ".Log"

    If intLine = 0 Then

        strErrorMessage = "Error Number :" & lngErrorNum & vbCrLf & _
            "Description  :" & strErrorDescription & vbCrLf & "Location     :" & _
            strLocation & vbCrLf & "Generated at :" & Format(Now, _
            "DDD DD MMM YYYY HH:MM:SS") & vbCrLf


    Else
        strErrorMessage = "Error Number :" & lngErrorNum & vbCrLf & _
            "Description  :" & strErrorDescription & vbCrLf & "Location     :" & _
            strLocation & vbCrLf & "Generated at :" & Format(Now, _
            "DDD DD MMM YYYY HH:MM:SS") & vbCrLf & "LINE         :" & intLine & vbCrLf

    End If


    lngFileNum = FreeFile

    Open strErrorLogPath For Append As lngFileNum

    Print #lngFileNum, strErrorMessage

    Close #lngFileNum


    GoTo CleanExit:

LogError_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ")" & vbCr & _
        "Found In Module: Module1" & vbCr & "Found In Procedure: LogError" & vbCr _
        & IIf(Erl > 0, "Found In Line:" & Erl, ""), vbCritical, "Error Occurred"

CleanExit:
    On Error GoTo 0


End Sub




