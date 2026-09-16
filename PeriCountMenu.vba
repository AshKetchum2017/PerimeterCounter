Option Explicit

' MacroRunner integration: no reference to the runner project is required.
Private pMRObserver As Object
Private pMRToken As String


Private Sub cmdCalculate_Click()
    Dim doc As Document
    Dim previousUnit As cdrUnit
    Dim selectedShapes As ShapeRange
    Dim shp As Shape
    Dim totalLength As Double

    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo 0

    If doc Is Nothing Then
        txtPerimeter.Text = "Tidak ada dokumen aktif."
        Exit Sub
    End If

    Set selectedShapes = ActiveSelectionRange
    If selectedShapes Is Nothing Then
        txtPerimeter.Text = "Pilih objek path terlebih dahulu."
        Exit Sub
    End If

    If selectedShapes.Count = 0 Then
        txtPerimeter.Text = "Pilih objek path terlebih dahulu."
        Exit Sub
    End If

    previousUnit = doc.Unit
    doc.Unit = cdrMillimeter

    On Error GoTo CleanFail

    For Each shp In selectedShapes
        totalLength = totalLength + GetShapePathLength(shp)
    Next shp

    doc.Unit = previousUnit
    txtPerimeter.Text = Format$(totalLength, "0.00") & " mm"
    Exit Sub

CleanFail:
    doc.Unit = previousUnit
    txtPerimeter.Text = "Gagal menghitung panjang path."
End Sub

Private Function GetShapePathLength(ByVal shp As Shape) As Double
    Dim childShape As Shape
    Dim crv As Curve
    Dim lengthValue As Double

    If shp.Type = cdrGroupShape Then
        For Each childShape In shp.Shapes
            lengthValue = lengthValue + GetShapePathLength(childShape)
        Next childShape

        GetShapePathLength = lengthValue
        Exit Function
    End If

    On Error Resume Next
    Set crv = shp.Curve
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        GetShapePathLength = 0
        Exit Function
    End If
    On Error GoTo 0

    If Not crv Is Nothing Then
        GetShapePathLength = crv.Length
    End If
End Function

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub txtPerimeter_Change()

End Sub

' Called only by MRTargetBridge; normal menu entry points remain unchanged.
Public Sub MRBindRunner(ByVal observer As Object, ByVal token As String)
    Set pMRObserver = observer
    pMRToken = token
End Sub

Public Sub MRDetachRunner()
    Set pMRObserver = Nothing
    pMRToken = vbNullString
End Sub

Private Sub UserForm_Terminate()
    Dim observer As Object, token As String
    On Error GoTo NotifyFailed
    Set observer = pMRObserver
    token = pMRToken
    MRDetachRunner
    If Not observer Is Nothing Then CallByName observer, "MacroUnloaded", VbMethod, token
    Exit Sub
NotifyFailed:
    MsgBox "Gagal memberitahu Macro Runner bahwa form sudah ditutup (" & CStr(Err.Number) & "): " & _
        Err.Description, vbExclamation, "Macro Runner"
End Sub
