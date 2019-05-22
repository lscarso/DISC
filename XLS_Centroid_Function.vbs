Function Centroid(oRng As Range, Optional sType As String) As Variant

'=Centroid(RANGE, "x")' will give the x-coordinate
'=Centroid(RANGE, "y")' will give the y-coordinate
'=Centroid(RANGE, "area")' will give the absolute area
'=Centroid(RANGE, "sarea")' will give the signed area
'=Centroid(RANGE)' will give the coordinate pair to 3 decimal places


    Dim vCoords As Variant
    Dim vRow() As Variant
    Dim i As Long
    Dim Area As Double
    Dim xPos As Double, yPos As Double
    
    vCoords = oRng.Value
    
    ReDim vRow(LBound(vCoords, 2) To UBound(vCoords, 2))
    
    For i = LBound(vCoords, 2) To UBound(vCoords, 2)
        vRow(i) = vCoords(1, i)
    Next i
    
    vCoords = AddRow(vCoords, vRow)
    
    Area = CalcArea(vCoords) 'Note that this is a signed area; if the points are numbered in clockwise order then the area will have a negative sign
    
    xPos = CalcxPos(vCoords, Area)
    yPos = CalcyPos(vCoords, Area)
    
    If UCase(sType) = "X" Then
        Centroid = xPos
    ElseIf UCase(sType) = "Y" Then
        Centroid = yPos
    ElseIf UCase(sType) = "AREA" Then
        Centroid = Abs(Area)
    ElseIf UCase(sType) = "SAREA" Then
        Centroid = Area
    Else
        Centroid = "(" & Round(xPos, 3) & "," & Round(yPos, 3) & ")"
    End If
    
    
End Function
Private Function CalcxPos(vCoords As Variant, Area As Double) As Double
    Dim i As Long
    
    For i = 1 To UBound(vCoords, 1) - 1
        CalcxPos = CalcxPos + (vCoords(i, 1) + vCoords(i + 1, 1)) * (vCoords(i, 1) * vCoords(i + 1, 2) - vCoords(i + 1, 1) * vCoords(i, 2))
    Next i
    
    CalcxPos = CalcxPos / (6 * Area)
    
End Function
Private Function CalcyPos(vCoords As Variant, Area As Double) As Double
    Dim i As Long
    
    For i = 1 To UBound(vCoords, 1) - 1
        CalcyPos = CalcyPos + (vCoords(i, 2) + vCoords(i + 1, 2)) * (vCoords(i, 1) * vCoords(i + 1, 2) - vCoords(i + 1, 1) * vCoords(i, 2))
    Next i
    
    CalcyPos = CalcyPos / (6 * Area)
    
End Function
Private Function CalcArea(vCoords As Variant) As Double
    Dim i As Long
    
    For i = 1 To UBound(vCoords, 1) - 1
        CalcArea = CalcArea + vCoords(i, 1) * vCoords(i + 1, 2) - vCoords(i + 1, 1) * vCoords(i, 2)
    Next i
    
    CalcArea = 0.5 * CalcArea
    
End Function
Private Function AddRow(InputArr As Variant, vRow As Variant) As Variant
    Dim vTemp As Variant
    Dim i As Long
    
    If LBound(vRow) <> LBound(InputArr, 2) Or UBound(vRow) <> UBound(InputArr, 2) Then AddRow = 0: Exit Function
    
    vTemp = TransposeArray(InputArr)
    
    ReDim Preserve vTemp(LBound(vTemp, 1) To UBound(vTemp, 1), LBound(vTemp, 2) To UBound(vTemp, 2) + 1)
    
    vTemp = TransposeArray(vTemp)
    
    For i = LBound(vTemp, 2) To UBound(vTemp, 2)
        vTemp(UBound(vTemp, 1), i) = vRow(i)
    Next i
    
    AddRow = vTemp
    
End Function
Private Function TransposeArray(InputArr As Variant) As Variant


Dim RowNdx As Long
Dim ColNdx As Long
Dim LB1 As Long
Dim LB2 As Long
Dim UB1 As Long
Dim UB2 As Long
Dim OutputArr() As Variant


LB1 = LBound(InputArr, 1)
LB2 = LBound(InputArr, 2)
UB1 = UBound(InputArr, 1)
UB2 = UBound(InputArr, 2)


ReDim OutputArr(LB2 To LB2 + UB2 - LB2, LB1 To LB1 + UB1 - LB1)


For RowNdx = LB2 To UB2
    For ColNdx = LB1 To UB1
        OutputArr(RowNdx, ColNdx) = InputArr(ColNdx, RowNdx)
    Next ColNdx
Next RowNdx


TransposeArray = OutputArr


End Function
