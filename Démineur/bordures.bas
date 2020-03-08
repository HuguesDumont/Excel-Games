Attribute VB_Name = "bordures"
Option Explicit

Public Sub bord(plage As Range)
    With plage.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With plage.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With plage.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With plage.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With plage.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With plage.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    plage.Borders(xlDiagonalDown).LineStyle = xlNone
    plage.Borders(xlDiagonalUp).LineStyle = xlNone
    With plage.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With plage.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With plage.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With plage.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With plage.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With plage.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Public Sub nettoyer()
    Dim rng As Range
    Dim bor As Border
    
    Set rng = ThisWorkbook.Sheets("Démineur").Range("A1:BN50")
    
    With rng
        .Cells.Clear
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    For Each bor In rng.Borders
        bor.LineStyle = xlNone
    Next bor
    
    With rng.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Set rng = ThisWorkbook.Sheets("Valeurs").Range("A1:BN50")
    With rng
        .Cells.Clear
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    For Each bor In rng.Borders
        bor.LineStyle = xlNone
    Next bor
    
    With rng.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
