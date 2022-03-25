VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)

Application.DisplayAlerts = False
Application.ScreenUpdating = False

Sheets("Répartition des comptes").Range("M1") = Now
Sheets("CAM détail").Select
    Range("U4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("V4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
        Range("A1").Select
    Sheets("IARD détail").Select
    
    Range("U4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("V4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False
        Range("A1").Select
    Sheets("VIE détail").Select
    
    
    Range("u4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Range("v4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Application.CutCopyMode = False
        Range("A1").Select
        Sheets("Notice").Select
        
        
'permet de copier la feuille de récupération des cours
Dim sht As Worksheet
Dim a As String
Dim exist As String
Dim exist2 As String

 Sheets("cours").Visible = True
 Sheets("donnees").Visible = True
 On Error Resume Next
 Sheets("Bloom1").Visible = True
Sheets("Bloom").Visible = True

        
For Each sht In ThisWorkbook.Worksheets
a = "Cours " & Format(Date, "dd mm yy")
exist = False
If sht.Name = a Then
        exist = True
        Exit For
    End If

Next sht
        
     If exist = True Then
     
    GoTo Label3
    Else
        
        
           Sheets("cours").Visible = True
         Sheets("cours").Copy After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Visible = True
    Sheets(Sheets.Count).Select
    Range("C3").Select
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Sheets(Sheets.Count).Name = "Cours " & Format(Date, "dd mm yy")
    
    'Replace(CStr(Date), "/", "-")
    Sheets(1).Visible = True
    Sheets(Sheets.Count).Move Before:=Sheets(1)
    'Sheets(Sheets.Count).Select
  
    
        'Sheets(Sheets.Count).Name = "cours"
        
        
        End If
        
        '//////////////////////////////////////DONNEES///////////////////////////////////////////////////////////
        
Label3:
For Each sht In ThisWorkbook.Worksheets
b = "donnees " & Format(Date, "dd mm yy")
exist2 = False
If sht.Name = b Then
        exist2 = True
        Exit For
    End If

Next sht
        
     If exist2 = True Then
     
     
        
        
        GoTo Label2
        Else
        
        
           Sheets("donnees").Visible = True
         Sheets("donnees").Copy After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Visible = True
    Sheets(Sheets.Count).Select
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Sheets(Sheets.Count).Name = "donnees " & Format(Date, "dd mm yy")
    
    'Replace(CStr(Date), "/", "-")
    Sheets(1).Visible = True
    Sheets(Sheets.Count).Move Before:=Sheets(1)
    'Sheets(Sheets.Count).Select
   
    
        'Sheets(Sheets.Count).Name = "cours"
        
        
        End If
        
        '''''''''''''''''''''''''permet de copier en dur la feuille carac'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Label2:
         If Sheets(Sheets("Bloom").Index + 1).Name = "Bloom1" Then
        GoTo Lable1
        Else
        
        Sheets("Bloom").Visible = True
           
         Sheets("Bloom").Copy After:=Sheets("Bloom")
    
    Sheets(Sheets("Bloom").Index + 1).Select
    Range("A1").Select
    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Sheets(Sheets("Bloom").Index + 1).Name = "Bloom1"
    
    'Replace(CStr(Date), "/", "-")
    
  End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''FIN''''''''''''''
        
        
Lable1:
        
   Dim I&
        Sheets("Notice").Select
    For I = 1 To Sheets("Notice").Index - 1
                Sheets(I).Visible = xlSheetVeryHidden
       
    Next I
 Sheets("cours").Visible = xlSheetVeryHidden
 Sheets("donnees").Visible = xlSheetVeryHidden
 On Error Resume Next
 Sheets("Bloom1").Visible = xlSheetVeryHidden
Sheets("Bloom").Visible = xlSheetVeryHidden
        
'ActiveWorkbook.SaveCopyAs "T:\finance & tresorerie & gestion immobiliere\commun direction financiere\Valorisation\justifs saisie manuelle\Edp Bloom\Edp Bloomberg_copie " & Format(Date, "dd mm yy") & ".xlsm"
        
        
        Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.AlertBeforeOverwriting = True
ActiveWorkbook.Close True

End Sub


