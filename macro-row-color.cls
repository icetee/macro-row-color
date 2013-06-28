VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Munka1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Const color As String = "jVeY7UZj"
Const cella As String = "A1"
Const feltetel As String = "+"

'Figyeli a válotzásokat a "Tartomany" részben
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Tartomany) Is Nothing Then Szinezes
End Sub

'A "GAZDINFO" táblában és a "Tartomany" területen változtatja a sor szinét, adott feltételre
Sub Szinezes()
    Dim cella As Range
    Dim oszlop As Integer

    With GAZDINFO
        For Each cella In Tartomany
            If cella.Text = feltetel Then
                For oszlop = 1 To Sorok.Columns.Count
                    Cells(cella.row, oszlop).Interior.color = 5287936
                    'Range("A5").Interior.color = 5287936
                Next
            Else
                For oszlop = 1 To Sorok.Columns.Count
                    Cells(cella.row, oszlop).Interior.ColorIndex = 2
                Next
            End If
       Next
    End With
End Sub

'Tartomány kijelölése
Public Sub Kijeloles()
    Dim jeloles As Range
    Dim jeloles2 As Range
    
    'Adat bekérési ablak
    On Error Resume Next
    Set jeloles = Application.InputBox(prompt:="Jelöld ki az ellenörzésre szánt területet:", Type:=8)
    On Error GoTo 0

    If jeloles Is Nothing Then
        Exit Sub
    Else
        'Adat eltárolása
        Sheets(color).Range(cella).Value = jeloles.Address
    End If
    
    '------------
    On Error Resume Next
    'Adat bekérési ablak
    Set jeloles2 = Application.InputBox(prompt:="Jelöld ki a színezésre szánt oszlopokat:", Type:=8)
    On Error GoTo 0
    
    If jeloles2 Is Nothing Then
        Exit Sub
    Else
        'Adat eltárolása
        Sheets(color).Range("A2").Value = jeloles2.Address
    End If

End Sub

Public Function Tartomany() As Range
    Munkalap (color)
    Call Adat
    'Adat
    Worksheets(color).Visible = True
    'Sheets(color).Visible = xlVeryHidden
    Set Tartomany = Range(Worksheets(color).Range(cella))
    
End Function

Public Function Sorok() As Range
    Set Sorok = Range(Worksheets(color).Range("A2"))
    
End Function

'Munkalap létezésének vizsgálata
Function Munkalap(Name As String) As Worksheet
    On Error Resume Next
    Set Munkalap = ThisWorkbook.Worksheets(Name)

    If Munkalap Is Nothing Then
        Set Munkalap = ThisWorkbook.Worksheets.Add
        Munkalap.Name = Name
        Call Kijeloles
    End If

End Function

'Tartomány érték meglétének vizsgálata
Function Adat() As Range
 On Error Resume Next
   Set Adat = Range(Worksheets(color).Range(cella))
   
    If Adat Is Nothing Then
        Call Kijeloles
    End If
End Function

Function test() As Range
    MsgBox Sorok.Columns.Count
End Function
