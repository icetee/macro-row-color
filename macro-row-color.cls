VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Munka1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'Figyeli a válotzásokat a "Tartomany" részben
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Tartomany) Is Nothing Then Szinezes
End Sub

'A "GAZDINFO" táblában és a "Tartomany" területen változtatja a sor szinét, adott feltételre
Sub Szinezes()
Attribute Szinezes.VB_ProcData.VB_Invoke_Func = "e\n14"
    Dim cella As Range
    Dim oszlop As Integer

    With GAZDINFO
        For Each cella In Tartomany
            If cella.Text = "+" Then
                For oszlop = 1 To 5
                    Cells(cella.row, oszlop).Interior.Color = 5287936
                Next
            Else
                For oszlop = 1 To 5
                    Cells(cella.row, oszlop).Interior.ColorIndex = 2
                Next
            End If
       Next
    End With
End Sub

'Tartomány funkció, a 400-as hiba elkerülése érdekében
Function Tartomany() As Range
    Set t1 = Range("B4:B12")
    Set t2 = Range("B14:B25")
    Set t3 = Range("B27:B44")
    Set t4 = Range("B46:B47")
    Set t5 = Range("B50:B58")
    Set t6 = Range("B60:B63")
    
    Set Tartomany = Union(t1, t2, t3, t4, t5, t6)
End Function
