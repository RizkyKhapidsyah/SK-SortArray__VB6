VERSION 5.00
Begin VB.Form Sort 
   AutoRedraw      =   -1  'True
   Caption         =   "Array Sort"
   ClientHeight    =   2940
   ClientLeft      =   975
   ClientTop       =   1560
   ClientWidth     =   6585
   LinkTopic       =   "Form2"
   ScaleHeight     =   2940
   ScaleWidth      =   6585
   Begin VB.CommandButton Command1 
      Caption         =   "Sort On Column Two"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   1
      Top             =   2400
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sort On Column One"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   2400
      Width           =   2535
   End
End
Attribute VB_Name = "Sort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Created by Rizky Khapidsyah

Dim DataArray()
Sub LoadArray()
    ReDim DataArray(2, 9)
    DataArray(0, 0) = "Pisang": DataArray(1, 0) = "Kucing"   '0
    DataArray(0, 1) = "Mangga": DataArray(1, 1) = "Ikan"     '7
    DataArray(0, 2) = "Nanas": DataArray(1, 2) = "Ayam"    '4
    DataArray(0, 3) = "Jambu": DataArray(1, 3) = "Bebek"     '6
    DataArray(0, 4) = "Jeruk": DataArray(1, 4) = "Ular"   '8
    DataArray(0, 5) = "Anggur": DataArray(1, 5) = "Ulat"    '2
    DataArray(0, 6) = "Apel": DataArray(1, 6) = "Monyet"      '5
    DataArray(0, 7) = "Pepaya": DataArray(1, 7) = "OrangUtan"     '1
    DataArray(0, 8) = "Nanas": DataArray(1, 8) = "Harimau"    '3
End Sub

Public Sub SortArray(ByRef DArray(), Element As Integer)
    Dim gap As Integer, doneflag As Integer, SwapArray()
    Dim Index As Integer
    ReDim SwapArray(2, UBound(DArray, 1), UBound(DArray, 2))
    'Gap is half the records
    gap = Int(UBound(DArray, 2) / 2)
    Do While gap >= 1
        Do
            doneflag = 1
            For Index = 0 To (UBound(DArray, 2) - (gap + 1))
                'Compare 1st 1/2 to 2nd 1/2
                If DArray(Element, Index) > DArray(Element, (Index + gap)) Then
                    For acol = 0 To (UBound(SwapArray, 2) - 1)
                        'Swap Values if 1st > 2nd
                        SwapArray(0, acol, Index) = DArray(acol, Index)
                        SwapArray(1, acol, Index) = DArray(acol, Index + gap)
                    Next
                    For acol = 0 To (UBound(SwapArray, 2) - 1)
                        'Swap Values if 1st > 2nd
                        DArray(acol, Index) = SwapArray(1, acol, Index)
                        DArray(acol, Index + gap) = SwapArray(0, acol, Index)
                    Next
                    CNT = CNT + 1
                    doneflag = 0
                End If
            Next
        Loop Until doneflag = 1
        gap = Int(gap / 2)
    Loop
End Sub
Private Sub Command1_Click(Index As Integer)
    'Reload Unsorted Data
    'LoadArray
    'Pass the array with data to be sorted
    SortArray DataArray(), Index
    Sort.Cls
    For R = 0 To UBound(DataArray, 2) - 1
        For C = 0 To UBound(DataArray, 1) - 1
            Sort.Print DataArray(C, R) & Space(5);
        Next
        Sort.Print
    Next
End Sub

Private Sub Form_Load()
    'Load Unsorted Data
    LoadArray
    'Clear Form
    Sort.Cls
    'Display Raw data
    For R = 0 To UBound(DataArray, 2) - 1
        For C = 0 To UBound(DataArray, 1) - 1
            Sort.Print DataArray(C, R) & Space(5);
        Next
        Sort.Print
    Next
End Sub
