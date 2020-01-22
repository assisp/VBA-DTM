VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Rem -----------------------------------------------------------------------------------------------------------------------
Rem --- Internal Variables                                                                           ---
Rem -----------------------------------------------------------------------------------------------------------------------
Private myFuzz As Double 'Fuzz factor for point comparison
Private thisList()    As Variant
Private lsize         As Long 'list size

Rem -----------------------------------------------------------------------------------------------------------------------
Rem --- CONSTRUCTORS / DESTRUCTORS                                                                    ---
Rem -----------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    myFuzz = 0.001
    lsize = 0
End Sub      '   Constructor

Rem -----------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    lsize = 0
    Erase thisList
End Sub

Rem -----------------------------------------------------------------------------------------------------------------------
Rem --- CLASS GET/LET/SET PROPERTIES                                                                 ---
Rem -----------------------------------------------------------------------------------------------------------------------

Public Property Get Fuzz() As Double
    Fuzz = myFuzz
End Property

Public Property Let Fuzz(ByVal pValue As Double)
    myFuzz = pValue
End Property

Rem -----------------------------------------------------------------------------------------------------------------------
Rem --- CLASS METHODS                                                                              ---
Rem -----------------------------------------------------------------------------------------------------------------------

Rem ---
Rem --- returns the list size
Rem ---
Public Function Count() As Long
    Count = lsize
End Function

Rem ---
Rem --- empties the list
Rem ---
Public Sub Clean()
    lsize = 0
    Erase thisList
End Sub

Rem ---
Rem --- Add item to list
Rem ---
Public Sub Add(ByVal Item As Variant)
   lsize = lsize + 1
   ReDim Preserve thisList(1 To lsize)
   thisList(lsize) = Item
End Sub

Rem ---
Rem --- set item at index
Rem ---
Public Sub SetItem(ByVal index As Long, ByVal Item As Variant)
    If index <= lsize And index > 0 Then
        thisList(index) = Item
    End If
End Sub

Rem ---
Rem --- get item at index
Rem ---
Public Function Item(ByVal index As Long) As Variant
    If index <= lsize And index > 0 Then
        Item = thisList(index)
    Else
        Item = Empty
    End If
End Function

Rem ---
Rem --- swap items at index1 and index2
Rem ---
Public Sub swap(ByVal index1 As Long, ByVal index2 As Long)
    Dim aux As Variant
    
    If index1 <= lsize And index2 <= lsize And index1 > 0 And index2 > 0 Then
        aux = thisList(index1)
        thisList(index1) = thisList(index2)
        thisList(index2) = aux
    End If
End Sub

Rem ---
Rem --- Insertion sort by XY if items are Array(2) of Doubles (3d Point)
Rem --- FIXME: check if item is array(2) of Doubles
Rem ---
Public Sub SortXY()
    Dim j, i, n As Long
    Dim Pt As Variant
    ' Upper Bound of Array
    n = UBound(thisList)
    For i = 2 To n
        j = i
        'FIXME: check if item is array(2) of Doubles
        Pt = thisList(i)
        Do While j > 1
            If thisList(j - 1)(0) > Pt(0) Then
                ' Shift Big Numbers to The Right
                thisList(j) = thisList(j - 1)
                j = j - 1
            ElseIf thisList(j - 1)(0) = Pt(0) And thisList(j - 1)(1) > Pt(1) Then
                ' Shift Big Numbers to The Right
                thisList(j) = thisList(j - 1)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        thisList(j) = Pt
    Next
End Sub

Rem ---
Rem --- Remove item at index from list and resize list
Rem --- all item indexes above index are reset
Rem ---
Public Sub Remove(index As Long)
    Dim i As Long
    
    i = index
    If index < lsize Then
        i = index + 1
        While i < lsize
            thisList(i - 1) = thisList(i)
            i = i + 1
        Wend
    End If
    
    If index <= lsize Then
        lsize = lsize - 1
        ReDim Preserve thisList(1 To lsize)
    End If
End Sub

Rem ---
Rem --- remove all indexes in indexes array from list
Rem ---
Public Sub RemoveList(ByRef indexes() As Long)
    Dim i As Long
    Dim j As Long
    
    For j = 1 To UBound(indexes) - 1
        For i = indexes(j) To indexes(j + 1)
            If i < lsize Then
                thisList(i - (j - 1)) = thisList(i + 1)
            End If
        Next i
    Next j
    
    'move remaining entries to end of list
    j = UBound(indexes)
    For i = indexes(j) To lsize - 1
        thisList(i - (j - 1)) = thisList(i + 1)
    Next i
    
    lsize = lsize - j
    ReDim Preserve thisList(1 To lsize)
End Sub

Rem ---
Rem --- remove duplicate points - must be called after sort
Rem ---
Public Sub RemoveDupXY()
    Dim arr() As Long
    Dim i As Long
    Dim j As Long
    
    j = 0
    i = 2
    
    'get duplicate indexes
    For i = 2 To lsize
        If thisList(i - 1)(0) <= thisList(i)(0) + myFuzz And thisList(i - 1)(0) > thisList(i)(0) - myFuzz And thisList(i - 1)(1) <= thisList(i)(1) + myFuzz And thisList(i - 1)(1) > thisList(i)(1) - myFuzz Then
            j = j + 1
            ReDim Preserve arr(1 To j)
            arr(j) = i
        End If
    Next i
    
    If j > 0 Then
        'remove the duplicates
        RemoveList arr
    End If
End Sub

Rem ---
Rem --- Insertion sort array of Long (index 1)
Rem ---
Private Sub sortArr(ByRef arr() As Long)
    Dim j, i, n, aux As Long
    ' Upper Bound of Array
    n = UBound(arr)
    For i = 2 To n
        j = i
        aux = arr(i)
        Do While j > 1
            If arr(j - 1) > aux Then
                ' Shift Big Numbers to The Right
                arr(j) = arr(j - 1)
                j = j - 1
                Exit Do
            End If
        Loop
        arr(j) = aux
    Next
End Sub

Rem ---
Rem --- check if val is in array (index 1)
Rem ---
Private Function IsInArr(val As Long, ByRef arr() As Long) As Boolean
    Dim i As Long
    Dim ret As Boolean
    
    sortArr arr
    
    ret = False
    i = 1
    While i <= UBound(arr) And Not ret
        If val = arr(i) Then
            ret = True
        End If
        i = i + 1
    Wend
    
    IsInArr = ret
End Function
