Attribute VB_Name = "mBasicExamples"
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mBasicExamples: Examples in the SpecsAndUse document for
' =============================== for the mBasic component.
'
' Arry: https://github.com/warbe-maker/VBA-Basics/blob/master/SpecsAndUse.md#example-arry
' ---------------------------------------------------------------------------

Private arrMy1DimArray As Variant
Private arrMy3DimArray As Variant

' ---------------------------------------------------------------------------
' Encapsulation of read from, write to arrMy1DimArray, Empty as explicit
' default for non-active items.
' ---------------------------------------------------------------------------
Private Property Get My1DimArray(ByVal m_indices As Variant) As Variant
    My1DimArray = Arry(arrMy1DimArray, m_indices, Empty)
End Property

Private Property Let My1DimArray(Optional ByVal m_indices As Variant, ByVal m_item As Variant)
    Arry(arrMy1DimArray, m_indices, Empty) = m_item
End Property

Private Sub AnyProc1()
' ---------------------------------------------------------------------------
' Usage example 1-dim array
' ---------------------------------------------------------------------------
    
    Set arrMy1DimArray = Nothing
    My1DimArray(10) = "any"          ' writes item number 10, items 1 to 9 remain Empty,
                                 ' establishes the array arrMy1DimArray with its first item
    Debug.Assert My1DimArray(10) = "any"
    My1DimArray(100) = "another"     ' writes item number 100, items 11 to 99 remain Empty
    Debug.Assert My1DimArray(100) = "another"
    My1DimArray(10) = "any-updated"  ' updates item number 10
    Debug.Assert My1DimArray(10) = "any-updated"
    Debug.Assert My1DimArray(2) = Empty
    
End Sub

Private Sub Example3DimArray1()
' ---------------------------------------------------------------------------
' Usage example 3-dim array, direct use of Arry service, un-Redim-ed
' ---------------------------------------------------------------------------
    
    Dim s As String
    Dim l As Long
    
    Set arrMy3DimArray = Nothing
    Debug.Assert ArryDims(arrMy3DimArray) = 0
    
    '~~ Write first item thereby establishing the array
    Arry(arrMy3DimArray, "5,5,2") = "Item(5,5,2)"
    s = Arry(arrMy3DimArray, "5,5,2")
    Debug.Assert s = "Item(5,5,2)"
    Debug.Assert ArryDims(arrMy3DimArray) = 3
    Debug.Assert ArryItems(arrMy3DimArray, True) = 1
    
    Arry(arrMy3DimArray, "6,1,3") = "Item(6,1,3)"        ' writes subsequent item, expand upper bound of dim 1 and 3
    s = Arry(arrMy3DimArray, "6,1,3")
    Debug.Assert s = "Item(6,1,3)"
    Debug.Assert ArryItems(arrMy3DimArray, True) = 2
    l = ArryItems(arrMy3DimArray) ' 7 x 6 x 4 = 168 (Base 0)
    Debug.Assert l = 168
      
End Sub

Private Sub Example3DimArray2()
' ---------------------------------------------------------------------------
' Usage example 3-dim array, direct use of Arry service, pre-Redim-ed
' ---------------------------------------------------------------------------
    
    Dim s As String
    Dim l As Long
    
    Set arrMy3DimArray = Nothing
    ReDim arrMy3DimArray(2 To 3, 1 To 4, 1 To 3)
    Debug.Assert ArryDims(arrMy3DimArray) = 3
    Debug.Assert ArryItems(arrMy3DimArray) = 24 ' 2 x 4 x 3
    
    '~~ Write first item thereby establishing the array
    Arry(arrMy3DimArray, "5,5,2") = "Item(5,5,2)"
    s = Arry(arrMy3DimArray, "5,5,2")
    Debug.Assert s = "Item(5,5,2)"
    Debug.Assert ArryDims(arrMy3DimArray) = 3
    Debug.Assert ArryItems(arrMy3DimArray, True) = 1    ' count non-empty only
    l = ArryItems(arrMy3DimArray) ' 4 x 5 x 3 = 60
    Debug.Assert l = 60
    
    Arry(arrMy3DimArray, "6,1,3") = "Item(6,1,3)"        ' writes subsequent item, expands upper bound of dim 1
    s = Arry(arrMy3DimArray, "6,1,3")
    Debug.Assert s = "Item(6,1,3)"
    Debug.Assert ArryItems(arrMy3DimArray, True) = 2
    l = ArryItems(arrMy3DimArray) ' 5 x 5 x 3 = 75
    Debug.Assert l = 75
      
End Sub

