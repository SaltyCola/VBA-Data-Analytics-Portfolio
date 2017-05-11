VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TC_SummaryColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=================================================================================='
'=================================== Properties ==================================='
'=================================================================================='

Private pTitle As String 'The summary column's Title
Private pOpIndex As Integer 'The operation index of the Title
Private pNumberOfCategories As Integer 'The number of category rows in the summary table
Private pCategoryList() As TC_SummaryCategory 'The list of Category Row objects (which contain category title, color, and total)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Summary Column's Title
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get Title() As String
    Title = pTitle
End Property
Public Property Let Title(Value As String)
    pTitle = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Summary Column's Linked Op Index
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get OpIndex() As Integer
    OpIndex = pOpIndex
End Property
Public Property Let OpIndex(Value As Integer)
    pOpIndex = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Number of Categories in Summary Table
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get NumberOfCategories() As Integer
    NumberOfCategories = pNumberOfCategories
End Property
Public Property Let NumberOfCategories(Value As Integer)
    pNumberOfCategories = Value
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Array: Category object list
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get CategoryList(ByVal indexSC As Integer) As TC_SummaryCategory
    Set CategoryList = pCategoryList(indexSC)
End Property
Public Property Let CategoryList(ByVal indexSC As Integer, Value As TC_SummaryCategory)
    Set pCategoryList = Value(indexSC)
End Property



'=================================================================================='
'=================================== Methods ======================================'
'=================================================================================='

Public Sub InitializeCategories()
'Must have Number of Categories set before running this method.

    Dim i As Integer 'iterator
    
    'iterate for number of categories
    For i = 1 To pNumberOfCategories
    
        'set category list size
        ReDim Preserve pCategoryList(1 To pNumberOfCategories)
        
        'initialize category object
        Set pCategoryList(i) = New TC_SummaryCategory
    
    Next i

End Sub
