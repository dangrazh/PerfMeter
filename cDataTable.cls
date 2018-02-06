VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDataTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1
'Option Compare Text

'# <author> Daniel Grass
'# <mail> dani.grass@bluewin.ch

'#Region
'# Public Subs, Functions and Properties
'#======================================================================================================================
'# Accessible in this class
'#======================================================================================================================
'     |> Get | --- About                :: Returns description of the class.
'     |> --------- CreateLogFile        :: Creates a logfile for output and stores it in the specified [storageDirectory].
'     |> Get | Let DirectoryPath        :: Returns the directory location of the current logfile.
'     |> Get | --- Name                 :: Returns name of the class.
'     |> Get | --- Version              :: Returns version string for the class [e.g. #.# (year)].


'#======================================================================================================================
'# References
'#======================================================================================================================
'Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal bytes As Long)

#If Win64 Then
    Private Const PTR_LENGTH As Long = 8
    'Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    'Private Declare PtrSafe Function VarPtrArray Lib "VBE7" Alias "VarPtr" (ByRef Var() As Any) As LongPtr
    Private Declare PtrSafe Sub Mem_Copy Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
    Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal length As Long, ByVal Fill As Byte)
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)
#Else
    Private Const PTR_LENGTH As Long = 4
    'Public Declare Function GetTickCount Lib "kernel32" () As Long
    'Private Declare Function VarPtrArray Lib "VBE7" Alias "VarPtr" (ByRef Var() As Any) As LongPtr
    Private Declare Sub Mem_Copy Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
    Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal length As Long, ByVal Fill As Byte)
    Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)
#End If


'#======================================================================================================================
'# Dependencies to other classes
'#======================================================================================================================
' cIndex                                             :: is a single Index of a data column
' cHashTable (directly and thourgh cIndex as well)   :: is used to access column indices & as underlying data structure
'                                                       to unique indices
' cBPlusTree (thourgh cIndex)                        :: is used as underlying data structure to non-unique indices


'#======================================================================================================================
'# Application Constants, Enumerations & Types
'#======================================================================================================================
Const C_WANT_FREE_PERCENT = 0.1   ' 0.1 translates in 10% free space
Const C_MIN_FREE = 10             ' Min unused space when resizing

Private Const C_CUT_OFF As Long = 100 'move into Insertionsort if list (or sublist) has less 100 items
Private Const C_Name As String = "cDataTable.cls"

Private Enum ResizeType
    IncludingWantFreeSpace = 1
    WithoutEmptyTrailRecords = 2
End Enum

Public Enum MatchType
    Equals = 1
    MatchesPattern = 2
    DoesNotMatchPattern
    GreaterThen
    GreaterThenOrEqual
    LessThen
    LessThenOrEqual
End Enum


Private Type ErrorCode
    errNumber As Long
    errDescrption As String
End Type
    
    
'#======================================================================================================================
'# API Constants, Enumerations & Types
'#======================================================================================================================

'Type Declarations needed for SafeArray hacks

'The bounds of the SafeArray
Private Type SAFEARRAYBOUND
    cElements    As Long
    lLbound      As Long
End Type

Private Type SAFEARRAY1D
    cDims           As Integer
    fFeatures       As Integer
    cbElements      As Long
    cLocks          As Long
    pvData          As Long
    Bounds(0 To 0)  As SAFEARRAYBOUND
End Type

Private Type SAFEARRAY2D
    cDims           As Integer
    fFeatures       As Integer
    cbElements      As Long
    cLocks          As Long
    pvData          As Long
    Bounds(0 To 1)  As SAFEARRAYBOUND
End Type

'Constants needed for SafeArray hacks
Private Const FADF_AUTO = &H1            'The array is in the Stack
Private Const FADF_STATIC = &H2          'A static array
Private Const FADF_EMBEDDED = &H4        'The field is embedded in a structure
Private Const FADF_FIXEDSIZE = &H10      'The bounds of the array are fixed
Private Const FADF_RECORD = &H20         'The array contains records
Private Const FADF_HAVEIID = &H40        'The array is identifyable over a IID-interface
Private Const FADF_HAVEVARTYPE = &H80    'Field of type VT
Private Const FADF_BSTR = &H100          'String array
Private Const FADF_UNKNOWN = &H200       'Array with interface IUnknown
Private Const FADF_DISPATCH = &H400      'Array with interface IDispatch
Private Const FADF_VARIANT = &H800       'Array of type Variant


'#======================================================================================================================
'# Private Variables
'#======================================================================================================================
Private m_List() As Variant                     ' The list array.
Private m_ListTransposed() As Variant           ' The transposed list array
Private m_NumCols As Long                       ' Number of Columns
Private m_NumItems As Long                      ' Last index in use.
Private m_ArraySize As Long                     ' Size of the list array.
Private m_ShrinkWhen As Long                    ' Shrink if m_NumItems < this.
Private m_GarbageCount As Long                  ' The number of garbage entries.
Private m_MaxGarbage As Long                    ' Collect when m_GarbageCount > this.
Private m_FirstGarbage As Long                  ' The index of the 1st record containing garbage
Private m_LastGarbage As Long                   ' The index of the last record containing garbage
Private m_CollectGarbage As Boolean             ' Flag to determine if we should collect garbage
Private m_Columns() As Variant                  ' Array containing the column headers
Private m_SortedColumn As Long                  ' Contains the column index of the sorted column
Private m_Bookmark As Long                      ' Bookmark pointing to the current record
Private m_HasHeaders As Boolean                 ' Indicates if column headers were provided
Private m_IsDirty As Boolean                    ' Indicates if the array needs cleanup and a refresh of the transposed list
Private m_BOF As Boolean                        ' Indicates that the cursor has reached the beginning of the table
Private m_EOF As Boolean                        ' Indicates that the cursor has reached the end of the table
Private m_Errors(1 To 7) As ErrorCode           ' List of all error codes

'#Region
'#======================================================================================================================
'# Class Initialization, Termination & Properties
'#======================================================================================================================


Private Sub Class_Initialize()
' ************************************************
' Class constructor.
' ************************************************

    'Debug.Print "|> Initializing:= " & Me.Name
    
    'Set list to minimal size
    ReDim m_List(1, 1)
    
    'initialze the boolean states
    m_HasHeaders = False
    m_IsDirty = False
    m_CollectGarbage = True
    m_EOF = False
    m_BOF = False

    
        
    'initialize any default / start values
    m_NumCols = 1
    m_SortedColumn = -1
    m_Bookmark = -1
    
    'initialize error codes
    m_Errors(1).errNumber = 513:  m_Errors(1).errDescrption = "List is not empty. Range can only be added to empty list!"
    m_Errors(2).errNumber = 514:  m_Errors(2).errDescrption = "The list of Header Names does not match the number of columns you specified!"
    m_Errors(3).errNumber = 515:  m_Errors(3).errDescrption = "The no of items in the input record does not match the number of columns in the data table!"
    m_Errors(4).errNumber = 516:  m_Errors(4).errDescrption = "Assigning a value to an item out of boundaries of the data table is not possible!"
    m_Errors(5).errNumber = 517:  m_Errors(5).errDescrption = "The column name you specified could not be found in the data table!"
    m_Errors(6).errNumber = 518:  m_Errors(6).errDescrption = "Accessing an item needs either a row number or an active cursor from one of the Rec... methods"
    m_Errors(7).errNumber = 519:  m_Errors(7).errDescrption = "The two defined tables do not have the same data structure. Cannot append two tables with different data strucutres!"
    

End Sub


Private Sub Class_Terminate()
' ************************************************
' Class destructor.
' ************************************************
    
    Dim ptrList As Long
    Dim ptrListTransposed As Long

    'Debug.Print "|> Terminating:= " & Me.Name
    
    
    
    'empty the list
    ReDim m_List(1, 1)
    ReDim m_ListTransposed(1, 1)


End Sub

Public Property Get Version() As String
' ************************************************
' Version string of the current class.
' Contains a list of (historical) changes to the class within the comments of the procedure.
' ************************************************
   Version = "Version 1.0 (2017)"  'Initial release

End Property

Public Property Get About() As String
' ***********************************************
' String that describes the current class.
' ***********************************************
    About = "Data Table Class providing advanced array based data handling. Version: " & Me.Version & "." & VBA.vbCrLf & VBA.vbCrLf
    About = About & "For additional details please contact the author."
End Property

Public Property Get Name() As String
' ***********************************************
' Returns the name of the class.
' ***********************************************
    Name = C_Name
End Property

Public Property Get Headers() As Collection
' ***********************************************
' Returns the headers of the table as collection
' ***********************************************

    Dim cHeaders As New Collection
    Dim i As Long

    For i = LBound(m_Columns) To UBound(m_Columns)
        cHeaders.Add m_Columns(i)
    Next i
    
    Set Headers = cHeaders

End Property


Public Property Get HasHeaders() As Boolean
' ***********************************************
' Returns the HasHeaders state of the class.
' ***********************************************

        HasHeaders = m_HasHeaders

End Property

'Public Property Let HasHeaders(bValue As Boolean)
' ***********************************************
' Sets the HasHeaders state of the class.
' ***********************************************

'        m_HasHeaders = bValue

'End Property

Public Property Get NumItems() As Long
' ***********************************************
' Return the number of items.
' ***********************************************

    NumItems = m_NumItems

End Property

Private Property Get ArraySize() As Long
' ***********************************************
' Return the size of the array.
' ***********************************************

    ArraySize = m_ArraySize

End Property

Public Property Get NumCols() As Long
' ***********************************************
' Return the number of columns.
' ***********************************************

        NumCols = m_NumCols

End Property

Public Property Get GarbageCollection() As Boolean
' ***********************************************
' Returns the CollectGarbage state of the class.
' ***********************************************

    GarbageCollection = m_CollectGarbage

End Property

Public Property Let GarbageCollection(value As Boolean)
' ***********************************************
' Sets the CollectGarbage state of the class.
' ***********************************************
    
    If value = False Then
        'Set the state to false
        m_CollectGarbage = False
     Else
        'set the state to true
        m_CollectGarbage = True
        'If table is dirty, collect garbage
        If m_IsDirty = True Then
            CollectGarbage
        End If
    End If

End Property



Public Property Get SADescrPtr() As Long
' ***********************************************
' Return the pointer to the SafeArray Descriptor
' ***********************************************


    SADescrPtr = getSafeArrayDescrPtr(m_List)

End Property

Public Property Get SAStructPtr() As Long
' ***********************************************
' Return the pointer to the SafeArray Descriptor
' ***********************************************


    SAStructPtr = getSafeArrayStructPtr(m_List)

End Property


Public Property Get RsEOF() As Boolean
' ***********************************************
' Return the EOF State of the recordset
' ***********************************************

    RsEOF = m_EOF

End Property

Public Property Get RsBOF() As Boolean
' ***********************************************
' Return the EOF State of the recordset
' ***********************************************

    RsBOF = m_BOF

End Property

Public Property Get RsBookmark() As Long
' ***********************************************
' Return the pointer to the current record within the recordset
' ***********************************************

    RsBookmark = m_Bookmark

End Property

Public Property Let RsBookmark(value As Long)
' ***********************************************
' Set the pointer of the current record within the recordset
' to a specific position
' ***********************************************

    m_Bookmark = value

End Property


Public Property Get ItemName(Index As Variant) As String
' ***********************************************
' Return the name of the Item (Column Header)
' ***********************************************

    ItemName = GetColumnName(Index)

End Property


Public Property Get Item(Index As Variant, Optional Row As Long = -1) As Variant
' ***********************************************
' Return a given item at index / position
' ***********************************************

    If Row = -1 And m_Bookmark = -1 Then
        'raise error as we need either of the information
        ' For class errors, you add vbObjectError to the error number.
        Err.Raise vbObjectError + m_Errors(6).errNumber, Me.Name, m_Errors(6).errDescrption
     ElseIf Row > -1 Then
        'user provided a position, use this one
        Item = ItemGet(Row, CStr(Index))
     ElseIf m_Bookmark > -1 Then
        'we have a valid bookmark so let's use this
        Item = ItemGet(m_Bookmark, CStr(Index))
     Else
        MsgBox "Unhandled condition in Property Get 'Item'", vbCritical
        Exit Property
    End If


End Property


Public Property Let Item(Index As Variant, Optional Row As Long = -1, value As Variant)
' ***********************************************
' Set the value of a given item at position / index
' ***********************************************

    If Row = -1 And m_Bookmark = -1 Then
        'raise error as we need either of the information
        ' For class errors, you add vbObjectError to the error number.
        Err.Raise vbObjectError + m_Errors(6).errNumber, Me.Name, m_Errors(6).errDescrption
     ElseIf Row > -1 Then
        'user provided a position, use this one
            Call ItemSet(Row, CStr(Index), value)
     ElseIf m_Bookmark > -1 Then
        'we have a valid bookmark so let's use this
        Call ItemSet(m_Bookmark, CStr(Index), value)
     Else
        MsgBox "Unhandled condition in Property Let 'Item'", vbCritical
        Exit Property
    End If

End Property

Public Property Get Record(Optional Row As Long = -1) As Variant
' ***********************************************
' Return an record from the list by position.
' ***********************************************

    Dim iCol As Long
    Dim tmpRec() As Variant
    Dim position As Long
        
        
    If Row = -1 And m_Bookmark = -1 Then
        'raise error as we need either of the information
        ' For class errors, you add vbObjectError to the error number.
        Err.Raise vbObjectError + m_Errors(6).errNumber, Me.Name, m_Errors(6).errDescrption
     ElseIf Row > -1 Then
        'user provided a position, use this one
        position = Row
     ElseIf m_Bookmark > -1 Then
        'we have a valid bookmark so let's use this
        position = m_Bookmark
     Else
        MsgBox "Unhandled condition in Property Get 'Record'", vbCritical
        Exit Property
    End If
        
        
        
    If position < 1 Or position > m_NumItems Then
        ' Out of bounds. Return Null.
        Record = Null
     Else

        ReDim tmpRec(1 To m_NumCols)
        
        For iCol = 1 To m_NumCols
                ' Return the record.
                tmpRec(iCol) = m_List(iCol, position)
        Next iCol
        
        Record = tmpRec
                
    End If

End Property

Public Property Let Record(Row As Long, Record As Variant)
' ***********************************************
' Update a record in the list by position.
' ***********************************************

    Dim iCol As Long
    Dim tmpRec() As Variant
    Dim position As Long
        
    
    If Row = -1 And m_Bookmark = -1 Then
        'raise error as we need either of the information
        ' For class errors, you add vbObjectError to the error number.
        Err.Raise vbObjectError + m_Errors(6).errNumber, Me.Name, m_Errors(6).errDescrption
     ElseIf Row > -1 Then
        'user provided a position, use this one
        position = Row
     ElseIf m_Bookmark > -1 Then
        'we have a valid bookmark so let's use this
        position = m_Bookmark
     Else
        MsgBox "Unhandled condition in Property Get 'Record'", vbCritical
        Exit Property
    End If
        
        
    If position < 1 Or position > m_NumItems Then
        ' Out of bounds, raise error
        ' The range 513-65535 is available for user errors.
        ' For class errors, you add vbObjectError to the error number.
        Err.Raise vbObjectError + m_Errors(4).errNumber, Me.Name, m_Errors(4).errDescrption
     Else

        If UBound(Record) = UBound(m_List, 1) Then
            For iCol = 1 To m_NumCols
                ' Set the record values.
                m_List(iCol, position) = Record(iCol)
            Next iCol
        Else
            'raise error -> Not the same no of column Names as NoOfColumns specified
            ' The range 513-65535 is available for user errors.
            ' For class errors, you add vbObjectError to the error number.
            Err.Raise vbObjectError + 515, Me.Name, "The no of items in the input record does not match the number of columns in the data table!"
        End If
                
    End If

    

End Property


'#Region
'#======================================================================================================================
'# Table Definition, Index Handling and Statistics
'#======================================================================================================================
Public Sub DefineTable(NoOfColumns As Long, Optional ColumnHeaders As String = "n/a", Optional NoOfRows As Long = 1)
' ***********************************************
' Define structure of array and create Header List
' ***********************************************

        Dim i As Long
        Dim aTmp As Variant

        'Size the list
        ReDim m_List(1 To NoOfColumns, 1 To NoOfRows)
        m_ArraySize = NoOfRows
        
        'fill in the Headers
        ReDim m_Columns(1 To NoOfColumns)
        
        'Name the Fields Field<n> if no field list is provided
        If ColumnHeaders = "n/a" Or InStr(ColumnHeaders, ",") = 0 Then
                For i = 1 To NoOfColumns
                        m_Columns(i) = "Field" & i
                Next i
            
                m_NumCols = NoOfColumns
            
         ElseIf InStr(ColumnHeaders, ",") > 0 Then
                aTmp = Split(ColumnHeaders, ",")
                If UBound(aTmp) + 1 = NoOfColumns Then
                    For i = 1 To UBound(aTmp) + 1
                            m_Columns(i) = Trim(aTmp(i - 1))
                    Next i
                    
                    m_NumCols = NoOfColumns
                    m_HasHeaders = True
                    
                Else
                    'raise error -> Not the same no of column Names as NoOfColumns specified
                    ' The range 513-65535 is available for user errors.
                    ' For class errors, you add vbObjectError to the error number.
                    Err.Raise vbObjectError + 514, Me.Name, "The list of Header Names does not match the number of columns you specified!"
                End If
        End If
        
End Sub

'Public Sub AnalyzeDataTypes()
' ***********************************************
' Do random spot checks of data contents of each column
' and determine the data type of the column
' ***********************************************

        '§§§ Future feature §§§

'End Sub


'Public Sub AddIndex(stColumnName as String)
' ***********************************************
' Create an index on a given column
' ***********************************************

        '§§§ Future feature §§§

'End Sub


'Public Sub RefreshIndices()
' ***********************************************
' Refresh all indices in the table
' ***********************************************

        '§§§ Future feature §§§

'End Sub

'Public Sub RunStats()
' ***********************************************
' Refresh all stats and indices
' ***********************************************

        '§§§ Future feature §§§

'End Sub


'#Region
'#======================================================================================================================
'# Data handling (add / remove) and list maintenance
'#======================================================================================================================
        
Public Sub TruncateTable(AskForConfirmation As Boolean)

    Dim retVal As Integer

    If AskForConfirmation = True Then
        MsgBox "This will delete all data from the table! Do you really want to continue?", vbQuestion + vbYesNoCancel, "Truncating Table"
        If retVal = vbNo Or retVal = vbCancel Then
            Exit Sub
        End If
    End If
        
    'clean the list by resizing it to a single record
    ReDim m_List(1 To m_NumCols, 1 To 1)
    'reset the number of items to 0
    m_NumItems = 0
    'reset the array size to 1
    m_ArraySize = 1
    

End Sub


Private Sub ItemSet(ByVal position As Long, ByVal Index As String, value As Variant)
' ***********************************************
' Set the value of an item in the list by position.
' ***********************************************

    Dim iCol As Long
    'Dim bColFound As Boolean: bColFound = False

    If position < 1 Or position > m_NumItems Then
        ' Out of bounds, raise error
        ' The range 513-65535 is available for user errors.
        ' For class errors, you add vbObjectError to the error number.
        Err.Raise vbObjectError + m_Errors(4).errNumber, Me.Name, m_Errors(4).errDescrption
    Else
    
        iCol = GetColumnIndex(Index)
        If iCol = -1 Then
            'Column Name not found, raise error
            ' The range 513-65535 is available for user errors.
            ' For class errors, you add vbObjectError to the error number.
            Err.Raise vbObjectError + m_Errors(5).errNumber, Me.Name, m_Errors(5).errDescrption
         Else
            ' Set the value of the item
            m_List(iCol, position) = value
        End If
                
    End If
        
End Sub
        
             
        
Public Sub RecordAdd(ByVal aRecord As Variant)
' ***********************************************
' Add a record to the list
' ***********************************************

        Dim iCol As Long

 
        'Check if no of items in the record is in line with the number of colums in the list
        If UBound(aRecord) = UBound(m_List, 1) Then
        
                'increase the row index
                m_NumItems = m_NumItems + 1
                'resize the list if required
                If m_NumItems > m_ArraySize Then ResizeList (IncludingWantFreeSpace)

                'add the record to the list
                For iCol = 1 To m_NumCols
                        m_List(iCol, m_NumItems) = aRecord(iCol)
                Next iCol
                
                'mark data table as dirty
                m_IsDirty = True
                
         Else
         
                'raise error -> Not the same no of column Names as NoOfColumns specified
                ' The range 513-65535 is available for user errors.
                ' For class errors, you add vbObjectError to the error number.
                Err.Raise vbObjectError + 515, Me.Name, "The no of items in the input record does not match the number of columns in the data table!"
                
        End If
        
End Sub


Public Sub RecordRemove(ByVal position As Long)
' ***********************************************
' Remove an item from the list.
' ***********************************************

    Dim iCol As Long

    '*** this approach is too slow if the table has a lot of columns ****
    ' set the record to Null
    For iCol = 1 To m_NumCols
            m_List(iCol, position) = "#DELETED"
    Next iCol

'    'mark the 1st element of the record as deleted (faster approach)
'    m_List(1, position) = "#DELETED"
        
    'increase the garbage counter
    m_GarbageCount = m_GarbageCount + 1
    
    'Update the 1st and the last garbage if required
    If position < m_FirstGarbage Or m_FirstGarbage = 0 Then
        m_FirstGarbage = position
    End If
    
    If position > m_LastGarbage Or m_LastGarbage = 0 Then
        m_LastGarbage = position
    End If
    
    'Check if garbage has to be collected and do so if yes
    If m_CollectGarbage = True Then
        'collect garbage if required
        If m_GarbageCount > m_MaxGarbage Then
            CollectGarbage
        End If
    End If

    'mark data table as dirty
    m_IsDirty = True
        
End Sub


Public Sub LoadRange(InputTable As Range, TableHasHeaders As Boolean)
' ***********************************************
' Add a cell range to the empty table
' ***********************************************


        Dim iCol As Long
        Dim iColCount As Long
        
        Dim rHeader As Range
        Dim cField As Range
        Dim iRowStart As Long: iRowStart = InputTable.Row
        Dim iColStart As Long: iColStart = InputTable.Column
        Dim iCols As Long: iCols = InputTable.Columns.Count
        Dim iRows As Long: iRows = InputTable.Rows.Count
        
        Dim ListTemp() As Variant
        Dim ptrList As Long
        Dim ptrListTransposed As Long
        Dim ptrListTemp As Long
            
        Dim want_free As Long
        
        
        'Dim mTimer As New cStopWatch

        If m_NumItems > 0 Then
                'raise error -> only allowed with empty list
                ' The range 513-65535 is available for user errors.
                ' For class errors, you add vbObjectError to the error number.
                Err.Raise vbObjectError + 513, Me.Name, "List is not empty. Range can only be added to empty list!"
        End If
        
        'Performance Improvement: if has headers, take 1st row for header creation and then reduce range size by 1st row
        'this eliminates the need to call "CleanUpTable" (i.e. collect the garbage) at the end of the procedure
        iColCount = iCols
        ReDim m_Columns(1 To iColCount)
        If TableHasHeaders = True Then
                For iCol = 1 To iColCount
                        m_Columns(iCol) = InputTable.Cells(1, (iColStart - 1) + iCol).value
                Next iCol
                Set InputTable = InputTable.Offset(1, 0).Resize(iRows - 1, iCols)
        Else
                For iCol = 1 To iColCount
                        m_Columns(iCol) = "Field" & iCol
                Next iCol
        End If
        m_HasHeaders = True
        
        
        'Now load the data range without header into the array
        'mTimer.Reset
        m_List = InputTable
        'Debug.Print "Assigning the input range to the variant array took: " & mTimer.MillisecondsElapsed & " milliseconds..."
        
        m_NumItems = UBound(m_List)
        m_ArraySize = UBound(m_List)
        m_NumCols = UBound(m_List, 2)
        
        
        Call Transpose
        
        Dim switch As Integer: switch = 2
        
        Debug.Print "Switch is: " & switch
        
        If switch = 1 Then
            
            'mTimer.Reset
            'swap transposed list and list
            ListTemp = m_List
            m_List = m_ListTransposed
            m_ListTransposed = ListTemp
            'Debug.Print "Swapping the lists (classic) took: " & mTimer.MillisecondsElapsed & " milliseconds..."
            
         ElseIf switch = 2 Then
            
            'Performance Improvement: rather than copying the transposed list into the list to
            'transpose it back further down with 'CleanUpTable', simply swap the pointers to the base array info
            'To do so, we need three pointers to the descriptor of m_List, m_ListTransposed and a temp list we need
            'as a temporary buffer to store the pointer
            
            'mTimer.Reset
                        
            'Get the pointers to the SafeArray
            ptrList = getSafeArrayDescrPtr(m_List)
            ptrListTransposed = getSafeArrayDescrPtr(m_ListTransposed)
            ptrListTemp = getSafeArrayDescrPtr(ListTemp)
            
            'Swap list Transposed and List
            Call CopyMemory(ByVal ptrListTemp, ByVal ptrList, 4)
            Call CopyMemory(ByVal ptrList, ByVal ptrListTransposed, 4)
            Call CopyMemory(ByVal ptrListTransposed, ByVal ptrListTemp, 4)
            
            'clean pointer of temp list to avoid crash on extiting the programm
            Call CopyMemory(ByVal ptrListTemp, 0&, 4&)
            
            'Debug.Print "Swapping the lists (CopyMemory) took: " & mTimer.MillisecondsElapsed & " milliseconds..."
            
        End If
        
        'Performance Improvement: rather than run 'CleanUpTable' to get rid of the deleted 1st row and to transpose the table,
        'set the m_IsDirty flag manually (no need to transpose as earlier on the two arrays have been swapped anyway)
        'Call CleanUpTable
        m_IsDirty = False
        
        
        '*****
        'set ShrinkWhen and MaxGarbage properties
        '*****
        want_free = C_WANT_FREE_PERCENT * m_NumItems
        If want_free < C_MIN_FREE Then
            want_free = C_MIN_FREE
        End If
                
        ' We will shrink the array if SL_NumItems < SL_ShrinkWhen.
        m_ShrinkWhen = m_NumItems - want_free
        
        ' Collect garbage when there are more than
        ' want_free garbage items in the list.
        m_MaxGarbage = want_free

        
        
End Sub

Public Sub LoadDelimTextFile(filename As String, FieldDelimiter As String, RecordDelimiter As String, TableHasHeaders As Boolean, Optional TextQualifier As String)
' ***********************************************
' Add a delimiter File to the empty table
' ***********************************************

    Dim iCol As Long
    Dim iColCount As Long
    Dim lFn As Integer
        
    

    If m_NumItems > 0 Then
            'raise error -> only allowed with empty list
            ' The range 513-65535 is available for user errors.
            ' For class errors, you add vbObjectError to the error number.
            Err.Raise vbObjectError + 513, Me.Name, "List is not empty. Range can only be added to empty list!"
     Else
        
        lFn = VBA.FreeFile
        
        Open filename For Input As #lFn
        
        '1. Open the file, parse the records
        '2. Check if table has headers
        '3. Count the columns & Define the table structure
        '4. Read the headers first and add header info
        '5. loop through all records and add them to the table
    End If
        


End Sub

Public Sub AppendToTable(ByRef tSource As cDataTable)
' ***********************************************
' Append the data of a given cDataTable object to the list
' ***********************************************

    Dim lSourceColCount As Long
    Dim lTargetColCount As Long
    Dim lSturctureCheck As Long
    Dim lHeaderCheck As Long
    Dim lNewColCount As Long
    Dim arrFieldMapping() As Variant
    Dim i As Long
    Dim k As Long
    
    lSourceColCount = tSource.NumCols
    lTargetColCount = Me.NumCols
    
    ReDim arrFieldMapping(1 To lSourceColCount, 1 To 2)
    
    '§§§ UNDER CONSTRUCTION !!! §§§
    
    
    '1. Check if structure is the same
    If lSourceColCount = lTargetColCount Then
        'the tables have the same no of columns
        'let's check if they are in the same order
        If tSource.HasHeaders = True Then
            For i = 1 To lSourceColCount
                'create the source mapping in case needed
                arrFieldMapping(i, 1) = tSource.ItemName(i)
                If tSource.ItemName(i) = Me.ItemName(i) Then
                    'Source and Target Column are the same
                    lHeaderCheck = 1
                 Else
                    lHeaderCheck = 2
                End If
            Next i
            
            For i = 1 To lSourceColCount
                For k = 1 To lSourceColCount
                    If arrFieldMapping(i, 1) = Me.ItemName(i) Then
                        arrFieldMapping(i, 2) = Me.ItemName(i)
                        Exit For
                     Else
                        arrFieldMapping(i, 2) = "No Match"
                    End If
                        
               Next k
            
                'check if we found a match
                If arrFieldMapping(i, 2) = "No Match" Then
                    'raise error -> the column names do not match
                    ' For class errors, you add vbObjectError to the error number.
                    Err.Raise vbObjectError + m_Errors(7).errNumber, Me.Name, m_Errors(7).errDescrption

                End If
            Next i
        End If
                
        lSturctureCheck = 1
     ElseIf lSourceColCount > lTargetColCount Then
        'the source has more columns than the target
        lNewColCount = lSourceColCount
        lSturctureCheck = 2
     ElseIf lSourceColCount < lTargetColCount Then
        'the source has less columns, let's append what we have
        lSturctureCheck = 3
    End If
     
                 
    '2.


    '§§§ UNDER CONSTRUCTION !!! §§§


End Sub


Private Sub ResizeList(tyResize As ResizeType)
' ***********************************************
' Resize the list to have C_WANT_FREE_PERCENT unused entries. Set m_ShrinkWhen so we know to
' resize the array when the number of used entries is too small. Set this value to resize
' the array when there is more than twice the desired amount of memory free.
' ***********************************************

    Dim want_free As Long

    want_free = C_WANT_FREE_PERCENT * m_NumItems
    If want_free < C_MIN_FREE Then
        want_free = C_MIN_FREE
    End If
        
    If tyResize = IncludingWantFreeSpace Then
            m_ArraySize = m_NumItems + want_free
     Else
            m_ArraySize = m_NumItems
    End If
    
    ReDim Preserve m_List(1 To m_NumCols, 1 To m_ArraySize)

        
    ' We will shrink the array if SL_NumItems < SL_ShrinkWhen.
    m_ShrinkWhen = m_NumItems - want_free
    
    ' Collect garbage when there are more than
    ' want_free garbage items in the list.
    m_MaxGarbage = want_free
    
    'Debug.Print "Resizing." & vbCrLf & vbCrLf & _
        "Items:" & Str$(m_ArraySize) & vbCrLf & _
        "Unused: " & Str$(want_free) & vbCrLf & _
        "ShrinkWhen:" & Str$(m_ShrinkWhen) & vbCrLf & _
        "MaxGarbage:" & Str$(m_MaxGarbage)
                
End Sub

Private Sub CollectGarbage()
' ************************************************
' Collect the garbage.
' ************************************************

    Dim i As Long
    Dim good As Long
    Dim iCol As Long
    
    'Debug.Print "Collecting garbage..."
            
    good = 1 ' The first good item goes here.
    For i = 1 To m_NumItems
        ' If 1st field in the record is not garbage,
        ' move the record to its new location.
        If Not m_List(1, i) = "#DELETED" Then
            For iCol = 1 To m_NumCols
                m_List(iCol, good) = m_List(iCol, i)
            Next iCol
            good = good + 1
        End If
    Next i
    
    ' This is where the last good item is.
    m_NumItems = good - 1

    ' See if we should resize the list.
    If m_NumItems < m_ShrinkWhen Then ResizeList (IncludingWantFreeSpace)
        
End Sub


Private Sub CleanUpTable()
' ************************************************
' Clean up the data table to prepare for sorting etc:
' Collect the garbage + remove free tail records
' ************************************************
        'Collect the garbage
        CollectGarbage
        
        'resize List with option to remove the free tail records
        ResizeList (WithoutEmptyTrailRecords)

        'Transpose the list
        Transpose
    
        m_IsDirty = False

End Sub



'#Region
'#======================================================================================================================
'# Record Navigation
'#======================================================================================================================
Public Sub RsMoveFirst()
' ***********************************************
' Moves the bookmark to the 1st valid record
' ***********************************************

    Dim lPointer As Long
    
    m_EOF = False
    m_BOF = False
    
    For lPointer = LBound(m_List, 2) To m_NumItems
        If m_List(1, lPointer) <> "#DELETED" Then
            Exit For
        End If
    Next lPointer

    m_Bookmark = lPointer


End Sub

Public Sub RsMoveLast()
' ***********************************************
' Moves the bookmark to the last valid record
' ***********************************************

    Dim lPointer As Long
    
    m_EOF = False
    m_BOF = False
    
    
    For lPointer = m_NumItems To LBound(m_List, 2) Step -1
        If m_List(1, lPointer) <> "#DELETED" Then
            Exit For
        End If
    Next lPointer

    m_Bookmark = lPointer


End Sub

Public Sub RsMoveNext()
' ***********************************************
' Moves the bookmark to the next valid record
' ***********************************************

    Dim lPointer As Long
    Dim iIdx As Long
    
    iIdx = m_Bookmark
    
    If iIdx + 1 > m_NumItems Then
        'we are passing by the last record, set EOF flag
        m_EOF = True
     Else
    
        For lPointer = iIdx + 1 To m_NumItems
            If m_List(1, lPointer) <> "#DELETED" Then
                Exit For
            End If
        Next lPointer
    
        m_Bookmark = lPointer
        
    End If

End Sub

Public Sub RsMovePrevious()
' ***********************************************
' Moves the bookmark to the next valid record
' ***********************************************

    Dim lPointer As Long
    Dim iIdx As Long
    
    iIdx = m_Bookmark
    
    If iIdx - 1 < LBound(m_List, 2) Then
        'we are passing by the first record, set BOF flag
        m_BOF = True
     Else
       
        For lPointer = iIdx - 1 To LBound(m_List, 2) Step -1
            If m_List(1, lPointer) <> "#DELETED" Then
                Exit For
            End If
        Next lPointer
    
        m_Bookmark = lPointer
    End If
    
End Sub

Public Sub RsFind(Index As Variant, Match As MatchType, Criteria As Variant)
' ***********************************************
' Moves the bookmark to the next valid record
' ***********************************************

    Dim lPointer As Long
    Dim iIdx As Long
    
    'set the pointer to the beginning of the table or at the next record
    If m_Bookmark = -1 Then
        lPointer = 1
     Else
        If m_Bookmark + 1 > m_NumItems Then
            'last record reached, set EOF flag
            m_EOF = True
            Exit Sub
         Else
            RsMoveNext
            lPointer = m_Bookmark
        End If
    End If
    
    'Find the next record matching the criteria and set cursor to this record
    If Search(Criteria, Match, CStr(Index), lPointer) = False Then
        'item not found
        m_EOF = True
    End If
    
    '§§§
    '§§§ CONSTRUCTION SITE!! NEED TO FIGURE OUT HOW TO RESOLVE THE CRITERIA INTO A EXECUTABLE SEARCH!! §§§
    '§§§ TEMPORARY SOLUTION IS TO USE LINEAR SEARCH ON ONE COLUMN ONLY... §§§
    '§§§
    
'    For iIdx = lPointer To m_NumItems
'    Next iIdx

    '§§§
    '§§§ END OF CONSTRUCTION SITE
    '§§§


End Sub



'#Region
'#======================================================================================================================
'# Data Selection
'#======================================================================================================================

Private Function ItemGet(ByVal position As Long, ByVal Index As String) As Variant
' ***********************************************
' Return an item from the list by position.
' ***********************************************

    Dim iCol As Long
    'Dim bColFound As Boolean: bColFound = False

    If position < 1 Or position > m_NumItems Then
        ' Out of bounds. Return Null.
        ItemGet = Null
    Else
    
        iCol = GetColumnIndex(Index)
        
        If iCol = -1 Then
            'Column Name not found. Return Null.
            ItemGet = Null
         Else
            ' Return the item.
            ItemGet = m_List(iCol, position)
        End If
                
    End If
        
End Function


Public Function SelectData(FieldList As String, WhereClause As String, Optional GroupBy As String = "", Optional Having As String = "", Optional SortBy As String = "") As cDataTable
' ***********************************************
' Return a new data table object containing recrods
' with the specified Field List and matching the WhereClause
' ***********************************************

    Dim dtOut As cDataTable
    
    Set dtOut = New cDataTable

    '§§§ Future feature §§§
    
    '0. Parse all parameters into a parse tree
    '1. Define Output Datetable Structure based on Field List
    '2. Evaluate WhereClause
        '2a. Check if sorted / indexed
        '2b. Find relevant records
    '5. Add each relevent record to new data table
    '6. create new data table and Apply GroupBy clause
    '7. evaluate HavingClause -> remove records for which having clause is not true
    '8. If distinct -> remove duplicates from output
    '9. Order the table by the OrderBy clause
    
    'Assign the selection result to the function output
    SelectData = dtOut

End Function


'#Region
'#======================================================================================================================
'# Data Output
'#======================================================================================================================


Public Sub DumpToRange(TargetWorksheet As Worksheet, TargetCell As Range, Optional IncludeHeader As Boolean = True)
' ***********************************************
' Output the array to a range on a given worksheet
' This works with MS Excel only...
' ***********************************************

    Dim iRow As Integer: iRow = TargetCell.Row
    Dim iCol As Integer: iCol = TargetCell.Column
    Dim iColIdx As Integer
    
    If m_IsDirty = True Then
        Call CleanUpTable
        'Call Transpose -> not needed as included in CleanUpTable
    End If
    
    If IncludeHeader = True Then
        For iColIdx = 1 To m_NumCols
            TargetWorksheet.Cells(iRow, iCol + (iColIdx - 1)).value = m_Columns(iColIdx)
        Next iColIdx
        
        TargetWorksheet.Cells(iRow + 1, iCol).Resize(UBound(m_ListTransposed), UBound(m_ListTransposed, 2)) = m_ListTransposed
     Else
        TargetWorksheet.Cells(iRow, iCol).Resize(UBound(m_ListTransposed), UBound(m_ListTransposed, 2)) = m_ListTransposed
    End If

End Sub

Public Sub DumpToFile(TargetFile As String, Delimiter As String)
' ***********************************************
' Output the array to a text file applying any
' given delimiter
' ***********************************************
        
        Dim fn As Integer
        Dim stRecored As String
        Dim i As Long
        Dim ii As Long
        Dim stRecord As String
        
        If m_IsDirty = True Then
            Call CleanUpTable
        End If
        
    fn = VBA.FreeFile
        
    Open TargetFile For Output As #fn
        
        For i = 1 To UBound(m_ListTransposed)
                For ii = 1 To UBound(m_ListTransposed, 1)
                        'construct the record ...
                        If stRecord = "" Then
                                stRecord = m_ListTransposed(i, ii)
                         Else
                                stRecord = stRecord & Delimiter & m_ListTransposed(i, ii)
                        End If
                Next ii
          Print #fn, stRecord
          stRecord = ""
        Next i

        Close #fn
        
End Sub

'#Region
'#======================================================================================================================
'# Sorting
'#======================================================================================================================

Public Sub Sort(ParamArray SortOrder() As Variant)
' ************************************************
' Sorts the list with the hybrid mergesort
' The ParamArray SortOrder is expected to be the number
' of the column to sort and a sort direction (as "asc" / "desc").
' If the sort direction is omitted, ascending sort direction
' is choosen as the default.
' ************************************************

    Dim i As Long
    Dim iColToSort As Long
    Dim iCol As Long
    Dim aScratch() As Variant
    'Dim List() As Variant
        
    
    'List = m_List
    
    'Make sure we have a clean table if it's dirty
    If m_IsDirty = True Then
        'Collect the garbage
        CollectGarbage
        'resize List with option to remove the free tail records
        ResizeList (WithoutEmptyTrailRecords)
    End If
        
    For i = LBound(SortOrder) To UBound(SortOrder)
        
        
        'Only process numeric pointers to the columns to sort
        If IsTableColumn(SortOrder(i)) Then
            iColToSort = GetColumnIndex(SortOrder(i))
            'check if we are not in the last parameter, then check the sort direction
            If i + 1 <= UBound(SortOrder) Then
                If Not IsTableColumn(SortOrder(i + 1)) Then
                    If UCase(SortOrder(i + 1)) = "ASC" Then
                        ReDim aScratch(LBound(m_List) To UBound(m_List), LBound(m_List, 2) To UBound(m_List, 2))
                        Call MergeSort(m_List(), iColToSort, aScratch(), LBound(m_List, 2), UBound(m_List, 2))              '§§§ changed to 2nd dimension
                        m_SortedColumn = iColToSort
                     ElseIf UCase(SortOrder(i + 1)) = "DESC" Then
                        ReDim aScratch(LBound(m_List) To UBound(m_List), LBound(m_List, 2) To UBound(m_List, 2))
                        Call MergeSortDesc(m_List(), iColToSort, aScratch(), LBound(m_List, 2), UBound(m_List, 2))          '§§§ changed to 2nd dimension
                        m_SortedColumn = iColToSort
                     Else
                        MsgBox "Invalid sort direction supplied for column " & GetColumnName(iColToSort) & "!", vbCritical
                        Exit Sub
                    End If
                 Else
                    ReDim aScratch(LBound(m_List) To UBound(m_List), LBound(m_List, 2) To UBound(m_List, 2))
                    Call MergeSort(m_List(), iColToSort, aScratch(), LBound(m_List, 2), UBound(m_List, 2))                  '§§§ changed to 2nd dimension
                    m_SortedColumn = iColToSort
               End If
             Else
                'we are looking at the last parameter - only process if it is a numeric pointer to the column to sort
                If IsTableColumn(SortOrder(i)) Then
                    ReDim aScratch(LBound(m_List) To UBound(m_List), LBound(m_List, 2) To UBound(m_List, 2))
                    Call MergeSort(m_List(), iColToSort, aScratch(), LBound(m_List, 2), UBound(m_List, 2))                  '§§§ changed to 2nd dimension
                    m_SortedColumn = iColToSort
                End If
            End If
        End If
        
    Next i

    'Mark table as dirty, as we have changed the order of the items
    m_IsDirty = True

End Sub

Private Sub MergeSort(List() As Variant, ColToSort As Long, Scratch() As Variant, ByVal min As Long, ByVal max As Long)
' ************************************************
' Hybrid Mergesort with:
'   - Stops recursion when there are fewer than
'       C_CUT_OFF items in the list. Finishes with
'       InsertionSort.
' ************************************************
    
    Dim middle As Long
    Dim i1 As Long
    Dim i2 As Long
    Dim i3 As Long

    Dim iCols As Long
    
    'Determine the number of Columns of the list
    Dim iColLower As Long: iColLower = LBound(List)          '§§§ changed to 1st dimension
    Dim iColUpper As Long: iColUpper = UBound(List)          '§§§ changed to 1st dimension
    

    ' If the list has no more than C_CUT_OFF elements,
    ' finish it off with InsertionSort.
    If max - min < C_CUT_OFF Then
        InsertionSort List(), ColToSort, min, max
        Exit Sub
    End If
    
    ' Recursively sort the sublists.
    middle = max \ 2 + min \ 2
    MergeSort List(), ColToSort, Scratch(), min, middle
    MergeSort List(), ColToSort, Scratch(), middle + 1, max

    ' Merge the sorted lists.
    i1 = min        ' Index in list 1
    i2 = middle + 1 ' Index in list 2
    i3 = min        ' Index in merged list
    Do While i1 <= middle And i2 <= max
        If List(ColToSort, i1) <= List(ColToSort, i2) Then          '§§§ inverted i<n> and ColToSort
            For iCols = iColLower To iColUpper
                Scratch(iCols, i3) = List(iCols, i1)                '§§§ inverted i<n> and iCols
            Next iCols
            i1 = i1 + 1
        Else
            For iCols = iColLower To iColUpper
                Scratch(iCols, i3) = List(iCols, i2)                '§§§ inverted i<n> and iCols
            Next iCols
            i2 = i2 + 1
        End If
        i3 = i3 + 1
    Loop

    ' Empty out whichever list is not already empty.
    Do While i1 <= middle
        For iCols = iColLower To iColUpper
            Scratch(iCols, i3) = List(iCols, i1)                    '§§§ inverted i<n> and iCols
        Next iCols
        i1 = i1 + 1
       i3 = i3 + 1
    Loop
    Do While i2 <= max
        For iCols = iColLower To iColUpper
            Scratch(iCols, i3) = List(iCols, i2)                    '§§§ inverted i<n> and iCols
        Next iCols
        i2 = i2 + 1
        i3 = i3 + 1
    Loop

    ' Move the merged list back into list.
    For i3 = min To max
        For iCols = iColLower To iColUpper
            List(iCols, i3) = Scratch(iCols, i3)                    '§§§ inverted i<n> and iCols
        Next iCols
    Next i3
End Sub

Private Sub MergeSortDesc(List() As Variant, ColToSort As Long, Scratch() As Variant, ByVal min As Long, ByVal max As Long)
' ************************************************
' Mergesort with:
'   - Stops recursion when there are fewer than
'       C_CUT_OFF items in the list. Finishes with
'       InsertionSort.
' ************************************************
    
    Dim middle As Long
    Dim i1 As Long
    Dim i2 As Long
    Dim i3 As Long

    Dim iCols As Long
    
    'Determine the number of Columns of the list
    Dim iColLower As Long: iColLower = LBound(List)                      '§§§ changed to 1st dimension
    Dim iColUpper As Long: iColUpper = UBound(List)                      '§§§ changed to 1st dimension
    

    ' If the list has no more than C_CUT_OFF elements,
    ' finish it off with InsertionSort.
    If max - min < C_CUT_OFF Then
        InsertionSortDesc List(), ColToSort, min, max
        Exit Sub
    End If
    
    ' Recursively sort the sublists.
    middle = max \ 2 + min \ 2
    MergeSortDesc List(), ColToSort, Scratch(), min, middle
    MergeSortDesc List(), ColToSort, Scratch(), middle + 1, max

    ' Merge the sorted lists.
    i1 = min        ' Index in list 1
    i2 = middle + 1 ' Index in list 2
    i3 = min        ' Index in merged list
    Do While i1 <= middle And i2 <= max
        If List(ColToSort, i1) >= List(ColToSort, i2) Then              '§§§ inverted i<n> and ColToSort
            For iCols = iColLower To iColUpper
                Scratch(iCols, i3) = List(iCols, i1)                    '§§§ inverted i<n> and iCols
            Next iCols
            i1 = i1 + 1
        Else
            For iCols = iColLower To iColUpper
                Scratch(iCols, i3) = List(iCols, i2)                    '§§§ inverted i<n> and iCols
            Next iCols
            i2 = i2 + 1
        End If
        i3 = i3 + 1
    Loop

    ' Empty out whichever list is not already empty.
    Do While i1 <= middle
        For iCols = iColLower To iColUpper
            Scratch(iCols, i3) = List(iCols, i1)                        '§§§ inverted i<n> and iCols
        Next iCols
        i1 = i1 + 1
       i3 = i3 + 1
    Loop
    Do While i2 <= max
        For iCols = iColLower To iColUpper
            Scratch(iCols, i3) = List(iCols, i2)                        '§§§ inverted i<n> and iCols
        Next iCols
        i2 = i2 + 1
        i3 = i3 + 1
    Loop

    ' Move the merged list back into list.
    For i3 = min To max
        For iCols = iColLower To iColUpper
            List(iCols, i3) = Scratch(iCols, i3)                        '§§§ inverted i<n> and iCols
        Next iCols
    Next i3
End Sub


Private Sub InsertionSort(List() As Variant, ColToSort As Long, min As Long, max As Long)
' ************************************************
' Standard insertion sort algorythm
' ************************************************
    
    Dim iInnerLoop As Long
    Dim iOuterLoop As Long
    Dim next_item As Variant
    Dim next_item_rec As Variant
    
    Dim iCols As Long
    Dim iColLower As Long: iColLower = LBound(List)                             '§§§ changed to 1st dimension
    Dim iColUpper As Long: iColUpper = UBound(List)                             '§§§ changed to 1st dimension

    ReDim next_item_rec(iColLower To iColUpper)
    
    For iOuterLoop = min + 1 To max
        
        'this is the next item / record
        next_item = List(ColToSort, iOuterLoop)                                 '§§§ inverted iOuterLoop and ColToSort
        For iCols = iColLower To iColUpper
            next_item_rec(iCols) = List(iCols, iOuterLoop)                      '§§§ inverted iOuterLoop and ColToSort
        Next iCols

        'reorder the items
        For iInnerLoop = iOuterLoop To min + 1 Step -1
            If next_item < List(ColToSort, iInnerLoop - 1) Then                 '§§§ inverted iInnerLoop and ColToSort
                For iCols = iColLower To iColUpper
                    List(iCols, iInnerLoop) = List(iCols, iInnerLoop - 1)       '§§§ inverted iInnerLoop and iCols
                Next iCols
             Else
                Exit For
            End If
        Next iInnerLoop
        
        'insert the new item
        For iCols = iColLower To iColUpper
            List(iCols, iInnerLoop) = next_item_rec(iCols)                      '§§§ inverted iInnerLoop and iCols
        Next iCols
        
    Next iOuterLoop
End Sub


Private Sub InsertionSortDesc(List() As Variant, ColToSort As Long, min As Long, max As Long)
' ************************************************
' Standard insertion sort algorythm
' ************************************************
    
    Dim iInnerLoop As Long
    Dim iOuterLoop As Long
    Dim next_item As Variant
    Dim next_item_rec As Variant
    
    Dim iCols As Long
    Dim iColLower As Long: iColLower = LBound(List)                                         '§§§ changed to 1st dimension
    Dim iColUpper As Long: iColUpper = UBound(List)                                         '§§§ changed to 1st dimension

    ReDim next_item_rec(iColLower To iColUpper)
    
    For iOuterLoop = min + 1 To max
        
        'this is the next item / record
        next_item = List(ColToSort, iOuterLoop)                                             '§§§ inverted iOuterLoop and ColToSort
        For iCols = iColLower To iColUpper
            next_item_rec(iCols) = List(iCols, iOuterLoop)                                  '§§§ inverted iOuterLoop and iCols
        Next iCols

        'reorder the items
        For iInnerLoop = iOuterLoop To min + 1 Step -1
            If next_item > List(ColToSort, iInnerLoop - 1) Then                             '§§§ inverted iInnerLoop and ColToSort
                For iCols = iColLower To iColUpper
                    List(iCols, iInnerLoop) = List(iCols, iInnerLoop - 1)                   '§§§ inverted iInnerLoop and iCols
                Next iCols
             Else
                Exit For
            End If
        Next iInnerLoop
        
        'insert the new item
        For iCols = iColLower To iColUpper
            List(iCols, iInnerLoop) = next_item_rec(iCols)                                  '§§§ inverted iInnerLoop and iCols
        Next iCols
        
    Next iOuterLoop
    
End Sub




Private Sub QuickSort(cToSort As Long, Optional ByVal lngStart As Variant, Optional ByVal lngEnd As Variant)
' ***********************************************
' QuickSort-Algorithm
' ***********************************************

'§§§ currently not used in this class - however could be faster than the hybrid sort, §§§
'§§§ therefore here for eventual later use. §§§
'§§§ double check if this works for top-down or left-right oriented array §§§
'§§§ enhance with sort direction §§§

  ' Sort the entire array if the start / end boundaries are not defined
  If IsMissing(lngStart) Then lngStart = LBound(m_List)
  If IsMissing(lngEnd) Then lngEnd = UBound(m_List)

  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim H As Variant
  Dim x As Variant
  
  Dim lCStart As Long
  Dim lCEnd As Long
  
  lCStart = LBound(m_List, 2)
  lCEnd = UBound(m_List, 2)

  i = lngStart: j = lngEnd
  x = m_List((lngStart + lngEnd) / 2, cToSort)

  ReDim H(lCEnd)

  ' Split the array
  Do

    While (m_List(i, cToSort) < x): i = i + 1: Wend
    While (m_List(j, cToSort) > x): j = j - 1: Wend

    If (i <= j) Then
      ' Flip value pairs
      For k = lCStart To lCEnd
        H(k) = m_List(i, k)
      Next k
      For k = lCStart To lCEnd
        m_List(i, k) = m_List(j, k)
      Next k
      For k = lCStart To lCEnd
        m_List(j, k) = H(k)
      Next k
      i = i + 1: j = j - 1
    End If
  Loop Until (i > j)

  ' Recursion
  If (lngStart < j) Then QuickSort cToSort, lngStart, j
  If (i < lngEnd) Then QuickSort cToSort, i, lngEnd
  
End Sub

'#Region
'#======================================================================================================================
'# Searching
'#======================================================================================================================

Private Function Search(SearchFor As Variant, Match As MatchType, Index As String, Optional Start As Long = 1) As Boolean
' ************************************************
' Function to search the list
' Currently only calling the full table scan
' §§§ Future usage considering sorted colum, indexes and data types §§§
' ************************************************

    Dim lPointer As Long
    
    lPointer = LinearSearch(SearchFor, Match, Index, Start)

    If lPointer = 0 Then
        Search = False
     Else
        Search = True
        m_Bookmark = lPointer
    End If


End Function


Private Function LinearSearch(SearchFor As Variant, Match As MatchType, Index As String, Optional Start As Long = 1) As Long
' ************************************************
' Locate the item using exhaustive search.
' This function returns the pointer to the recrod, not the record itself.
' If item is not found, a 0-Pointer is returned.
' ************************************************

    Dim target As Variant: target = SearchFor
    Dim i As Long
    Dim iCol As Long
    Dim NumSearches As Long
    
    iCol = GetColumnIndex(Index)
    
    'NumSearches = 0
    
    For i = Start To m_NumItems
        'NumSearches = NumSearches + 1
        
        'Old code checking for equals only
        'If m_List(iCol, i) = target Then
        '    Exit For
        'End If
        
        'New code using compare function
        If Compare(m_List(iCol, i), target, Match) = True Then
            Exit For
        End If
        
    Next i
    
    If i > m_NumItems Then
        LinearSearch = 0    ' Not found.
    'Old Code checking for equals only
    'ElseIf m_List(iCol, i) <> target Then
    '    LinearSearch = 0    ' Not found.
    ElseIf Compare(m_List(iCol, i), target, Match) = False Then
         LinearSearch = 0    ' Not found.
    Else
        LinearSearch = i    ' Found.
    End If
    
End Function


'#Region
'#======================================================================================================================
'# Array handling
'#======================================================================================================================

Private Sub Transpose()
' ***********************************************
' Transpose the array using a SafeArray hack
' ***********************************************

    Dim i As Long
    Dim k As Long
    Dim m As Long
    Dim tmpArr() As Variant
    Dim acurOld() As Currency
    Dim acurNew() As Currency
    Dim lngCount As Long
    Dim lngSafeArrayList As Long
    Dim lngSafeArray As Long
    Dim lngData As Long
    Dim lngCountX As Long
    Dim lngCountY As Long
    Dim lngLBoundX As Long
    Dim lngLBoundY As Long
    Dim SafeArray2 As SAFEARRAY2D
    
    'Dim mTimer As New cStopWatch
    
    'mTimer.Reset
    
    '§§§ Performance Improvment: This might be faster by using CopyMemory at a later stage §§§
    'copy the data from m_List to m_ListTransposed
    m_ListTransposed = m_List
    'Debug.Print "HP Copying m_List to m_ListTransposed took: " & mTimer.MillisecondsElapsed & " milliseconds..."
    
    
    'Get the pointer to the SafeArray Base of m_ListTransposed
    lngSafeArray = getSafeArrayStructPtr(m_ListTransposed)
    
    'Copy the SafeArray base in to it's structure
    CopyMemory SafeArray2, ByVal lngSafeArray, Len(SafeArray2)

    With SafeArray2
        'Get the pointer to the actual data
        lngData = .pvData
        'Get the LBound and the no of elements for each dimension
        lngCountX = .Bounds(0).cElements
        lngLBoundX = .Bounds(0).lLbound
        lngCountY = .Bounds(1).cElements
        lngLBoundY = .Bounds(1).lLbound
    End With
    
    'Stop
    
    'mTimer.Reset
    
    'Get to no of elements of the matrix
    lngCount = (lngCountX * lngCountY)

    'We need to maintain the order of the elements, so we need to create a buffer for the data
    'Each element requires 16 Byte space, so twice a Currency.
    'This does not actually hold the data itself but rather the variant info with the
    'pointers to the data.
    ReDim acurOld(1 To 2, 1 To lngCount)
    ReDim acurNew(1 To 2, 1 To lngCount)
    
    'copy the headerinfo of the data to the buffer
    CopyMemory acurOld(1, 1), ByVal lngData, lngCount * 16
    
    'swap the headerinfo  of the data and write it into the buffer acurNew
    For i = 1 To lngCountY
        For k = 1 To lngCountX
            m = m + 1
            acurNew(1, m) = acurOld(1, (k - 1) * lngCountY + i)
            acurNew(2, m) = acurOld(2, (k - 1) * lngCountY + i)
        Next
    Next
    
    'Now copy the swaped headerinfos back to their original position
    CopyMemory ByVal lngData, acurNew(1, 1), lngCount * 16

    'Now we pretend that the 1st dimension has the no of elements of the 2nd and vice versa
    CopyMemory ByVal lngSafeArray + 16, lngCountY, 4
    CopyMemory ByVal lngSafeArray + 20, lngLBoundY, 4
    CopyMemory ByVal lngSafeArray + 24, lngCountX, 4
    CopyMemory ByVal lngSafeArray + 28, lngLBoundX, 4
    
    'Debug.Print "HP Transposing the array took: " & mTimer.MillisecondsElapsed & " milliseconds..."
    
    'That's it -> the list is transposed

End Sub

Private Sub GetColumns(ByVal idx As Long, ByRef slicedArray As Variant)
' ***********************************************
' slicedArray can be passed as a 1d or 2d array
' sliceArray can also be part bound, eg  slicedArray(1 to 100) or slicedArray(10 to 100)
' ***********************************************

    Dim ptrToArrayVar As LongPtr
    Dim ptrToSafeArray As LongPtr
    Dim ptrToArrayData As LongPtr
    Dim ptrToArrayData2 As LongPtr
    Dim uSAFEARRAY As SAFEARRAY1D
    Dim ptrCursor As LongPtr
    Dim cbElements As Long
    Dim atsBound1 As Long
    Dim elSize As Long
        
    '§§§ TO BE TESTED... POTENTIALLY UNSAFE CODE §§§
    
    'determine bound1 of source array (ie row Count)
    atsBound1 = m_NumItems
    
    'get pointer to source array Safearray
    'ptrToArrayVar = VarPtrArray(m_List) 'does not work that well under VBA
    'CopyMemory ptrToSafeArray, ByVal ptrToArrayVar, PTR_LENGTH
    ptrToSafeArray = getSafeArrayDescrPtr(m_List)
    CopyMemory uSAFEARRAY, ByVal ptrToSafeArray, LenB(uSAFEARRAY)
    ptrToArrayData = uSAFEARRAY.pvData
    'determine byte size of source elements
    cbElements = uSAFEARRAY.cbElements

    'get pointer to destination array Safearray
    ptrToArrayVar = VarPtr(slicedArray) + 8 'Variant reserves first 8bytes
    CopyMemory ptrToSafeArray, ByVal ptrToArrayVar, PTR_LENGTH
    CopyMemory uSAFEARRAY, ByVal ptrToSafeArray, LenB(uSAFEARRAY)
    ptrToArrayData2 = uSAFEARRAY.pvData

    'determine elements size
    elSize = UBound(slicedArray, 1) - LBound(slicedArray, 1) + 1
    'determine start position of data in source array
    ptrCursor = ptrToArrayData + (((idx - 1) * atsBound1 + LBound(slicedArray, 1) - 1) * cbElements)
    'Copy source array to destination array
    CopyMemory ByVal ptrToArrayData2, ByVal ptrCursor, cbElements * elSize

End Sub

Private Function DimensionCount(ByRef ArrayToCheck As Variant) As Integer
' ***********************************************
' Get the dimensions of the array
' ***********************************************
        
        'Dim lVariableType As Long
        Dim lpSADescriptor As Long
        Dim lpSafeArray As Long
        'Dim SafeArray2 As SAFEARRAY2D
        'Dim aToCheck As Variant: aToCheck = m_List
        
        If Not IsArray(ArrayToCheck) Then
                DimensionCount = -1
                Exit Function
        End If

        ' Get the tpye of the variable
        'CopyMemory lVariableType, ByVal VarPtr(ArrayToCheck), 2
    
        ' Get address of SAFEARRAY-Descriptor
        CopyMemory lpSADescriptor, ByVal VarPtr(ArrayToCheck) + 8, 4

        ' Get address of SAFEARRAY-Structure
        CopyMemory lpSafeArray, ByVal lpSADescriptor, 4
        
        If lpSafeArray <> 0 Then
            ' Get the first 2 Bytes, which contain the dimensitons of the array:
            CopyMemory DimensionCount, ByVal lpSafeArray, 2
        End If
          
End Function


'#Region
'#======================================================================================================================
'# General routines
'#======================================================================================================================

Private Function GetColumnIndex(ByVal Index As String) As Long
' ***********************************************
' Get the column number of a given index
' Returns either the number provided (as Long rather than String)
' or the index based on the name of the column provided
' ***********************************************

    Dim iCol As Long: iCol = -1

    'initialize with Default not found = -1
    GetColumnIndex = iCol
    
    If IsNumeric(Index) Then
        'Use the Column index provided
        iCol = CLng(Index)
        GetColumnIndex = iCol
     Else
        'Find the column index based on the field name
        For iCol = 1 To m_NumCols
            If m_Columns(iCol) = Index Then
                    GetColumnIndex = iCol
                    Exit For
            End If
        Next iCol
    End If
    
End Function

Private Function GetColumnName(ByVal Index As Long) As String
' ***********************************************
' Get the column name of a given index
' ***********************************************

    GetColumnName = m_Columns(Index)

End Function

Private Function IsTableColumn(ByVal Index As String) As Boolean
' ***********************************************
' Checks if the column Name
' ***********************************************
    
    If GetColumnIndex(Index) = -1 Then
        IsTableColumn = False
     Else
        IsTableColumn = True
    End If


End Function


'#Region
'#======================================================================================================================
'# Service Routines
'#======================================================================================================================

Private Function Compare(value As Variant, CompareWith As Variant, Comparison As MatchType) As Boolean
' ***********************************************
' Compares Value1 with Value2 and returns TRUE
' if specific comparison is positive
' ***********************************************
        
    Dim stValue As String
    Dim stCompareWith As String
    Dim lLenCW As Long

    Compare = False

    Select Case Comparison
    
        Case Equals
            If value = CompareWith Then Compare = True
        Case GreaterThen
            If value > CompareWith Then Compare = True
        Case GreaterThenOrEqual
            If value >= CompareWith Then Compare = True
        Case LessThen
            If value < CompareWith Then Compare = True
        Case LessThenOrEqual
            If value <= CompareWith Then Compare = True
        Case MatchesPattern, DoesNotMatchPattern
            stValue = value
            stCompareWith = CompareWith
            
            If Left$(stCompareWith, 1) = "*" Then
                If Right$(stCompareWith, 1) = "*" Then
                    'Contains...
                    If Comparison = MatchesPattern Then
                        If InStr(stValue, stCompareWith) > 0 Then Compare = True
                     Else
                        If InStr(stValue, stCompareWith) <> 0 Then Compare = True
                    End If
                 Else
                    'Ends With
                    lLenCW = Len(stCompareWith) - 1
                    stValue = Right$(stValue, lLenCW)
                    stCompareWith = Left$(stCompareWith, lLenCW)
                    If Comparison = MatchesPattern Then
                        If stValue = stCompareWith Then Compare = True
                     Else
                        If stValue <> stCompareWith Then Compare = True
                    End If
                End If
             Else
                If Right$(stCompareWith, 1) = "*" Then
                    'Begins With
                    lLenCW = Len(stCompareWith) - 1
                    stValue = Left$(stValue, lLenCW)
                    stCompareWith = Left$(stCompareWith, lLenCW)
                    If Comparison = MatchesPattern Then
                        If stValue = stCompareWith Then Compare = True
                     Else
                        If stValue <> stCompareWith Then Compare = True
                    End If
                End If
            End If
        Case Else
            'nothing to do
    End Select

End Function

Private Function getSafeArrayDescrPtr(ByRef varInput As Variant) As Long

    Dim lngSafeArrayDescr As Long
    
    '1st pointer only - get the pointer to the Safearray descriptor of input array
    CopyMemory lngSafeArrayDescr, ByVal (VarPtr(varInput) + 8), 4
    
    getSafeArrayDescrPtr = lngSafeArrayDescr

End Function

Private Function getSafeArrayStructPtr(ByRef varInput As Variant) As Long

    Dim lngSafeArrayDescr As Long
    Dim lngSafeArrayStruct As Long
    
    '1st pointer - get the pointer to the Safearray desciptor of input array
    CopyMemory lngSafeArrayDescr, ByVal (VarPtr(varInput) + 8), 4
    
    '2nd pointer - get the pointer to the Safearray desciptor of input array
    Call CopyMemory(lngSafeArrayStruct, ByVal lngSafeArrayDescr, 4&)
    
    getSafeArrayStructPtr = lngSafeArrayStruct

End Function




'Gets and copies a SafeArray for debugging output
Private Sub DebugSafeArray(ArrayInput As Variant)
    'Dim MyArray() As Long
    Dim SafeArray As SAFEARRAY1D
    Dim SafeArray2 As SAFEARRAY2D
    Dim lDimensions As Long
    Dim PointerToPointer As Long
    Dim PointerToSafeArray As Long

    lDimensions = DimensionCount(ArrayInput)
    
    'Ersten Zeiger ermitteln
    PointerToPointer = getSafeArrayDescrPtr(ArrayInput)
    
    'Zweiten Zeiger auslesen
    Call CopyMemory(PointerToSafeArray, ByVal PointerToPointer, 4&)
    
    
    If lDimensions = 0 Then
        Debug.Print "Array not initialized!"
     ElseIf lDimensions = 1 Then
        'SafeArray in vorgegebene Strukur kopieren
        Call CopyMemory(SafeArray, ByVal PointerToSafeArray, _
        Len(SafeArray))
        'Resultat ausgeben
        Call DisplaySafeArray1D(SafeArray)
     ElseIf lDimensions = 2 Then
        'SafeArray in vorgegebene Strukur kopieren
        Call CopyMemory(SafeArray2, ByVal PointerToSafeArray, _
        Len(SafeArray2))
        'Resultat ausgeben
        Call DisplaySafeArray2D(SafeArray2)
    End If
    
End Sub

'Prints a given SafeArray structure - 1 dimensional
Private Sub DisplaySafeArray1D(SA As SAFEARRAY1D)
  
  Dim H As String
    
        H = ".cDims = " & SA.cDims & vbCrLf
        H = H & ".fFeatures = " & SA.fFeatures & vbCrLf
        H = H & ".cbElements = " & SA.cbElements & vbCrLf
        H = H & ".cLocks = " & SA.cLocks & vbCrLf
        H = H & ".pvData = " & SA.pvData & vbCrLf
    
        H = H & ".Bounds(0).cElements = " & SA.Bounds(0).cElements & vbCrLf

        H = H & ".Bounds(0).lLbound = " & SA.Bounds(0).lLbound & vbCrLf
        
    Debug.Print H
    
End Sub


'Prints a given SafeArray structure - 1 dimensional
Private Sub DisplaySafeArray2D(SA As SAFEARRAY2D)
  
  Dim H As String
    
        H = ".cDims = " & SA.cDims & vbCrLf
        H = H & ".fFeatures = " & SA.fFeatures & vbCrLf
        H = H & ".cbElements = " & SA.cbElements & vbCrLf
        H = H & ".cLocks = " & SA.cLocks & vbCrLf
        H = H & ".pvData = " & SA.pvData & vbCrLf
    
        H = H & ".Bounds(0).cElements = " & SA.Bounds(0).cElements & vbCrLf
        H = H & ".Bounds(1).cElements = " & SA.Bounds(1).cElements & vbCrLf
        
        H = H & ".Bounds(0).lLbound = " & SA.Bounds(0).lLbound & vbCrLf
        H = H & ".Bounds(1).lLbound = " & SA.Bounds(1).lLbound & vbCrLf
       
    Debug.Print H
    
End Sub




Private Function GetVarType(myVar) As String

' ---------------------------------------------------------------
' Written By Shanmuga Sundara Raman for http://vbadud.blogspot.com
' ---------------------------------------------------------------

'§§§ kept here for the sake of having all data types in one place, maybe used at a later stage §§§

    If VarType(myVar) = vbNull Then
        GetVarType = "Null (no valid data) "
    ElseIf VarType(myVar) = vbInteger Then
        GetVarType = "Integer "
    ElseIf VarType(myVar) = vbLong Then
        GetVarType = "Long integer "
    ElseIf VarType(myVar) = vbSingle Then
        GetVarType = "Single-precision floating-point number "
    ElseIf VarType(myVar) = vbDouble Then
        GetVarType = "Double-precision floating-point number "
    ElseIf VarType(myVar) = vbCurrency Then
        GetVarType = "Currency value "
    ElseIf VarType(myVar) = vbDate Then
        GetVarType = "Date value "
    ElseIf VarType(myVar) = vbString Then
        GetVarType = "String "
    ElseIf VarType(myVar) = vbObject Then
        GetVarType = "Object "
    ElseIf VarType(myVar) = vbError Then
        GetVarType = "Error value "
    ElseIf VarType(myVar) = vbBoolean Then
        GetVarType = "Boolean value "
    ElseIf VarType(myVar) = vbVariant Then
        GetVarType = "Variant (used only with arrays of variants) "
    ElseIf VarType(myVar) = vbDataObject Then
         GetVarType = "A data access object "
    ElseIf VarType(myVar) = vbDecimal Then
        GetVarType = "Decimal value "
    ElseIf VarType(myVar) = vbByte Then
        GetVarType = "Byte value "
    ElseIf VarType(myVar) = vbUserDefinedType Then
        GetVarType = "Variants that contain user-defined types "
    ElseIf VarType(myVar) = vbArray Then
        GetVarType = "Array "
    Else
        GetVarType = VarType(myVar)
    End If

End Function


