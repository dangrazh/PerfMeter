VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPerformanceMeter"
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
'     |> Get | --- Name                 :: Returns name of the class.
'     |> Get | --- Version              :: Returns version string for the class [e.g. #.# (year)].


'#======================================================================================================================
'# References
'#======================================================================================================================
'Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal bytes As Long)

'#======================================================================================================================
'# Dependencies to other classes
'#======================================================================================================================
' cDataTable      :: used to store the results of each execution
' cStopWatchHD    :: used to measure the elapsed time for execution in micro seconds


'#======================================================================================================================
'# Application Constants, Enumerations & Types
'#======================================================================================================================
Private Const C_Name As String = "cPerformanceMeter.cls"

Public Enum OutputChannel
    StringOut = 1
    ImmedateWindowOut = 2
    FileOut = 3
End Enum


'#======================================================================================================================
'# Private Variables
'#======================================================================================================================
Private m_Storage               As cDataTable           'The table holding the results of all runs
Private m_Timer                 As cStopWatchHD         'The stop watch itself
Private m_Runtime               As Double               'The runtime of a dedecated run
Private m_RunID                 As Long                 'The RunID of a dedicated run
Private m_SingleRun(1 To 3)     As Variant              'Array holding the result of a single run



'#Region
'#======================================================================================================================
'# Class Initialization, Termination & Properties
'#======================================================================================================================


Private Sub Class_Initialize()
' ************************************************
' Class constructor.
' ************************************************

    'Debug.Print "|> Initializing:= " & Me.Name
    
    Set m_Storage = New cDataTable
    Set m_Timer = New cStopWatchHD
    
    m_Storage.DefineTable 3, "RunID, RunDescription, RunTime (s)"

End Sub


Private Sub Class_Terminate()
' ************************************************
' Class destructor.
' ************************************************
    
    'Debug.Print "|> Terminating:= " & Me.Name
    
    Set m_Storage = Nothing
    Set m_Timer = Nothing

End Sub

Public Property Get Version() As String
' ************************************************
' Version string of the current class.
' Contains a list of (historical) changes to the class within the comments of the procedure.
' ************************************************
   'Version = "Version 1.0 (2017)"  'Initial release
   Version = "Version 1.1 (2018-08)"  'Redesigned PrintReport to prevent performance issue with string concatination in combination with high numbers of records.

End Property

Public Property Get About() As String
' ***********************************************
' String that describes the current class.
' ***********************************************
    About = "<Description of the Class> Version: " & Me.Version & "." & VBA.vbCrLf & VBA.vbCrLf
    About = About & "For additional details please contact the author."
End Property

Public Property Get Name() As String
' ***********************************************
' Returns the name of the class.
' ***********************************************
    Name = C_Name
    
End Property


'#Region
'#======================================================================================================================
'# cStopWatchHD Manipulation
'#======================================================================================================================

Public Sub Reset()

    m_Timer.Reset

End Sub

Public Sub Restart()

    m_Timer.Restart

End Sub

Public Sub Start()

    m_Timer.Start
    
End Sub

Public Sub Pause()

    m_Timer.Pause

End Sub

Public Sub LogRun(RunDescription As String, Optional RunID As Long = -1)
    
    m_Runtime = m_Timer.Elapsed
    
    m_SingleRun(1) = RunID
    m_SingleRun(2) = RunDescription
    m_SingleRun(3) = m_Runtime
    
    m_Storage.RecordAdd m_SingleRun

End Sub



'#Region
'#======================================================================================================================
'# Perf Meter Methods
'#======================================================================================================================

Public Sub PrintReport(Optional OutputTo As OutputChannel = ImmedateWindowOut, Optional IncludeDetails As Boolean = False, Optional ByRef stOut As String, Optional stFileOut As String = "") 'As String

    Dim stRunID As String
    Dim stRunDescr As String
    Dim stRunTime As String
    Dim stHeaderOut As String
    Dim stOutput As String
    Dim stSummaryHeaderOut As String
    Dim stSummaryOut As String
    Dim aRecord As Variant
    Dim stHeader As Variant
    Dim i As Long
    Dim htRunDescr As cHashTable
    Dim tSummary As cDataTable
    Dim rSummary(1 To 5) As Variant
    Dim lptrToData As Long
    Dim iFH As Integer
    
    'Dim stRunDescr As String
    Dim dblMin As Double: dblMin = 0
    Dim dblMax As Double: dblMax = 0
    Dim dblRunTime As Double: dblRunTime = 0
    
    Set tSummary = New cDataTable
    Set htRunDescr = New cHashTable
    
    tSummary.DefineTable 5, "NoOfRuns, RunDescription, MinRuntime, MaxRuntime, AvgRuntime"
    
    'create header
    For Each stHeader In m_Storage.Headers
        stHeaderOut = stHeaderOut & stHeader & vbTab
    Next
    
    stHeaderOut = stHeaderOut & vbCrLf


    'Check for detailed report with output channel other than file
    If IncludeDetails = True And OutputTo <> FileOut Then
        MsgBox "Detailed output only supported when writing to file. Summary report will be produced!", vbInformation
    End If

    'Create file handle if output to file
    If OutputTo = FileOut Then
        iFH = FreeFile
        Open stFileOut For Output As #iFH
    End If

    'process the logged data
    m_Storage.RsMoveFirst
    Do While Not m_Storage.RsEOF
        
        stRunID = m_Storage.Item(1)
        stRunDescr = m_Storage.Item(2)
        stRunTime = Format(m_Storage.Item(3), "0.000000")
        dblRunTime = m_Storage.Item(3)
        
        'create the detail record if required
        If IncludeDetails = True And OutputTo = FileOut Then
            'here goes the file output
            stOutput = stOutput & stRunID & vbTab
            stOutput = stOutput & stRunDescr & vbTab
            stOutput = stOutput & stRunTime
            
            Print #iFH, stOutput
            
            stOutput = ""
        End If
        
        'create the summary data
        If htRunDescr.Exists(stRunDescr) Then
        
            'get the pointer to the summary table
            lptrToData = htRunDescr.Item(stRunDescr)
            
            'update the summary table
            tSummary.Item("NoOfRuns", lptrToData) = tSummary.Item("NoOfRuns", lptrToData) + 1
            'maintain min and max values
            If dblRunTime > tSummary.Item("MaxRuntime", lptrToData) Then tSummary.Item("MaxRuntime", lptrToData) = dblRunTime
            If dblRunTime < tSummary.Item("MinRuntime", lptrToData) Then tSummary.Item("MinRuntime", lptrToData) = dblRunTime
            'calculate the average value - calculate the travelling mean: ('previous mean' * ('count' - 1) + 'new value') / 'count'
            tSummary.Item("AvgRuntime", lptrToData) = (tSummary.Item("AvgRuntime", lptrToData) * (tSummary.Item("NoOfRuns", lptrToData) - 1) + dblRunTime) / tSummary.Item("NoOfRuns", lptrToData)
         Else
            
            'add record to the summary table
            rSummary(1) = 1
            rSummary(2) = stRunDescr
            rSummary(3) = dblRunTime
            rSummary(4) = dblRunTime
            rSummary(5) = dblRunTime
            tSummary.RecordAdd rSummary
            
            'add key to the hash table with pointer to summary table
            htRunDescr.Add stRunDescr, tSummary.NumItems
            
        End If
            
      m_Storage.RsMoveNext
    Loop
    
    
    'create the summary
    
    'create summary header
    For Each stHeader In tSummary.Headers
        stSummaryHeaderOut = stSummaryHeaderOut & stHeader & vbTab
    Next
    
    stSummaryHeaderOut = stSummaryHeaderOut & vbCrLf
    
    'process the summary table
    tSummary.RsMoveFirst
    Do While Not tSummary.RsEOF
        stSummaryOut = stSummaryOut & tSummary.Item(1) & vbTab
        stSummaryOut = stSummaryOut & tSummary.Item(2) & vbTab
        stSummaryOut = stSummaryOut & Format(tSummary.Item(3), "0.000000") & vbTab
        stSummaryOut = stSummaryOut & Format(tSummary.Item(4), "0.000000") & vbTab
        stSummaryOut = stSummaryOut & Format(tSummary.Item(5), "0.000000") & vbCrLf
     tSummary.RsMoveNext
    Loop
    
    
    'produce the final output
    stOutput = "*** Report Start ***" & vbCrLf & vbCrLf & "*** Summary ***" & vbCrLf & stSummaryHeaderOut & stSummaryOut & vbCrLf & vbCrLf & "*** Report End ***"
    If OutputTo = FileOut Then
        If IncludeDetails = True Then
            'add the summary to the end of the file
            stOutput = vbCrLf & vbCrLf & "*** Summary ***" & vbCrLf & stSummaryHeaderOut & stSummaryOut & vbCrLf & vbCrLf & "*** Report End ***"
            Print #iFH, stOutput
         Else
            'simply dump the summary
            Print #iFH, stOutput
        End If
        Close #iFH
     Else
        If OutputTo = StringOut Then
            stOut = stOutput
         ElseIf OutputTo = ImmedateWindowOut Then
            Debug.Print stOutput
         Else
            MsgBox "Not supported output channel!", vbInformation
        End If
    End If
        
End Sub

Public Function Clear()

    m_Storage.TruncateTable False
    m_Timer.Reset

End Function
