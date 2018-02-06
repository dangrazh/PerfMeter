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


'#======================================================================================================================
'# Private Variables
'#======================================================================================================================
Private m_Storage               As cDataTable           'The table holding the results of all runs
Private m_Timer                 As cStopWatchHD         'The stop watch itself
Private m_Runtime               As Double               'The runtime of a dedecated run
Private m_RunID                 As Long                 'The RunID of a dedicated run
Private m_SingleRun(1 To 3)     As Variant          'Array holding the result of a single run



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
   Version = "Version 1.0 (2017)"  'Initial release

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


Public Function PrintReport(Optional IncludeDetails As Boolean = False) As String

    Dim stRunID As String
    Dim stRunDescr As String
    Dim stRunTime As String
    Dim stHeaderOut As String
    Dim stOut As String
    Dim stSummaryHeaderOut As String
    Dim stSummaryOut As String
    Dim aRecord As Variant
    Dim stHeader As Variant
    Dim i As Long
    Dim htRunDescr As Dictionary
    Dim tSummary As cDataTable
    Dim tSumTmp As cDataTable
    Dim rSummary(1 To 7) As Variant
    Dim lptrToData As Long
    
    'Dim stRunDescr As String
    Dim dblMin As Double: dblMin = 0
    Dim dblMax As Double: dblMax = 0
    Dim dblRunTime As Double: dblRunTime = 0
    Dim dblRunTimeSum As Double
    Dim dblRunTimeAvg As Double
    Dim dblMedian As Double
    Dim dbl80thPctl As Double
    
    Set tSummary = New cDataTable
    Set tSumTmp = New cDataTable
    Set htRunDescr = New Dictionary
    
    tSummary.DefineTable 7, "NoOfRuns, RunDescription, MinRuntime, MaxRuntime, AvgRuntime, MedianRuntime, 80thPctlRuntime"
    tSumTmp.DefineTable 3, "RunID, RunDescription, RunTime (s)"
    
    'create header
    For Each stHeader In m_Storage.Headers
        stHeaderOut = stHeaderOut & stHeader & vbTab
    Next
    
    stHeaderOut = stHeaderOut & vbCrLf

    'process the logged data
    m_Storage.RsMoveFirst
    Do While Not m_Storage.RsEOF
        
        stRunID = m_Storage.Item(1)
        stRunDescr = m_Storage.Item(2)
        stRunTime = Format(m_Storage.Item(3), "0.000000")
        dblRunTime = m_Storage.Item(3)
        
        'create the detail record
        stOut = stOut & stRunID & vbTab
        stOut = stOut & stRunDescr & vbTab
        stOut = stOut & stRunTime & vbCrLf
        
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
            rSummary(6) = 0
            rSummary(7) = 0
            tSummary.RecordAdd rSummary
            
            'add key to the hash table with pointer to summary table
            htRunDescr.Add stRunDescr, tSummary.NumItems
            
        End If
            
      m_Storage.RsMoveNext
    Loop
    
    
    'calculate the Average, Median and 80th Percentile
    tSummary.RsMoveFirst
    Do While Not tSummary.RsEOF
        stRunDescr = tSummary.Item("RunDescription")
        
        'find all records for a given Run Description
        
        'move to the beginning of the recordset
        m_Storage.RsMoveFirst
        
        'find the 1st entry
        m_Storage.RsFind "RunDescription", Equals, stRunDescr
        
        'check for RsEOF = False -> something found
        If m_Storage.RsEOF = False Then
            'process the entries
            Do While Not m_Storage.RsEOF
                tSumTmp.RecordAdd m_Storage.Record
                dblRunTimeSum = dblRunTimeSum + m_Storage.Item("RunTime (s)")
              m_Storage.RsFind "RunDescription", Equals, stRunDescr
            Loop
            
            'calculate the meadian
            dblMedian = Percentile(tSumTmp, "RunTime (s)", 0.5, True)
            'calculate the 80th percentile
            dbl80thPctl = Percentile(tSumTmp, "RunTime (s)", 0.8, False)
            
            'Update the summary table
            tSummary.Item("MedianRuntime") = dblMedian
            tSummary.Item("80thPctlRuntime") = dbl80thPctl
            
            'clean the temp storage
            tSumTmp.TruncateTable False
        End If
     
     tSummary.RsMoveNext
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
        stSummaryOut = stSummaryOut & Format(tSummary.Item(5), "0.000000") & vbTab
        stSummaryOut = stSummaryOut & Format(tSummary.Item(6), "0.000000") & vbTab
        stSummaryOut = stSummaryOut & Format(tSummary.Item(7), "0.000000") & vbCrLf
        
     tSummary.RsMoveNext
    Loop
    
    'produce the final output
    If IncludeDetails = True Then
        PrintReport = "*** Report Start ***" & vbCrLf & vbCrLf & "*** Summary ***" & vbCrLf & stSummaryHeaderOut & stSummaryOut & vbCrLf & vbCrLf & "*** Details ***" & vbCrLf & stHeaderOut & stOut & vbCrLf & vbCrLf & "*** Report End ***"
     Else
        PrintReport = "*** Report Start ***" & vbCrLf & vbCrLf & "*** Summary ***" & vbCrLf & stSummaryHeaderOut & stSummaryOut & vbCrLf & vbCrLf & "*** Report End ***"
    End If
    
End Function


Private Function Percentile(ByRef tInput As cDataTable, Index As Variant, Pctl As Double, Optional bNeedsSorting As Boolean = False) As Double

    Dim NumElems As Double
    Dim NthEl As Double
    Dim StrtPctle As Double
    Dim IncrPctle As Double
    Dim DeltPctle As Double
    
    If bNeedsSorting = True Then
        tInput.Sort Index, "asc"
    End If
    
    NumElems = tInput.NumItems
    NthEl = (NumElems - 1) * (Pctl) + 1
    If (NthEl = Int(NthEl)) Then
        Percentile = tInput.Item(Index, Int(NthEl))
     Else
        StrtPctle = tInput.Item(Index, Int(NthEl - 1))
        IncrPctle = (NthEl - Int(NthEl))
        DeltPctle = (tInput.Item(Index, Int(NthEl)) - tInput.Item(Index, Int(NthEl - 1)))
        Percentile = StrtPctle + (IncrPctle * DeltPctle)
    End If


End Function
