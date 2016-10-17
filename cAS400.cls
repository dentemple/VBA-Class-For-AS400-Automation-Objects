''''''''''''''''''''''''''''''''''''''''''''''''''
' Organization:
'   I.      Initialization & custom procedures
'   II.     Direct automation object wrappers
'   III.    autECLConnList Class
'   IV.     autECLPS Class
'   V.      autECLOIA Class

' Additional documentation for the automation objects can be found here: _
    http://

'       Created By:     Den Temple | dentemple.io
'       Creation Date:  May 8, 2015
'       Last Modified:
''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Error Constants
    Const CUSTOM_ERROR_CODE As Long = vbObjectError + 65535     'Maximum possible value
    Const ERROR_SOURCE As String = "cAS400.Property"
    Const ASSERT_ERROR_DESCRIPTION As String = "Either the AS400 is not active, or " & _
        "the AS400 attempted an action prior to the completion of a previous operation."
        
' Internal Constants

' iSeries automation objects (properties)
    Private pConnectionList As Object
    Private pConnectionManager As Object
    Private pOperatorInformation As Object
    Private pFieldList As Object
    Private pScreenDescription As Object
    Private pScreenRecognition As Object
    Private pSessionInformation As Object
    Private pWindowMetrics As Object
    Private pFileTransfer As Object
    Private pPageSettings As Object
    Private pPrinterSettings As Object
    Private pPresentationSpace As Object


''''''''''''''''''''''''''''''''''''''''''''''''''
' I. Initialization & custom procedures
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitializeAS400(ByVal sessionName As String)

' Sets the initial connection to the iSeries/AS400
' Input: sessionName (string). The iseries provides a single alphabetical _
    character to each application window, which can be called by name. _
    By default, the very first window is given the letter "A".
    
    ' Must be set independently
    Set pConnectionManager = CreateObject("PCOMM.autECLConnMgr")
    Set pScreenDescription = CreateObject("PCOMM.autECLScreenDesc")
    Set pScreenRecognition = CreateObject("PCOMM.autECLScreenReco")
    Set pSessionInformation = CreateObject("PCOMM.autECLSession")
    
    ' Note: autECLSession session must be set prior to sub-classing _
        any additional classes under it.
    pSessionInformation.SetConnectionByName (sessionName)
    
    ' Can be connected to a higher object for convenience
    Set pConnectionList = pConnectionManager.autECLConnList
    Set pOperatorInformation = pSessionInformation.autECLOIA
    Set pPresentationSpace = pSessionInformation.autECLPS
    Set pWindowMetrics = pSessionInformation.autECLWinMetrics
    Set pFileTransfer = pSessionInformation.autECLXfer
    Set pPageSettings = pSessionInformation.autECLPageSettings
    Set pPrinterSettings = pSessionInformation.autECLPrinterSettings
    
    ' Recommended by the documentation, but not required
    pConnectionList.Refresh
    
End Sub

Private Sub AssertReady()

    If Me.IsReady = False Then
        Err.Raise CUSTOM_ERROR_CODE, ERROR_SOURCE, ASSERT_ERROR_DESCRIPTION
    End If
    
End Sub

Public Sub RunComponentTests()
    
    Debug.Print "== Run Test =="
    Debug.Print ""
    
    With Me

        Debug.Print "== App Status =="
        Debug.Print "Is Started: " & .IsStarted
        Debug.Print "Is Connected: " & .IsConnected
        Debug.Print "Is Enabled: " & .IsEnabled
        Debug.Print "Is Ready: " & .IsReady
        Debug.Print ""
        Debug.Print "== autECLPS Class =="
        Debug.Print "Rows & Cols Count: (" & .RowsCount & ", " & .ColumnsCount & ")"
        Debug.Print "Cursor Position: (" & .CursorPositionRow & ", " & .CursorPositionColumn & ")"
        Debug.Print "Current Session: " & .CurrentConnectionName & ", " & _
            .CurrentConnectionHandle & ", " & .CurrentConnectionType
            
    End With
    
    Debug.Print ""
    Debug.Print "== End Test =="

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''
' II. Direct automation object wrappers
''''''''''''''''''''''''''''''''''''''''''''''''''
' The following properties wrap only the automation objects. These _
    wrappers will be used internally by the class.
' These wrappers are public and can also be called outside the class. _
    This can be useful for experimentation and debugging purposes.
' EXAMPLE:
'   >>> Dim App As cAS400: Set App = New cAS400: App.InitializeAS400 ("A")
'   >>> App.PresentationSpaceObject.GetText 1,1,50
''''''''''''''''''''''''''''''''''''''''''''''''''

Public Property Get ConnectionManagerObject() As Object
    Set ConnectionManagerObject = pConnectionManager
End Property

Public Property Get ScreenDescriptionObject() As Object
    Set ScreenDescriptionObject = pScreenDescription
End Property

Public Property Get ScreenRecognitionObject() As Object
    Set ScreenRecognitionObject = pScreenRecognition
End Property

Public Property Get SessionInformationObject() As Object
    Set SessionInformationObject = pSessionInformation
End Property

Public Property Get ConnectionListObject() As Object
    Set ConnectionListObject = pConnectionList
End Property

Public Property Get OperatorInformationObject() As Object
    Set OperatorInformationObject = pOperatorInformation
End Property

Public Property Get PresentationSpaceObject() As Object
    Set PresentationSpaceObject = pPresentationSpace
End Property

Public Property Get WindowMetricsObject() As Object
    Set WindowMetricsObject = pWindowMetrics
End Property

Public Property Get FileTransferObject() As Object
    Set FileTransferObject = pFileTransfer
End Property

Public Property Get PageSettingsObject() As Object
    Set PageSettingsObject = pPageSettings
End Property

Public Property Get PrinterSettingsObject() As Object
    Set PrinterSettingsObject = pPrinterSettings
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''
' III. autECLConnList Class
' autECLConnList contains information about all started connections.
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RefreshConnections()
    Me.ConnectionListObject.Refresh
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
' IV. autECLPS Class
' autECLPS performs operations on a presentation space.
''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get RowsCount() As Long
    AssertReady
    RefreshConnections
    RowsCount = Me.PresentationSpaceObject.NumRows
End Property

Public Property Get ColumnsCount() As Long
    AssertReady
    RefreshConnections
    ColumnsCount = Me.PresentationSpaceObject.NumCols
End Property

Public Property Get CursorPositionRow() As Long
    AssertReady
    RefreshConnections
    CursorPositionRow = Me.PresentationSpaceObject.CursorPosRow
End Property

Public Property Get CursorPositionColumn() As Long
    AssertReady
    RefreshConnections
    CursorPositionColumn = Me.PresentationSpaceObject.CursorPosCol
End Property

Public Property Get CurrentConnectionName() As String
    AssertReady
    CurrentConnectionName = Me.PresentationSpaceObject.Name
End Property

Public Property Get CurrentConnectionHandle() As Long
    AssertReady
    CurrentConnectionHandle = Me.PresentationSpaceObject.Handle
End Property

Public Property Get CurrentConnectionType() As String
    AssertReady
    CurrentConnectionType = Me.PresentationSpaceObject.ConnType
End Property

Public Property Get CurrentCodePage() As Long
    AssertReady
    CurrentCodePage = Me.PresentationSpaceObject.CodePage
End Property

Public Property Get IsStarted() As Boolean
    IsStarted = PresentationSpaceObject.Started
End Property

Public Property Get IsConnected() As Boolean
    IsConnected = PresentationSpaceObject.CommStarted
End Property

Public Property Get IsEnabled() As Boolean
    IsEnabled = PresentationSpaceObject.APIEnabled
End Property

Public Property Get IsReady() As Boolean
    ' Checks .Started, .CommStarted, and .APIEnabled
    IsReady = PresentationSpaceObject.Ready
End Property


''''''''''''''''''''''''''''''''''''''''''''''''''
' V. autECLOIA Class
' The autECLOIA object retrieves status from the Host Operator Information Area.

''''''''''''''''''''''''''''''''''''''''''''''''''
