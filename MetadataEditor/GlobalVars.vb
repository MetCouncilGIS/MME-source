Option Explicit On

'Imports System.Reflection
'Imports System.IO

Module GlobalVars

    ''' <summary>
    ''' Configuration from XML file.
    ''' </summary>
    ''' <remarks>We use this to give the user a way to configure some options that are not configurable through the user interface.</remarks>
    Dim config As MMEConfig

    ''' <summary>
    ''' Enum that defines application states we are interested in keeping track of
    ''' </summary>
    ''' <remarks></remarks>
    Private Enum appState As Byte
        uninitialized
        enabled
        disabled
    End Enum

    ''' <summary>
    ''' Variable that keeps track of the current application state
    ''' </summary>
    ''' <remarks></remarks>
    Private currentAppState As appState = appState.uninitialized

    ''' <summary>
    ''' Get/Set if the application is enabled
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property enabled As Boolean
        Get
            If currentAppState = appState.uninitialized Then
                If checkAndCopyTemplate() Then
                    config = DeserializeFromXmlFile(Utils.getAppDataFolder + "\config.xml", GetType(MMEConfig))
                    currentAppState = appState.enabled
                Else
                    currentAppState = appState.disabled
                    MessageBox.Show("Problem while initializing the application. Exiting...", GlobalVars.nameStr, MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
            Return currentAppState = appState.enabled
        End Get
        Set(ByVal enable As Boolean)
            If enable Then
                currentAppState = appState.enabled
            Else
                currentAppState = appState.disabled
            End If
        End Set
    End Property


    ''' <summary>
    ''' Path to supporting MSAccess database
    ''' </summary>
    ''' <remarks>This can be overridden by the user in config.xml</remarks>
    Public ReadOnly Property mdbPath()
        Get
            ' Try to get path from config.xml
            My.Settings.MdbFilepathname = config.MdbFilepathname.Trim

            If _
                My.Settings.MdbFilepathname Is Nothing OrElse _
                Not My.Settings.MdbFilepathname.EndsWith(".mdb") OrElse _
                Not System.IO.File.Exists(My.Settings.MdbFilepathname) _
            Then
                ' If the database reference does not look valid, use default database.
                If System.IO.File.Exists(mdbPathDefault()) Then
                    My.Settings.MdbFilepathname = mdbPathDefault()
                Else
                    Return Nothing
                End If
            End If

            ' If we got here, then use configured database.
            Return My.Settings.MdbFilepathname
        End Get
    End Property

    ''' <summary>
    ''' Default path to supporting MSAccess database
    ''' </summary>
    ''' <remarks>Actual path can be retrieved via mdbPath property.</remarks>
    Public ReadOnly Property mdbPathDefault() As String
        Get
            Return Utils.getAppDataFolder() & "\metadata.mdb"
        End Get
    End Property

    ''' <summary>
    ''' The OLEDB connection string used to connect to the MSAccess database.
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property connStr() As String
        Get
            Return "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & mdbPath
        End Get
    End Property

    ''' <summary>
    ''' The separator used internally by EME to seperate various postfixes from identifiers based on the XSL pattern for the element
    ''' </summary>
    ''' <remarks></remarks>
    Public idSep As String = "_____"

    ''' <summary>
    ''' Name string for this application used in a number of contexts. 
    ''' Some other name references had to be hardwired, so changing the value here alone is not sufficient.
    ''' </summary>
    ''' <remarks></remarks>
    Public nameStr As String = My.Application.Info.ProductName

    ''' <summary>
    ''' Variable that holds the metadata record copy being edited.
    ''' </summary>
    ''' <remarks></remarks>
    Public iXPS As XmlMetadata

    ''' <summary>
    ''' Indicator that the user has performed a save without closing EME.
    ''' </summary>
    ''' <remarks></remarks>
    Public savedSession As Boolean

    ''' <summary>
    ''' Indicator that a previously saved metadata record was recovered during current EME session.
    ''' </summary>
    ''' <remarks></remarks>
    Public recoveredSession As Boolean

    ''' <summary>
    ''' (Re)Initialize global variables as applicable.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub init()
        savedSession = False
        recoveredSession = False
    End Sub

    ''' <summary>
    ''' The user thesaurus being used if applicable.
    ''' </summary>
    ''' <remarks></remarks>
    Private _userThesaurus As String = Nothing

    ''' <summary>
    ''' The property providing controlled access to user thesaurus.
    ''' </summary>
    ''' <remarks></remarks>
    Public ReadOnly Property userThesaurus()
        Get
            Try
                If _userThesaurus Is Nothing Then
                    '_userThesaurus = Utils.getFromSingleValuedSQL("SELECT * FROM GetUserThesaurus").ToString.Split("'")(1)
                    '_userThesaurus = Utils.getFromSingleValuedSQL("SELECT themekt FROM 1l_KeywordsUser")
                    _userThesaurus = "User"
                End If
            Catch
                _userThesaurus = "User"
                MsgBox("Error while identifying user-specified thesaurus!")
            End Try
            Return _userThesaurus
        End Get
    End Property


    ''' <summary>
    ''' Keeps track of the HTML Help process (hh.exe) to avoid having more than one help window open.
    ''' </summary>
    ''' <remarks></remarks>
    Public proc As New Process()


    ''' <summary>
    ''' Enumeration for metadata validation modes:
    ''' <c>Webservice</c> uses EPA's validation webservice with fallback to local validation upon failure or timeout.
    ''' <c>Local</c> uses local validation service hardwired into the editor.
    ''' </summary>
    ''' <remarks>
    ''' At the time of release, webservice and local validation both yield the same results.
    ''' It is conceivable that the webservice validation may be updated ahead of or without an update to EME.
    ''' </remarks>
    Public Enum ValidationMode
        Webservice
        Local
    End Enum

    ''' <summary>
    ''' Holds the validation mode in effect.
    ''' </summary>
    ''' <remarks></remarks>
    Private _ValidationMode As ValidationMode

    ''' <summary>
    ''' Gets/sets the validation mode in effect.
    ''' </summary>
    ''' <remarks></remarks>
    Public Property CurrentValidationMode() As ValidationMode
        Get
            Return _ValidationMode
        End Get
        Set(ByVal Value As ValidationMode)
            _ValidationMode = Value
        End Set
    End Property

End Module
