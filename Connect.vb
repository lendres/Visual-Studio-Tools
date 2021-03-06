Imports System
Imports Microsoft.VisualStudio.CommandBars
Imports Extensibility
Imports EnvDTE
Imports EnvDTE80

Public Class Connect
	
	Implements IDTExtensibility2
	Implements IDTCommandTarget

	Dim _application As DTE2
	Dim _addininstance As AddIn

	Dim _textcommands As DPTextCommands
	Dim _properties As DPEnvDTEProperties
	Dim _toolbarutil As DPToolBarsUtils
	Dim _menuitems As DPMenus

	'''<summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
	Public Sub New()

	End Sub

	'''<summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
	'''<param name='application'>Root object of the host application.</param>
	'''<param name='connectMode'>Describes how the Add-in is being loaded.</param>
	'''<param name='addInInst'>Object representing this Add-in.</param>
	'''<remarks></remarks>
	Public Sub OnConnection(ByVal application As Object, ByVal connectMode As ext_ConnectMode, ByVal addInInst As Object, ByRef custom As Array) Implements IDTExtensibility2.OnConnection

		_application = CType(application, DTE2)
		_addininstance = CType(addInInst, AddIn)
		Dim commands As Commands2 = CType(_application.Commands, Commands2)

		DigPro.Application = _application

		_textcommands = New DPTextCommands(commands, _addininstance)
		_properties = New DPEnvDTEProperties()

		Try
			_toolbarutil = New DPToolBarsUtils()
			_menuitems = New DPMenus(_toolbarutil, _properties)

		Catch ex As Exception
			Try
				_textcommands.CreateCommands()

				_toolbarutil = New DPToolBarsUtils()

				Dim tlbar As CommandBar = _toolbarutil.GetToolbar(DPToolBarsUtils.TlBrName.TextCommands)
				If tlbar Is Nothing Then
					_toolbarutil.BuildDPMToolBars()
					tlbar = _toolbarutil.GetToolbar(DPToolBarsUtils.TlBrName.TextCommands)
				End If
				'For some reason the commands won't stick between sessions.  It seems that as a result the toolbar buttons
				'get destroyed as well.
				_textcommands.DebugMessage.AddControl(tlbar)
				_textcommands.CheckThrow.AddControl(tlbar, 2)

				_menuitems = New DPMenus(_toolbarutil, _properties)

			Catch ex2 As Exception
				MsgBox("An exception occurred and connection has aborted." + vbCrLf + vbCrLf + ex2.ToString(), CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, MsgBoxStyle), "Visual Studio Tools")
				Exit Sub
			End Try

		End Try

		'Seem to have to run this every time as well.
		_properties.SetKeyboardScheme()

	End Sub

	'''<summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
	'''<param name='disconnectMode'>Describes how the Add-in is being unloaded.</param>
	'''<param name='custom'>Array of parameters that are host application specific.</param>
	'''<remarks></remarks>
	Public Sub OnDisconnection(ByVal disconnectMode As ext_DisconnectMode, ByRef custom As Array) Implements IDTExtensibility2.OnDisconnection
	End Sub

	'''<summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification that the collection of Add-ins has changed.</summary>
	'''<param name='custom'>Array of parameters that are host application specific.</param>
	'''<remarks></remarks>
	Public Sub OnAddInsUpdate(ByRef custom As Array) Implements IDTExtensibility2.OnAddInsUpdate
	End Sub

	'''<summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
	'''<param name='custom'>Array of parameters that are host application specific.</param>
	'''<remarks></remarks>
	Public Sub OnStartupComplete(ByRef custom As Array) Implements IDTExtensibility2.OnStartupComplete
	End Sub

	'''<summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
	'''<param name='custom'>Array of parameters that are host application specific.</param>
	'''<remarks></remarks>
	Public Sub OnBeginShutdown(ByRef custom As Array) Implements IDTExtensibility2.OnBeginShutdown
	End Sub

	'''<summary>Implements the QueryStatus method of the IDTCommandTarget interface. This is called when the command's availability is updated</summary>
	'''<param name='commandName'>The name of the command to determine state for.</param>
	'''<param name='neededText'>Text that is needed for the command.</param>
	'''<param name='status'>The state of the command in the user interface.</param>
	'''<param name='commandText'>Text requested by the neededText parameter.</param>
	'''<remarks></remarks>
	Public Sub QueryStatus(ByVal commandName As String, ByVal neededText As vsCommandStatusTextWanted, ByRef status As vsCommandStatus, ByRef commandText As Object) Implements IDTCommandTarget.QueryStatus
		If neededText = vsCommandStatusTextWanted.vsCommandStatusTextWantedNone Then
			If commandName = "VisualStudioTools.Connect.VisualStudioTools" Then
				status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
			Else
				'status = vsCommandStatus.vsCommandStatusUnsupported
				'Try the text commands.
				_textcommands.QueryStatus(commandName, neededText, status, commandText)
			End If
		End If
	End Sub

	'''<summary>Implements the Exec method of the IDTCommandTarget interface. This is called when the command is invoked.</summary>
	'''<param name='commandName'>The name of the command to execute.</param>
	'''<param name='executeOption'>Describes how the command should be run.</param>
	'''<param name='varIn'>Parameters passed from the caller to the command handler.</param>
	'''<param name='varOut'>Parameters passed from the command handler to the caller.</param>
	'''<param name='handled'>Informs the caller if the command was handled or not.</param>
	'''<remarks></remarks>
	Public Sub Exec(ByVal commandName As String, ByVal executeOption As vsCommandExecOption, ByRef varIn As Object, ByRef varOut As Object, ByRef handled As Boolean) Implements IDTCommandTarget.Exec
		handled = False
		If executeOption = vsCommandExecOption.vsCommandExecOptionDoDefault Then
			If commandName = "VisualStudioTools.Connect.VisualStudioTools" Then
				handled = True
				Exit Sub
			End If
		End If

		'Try the text commands.
		_textcommands.Exec(commandName, executeOption, varIn, varOut, handled)
		If handled Then
			Exit Sub
		End If

	End Sub
End Class
