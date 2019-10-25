Imports Microsoft.VisualStudio.CommandBars
Imports EnvDTE
Imports EnvDTE80

Public Class DPToolBarsUtils

	'Toolbar results.
	Public Enum TlBrResult
		CreatedNew					'A new tool bar was created and returned.
		ReturnedExisting			'The tool bar was found to exist and was returned.
		ReturnedNothing				'The tool bar did not exist and could not be created (an error occurred).
	End Enum

	'Type of Button in TlBrControls.
	Public Enum TlBrButtonSource
		Command
		CreateNew
		CommandBarControl
	End Enum

	Public Enum TlBrName
		First = 0
		Standard = First
		TextEditing
		Build
		SolutionManagement
		TextCommands
		Outlining
		Last = Outlining
		Size = Last
	End Enum

	'Structure to pass data to create toolbar buttons.  If "BuiltIn" is true then the button is a built in
	'button for the application and does not have to be created new, just added to the toolbar.  In this
	'case the "ID" is given to determine which button to use.  Otherwise if "BuiltIn" is false the button
	'is added as a new button using everything except "BeginGroup" which is used by both options to specify
	'if the buttons begins a group.
	Public Structure TlBrControls
		Public Source As TlBrButtonSource
		Public CommandObj As Command
		Public CommandBarControlObj As CommandBarControl
		Public BeginGroup As Boolean
		Public ID As Integer
		Public Caption As String
		Public FaceID As Integer
		Public ClickHandler As Microsoft.VisualStudio.CommandBars._CommandBarButtonEvents_ClickEventHandler
		Public TooltipText As String
		Public DescriptionText As String
		Public Style As MsoButtonStyle
		Public Enabled As Boolean
		Public Visable As Boolean
		Public ControlType As MsoControlType

		Public Sub Initialize()
			Source = TlBrButtonSource.Command
			CommandObj = Nothing
			CommandBarControlObj = Nothing
			BeginGroup = False
			ID = 1
			Caption = ""
			FaceID = -1
			ClickHandler = Nothing
			TooltipText = ""
			DescriptionText = ""
			Style = MsoButtonStyle.msoButtonAutomatic
			Enabled = True
			Visable = True
			ControlType = MsoControlType.msoControlButton
		End Sub
	End Structure

	'To add a toolbar:
	'Add a new toolbar button array under the comment 'Toolbar buttons.
	'Add the toolbar name to the array ToolbarNames() in the function "New".
	'Add the toolbar buttons in the function "PopulateToolbarButtonArrays".

	'Toolbar names.
	Private ToolbarNames(TlBrName.Size) As String

	'Toolbar buttons
	Private TBButtonArrays(TlBrName.Size)() As TlBrControls
	Private TBStandardButtons(9) As TlBrControls
	Private TBTextEditingButtons(8) As TlBrControls
	Private TBBuildButtons(4) As TlBrControls
	Private TBSolutionManagement(2) As TlBrControls
	Private TBTextCommands(1) As TlBrControls
	Private TBOutliningButtons(1) As TlBrControls

	Public Sub New()
		ToolbarNames(TlBrName.Standard) = "DP Standard"
		ToolbarNames(TlBrName.TextEditing) = "DP Text Editing"
		ToolbarNames(TlBrName.Build) = "DP Build"
		ToolbarNames(TlBrName.SolutionManagement) = "DP Solution Management"
		ToolbarNames(TlBrName.TextCommands) = "DP Text Commands"
		ToolbarNames(TlBrName.Outlining) = "DP Outlining"

		'Fill in the button arrays.  Use a separate function so it can be done at the way bottom.
		PopulateToolbarButtonArrays()
	End Sub

	Public Sub BuildDPMToolBars()

		'Loop over all the toolbars to be built.
		For i As Integer = 0 To ToolbarNames.GetUpperBound(0)

			Dim title As String
			title = "Build DP Toolbars"

			'Try to create the toolbar.  If it exists, then we will ask the user if he/she wishes
			'to rebuild the toolbar.
			Dim result As DPToolBarsUtils.TlBrResult
			Dim toolbar As CommandBar = Nothing
			result = CreateNewBuiltInToolbar(toolbar, ToolbarNames(i), False)

			Select Case result
				Case DPToolBarsUtils.TlBrResult.CreatedNew
					PopulateToolbar(toolbar, TBButtonArrays(i))

				Case DPToolBarsUtils.TlBrResult.ReturnedExisting
					Dim cont As MsgBoxResult
					cont = MsgBox(Prompt:="The toolbar " + Chr(34) + ToolbarNames(i) + Chr(34) + _
									"already exists." + vbCrLf + "Do you wish to rebuild " + _
									"this toolbar?", Title:=title, Buttons:=CType(MsgBoxStyle.Question + MsgBoxStyle.YesNo, MsgBoxStyle))

					If cont = MsgBoxResult.Yes Then
						CreateNewBuiltInToolbar(toolbar, ToolbarNames(i), True)
						PopulateToolbar(toolbar, TBButtonArrays(i))
					End If

				Case DPToolBarsUtils.TlBrResult.ReturnedNothing
					MsgBox(Prompt:="The toolbar " + Chr(34) + ToolbarNames(i) + Chr(34) + _
									"could not be created.", Title:=title)
			End Select

			'Show the toolbar.
			toolbar.Visible = True

			'If ToolbarNames(i) = "DP Outlining" Then
			'End If

		Next i
	End Sub

	Public Sub GetToolbarIDs()
		Dim TlBrName As String
		Const title As String = "Get the IDs on a Toolbar"

		TlBrName = InputBox(Prompt:="Enter the name of the toolbar who's IDs you wish to list.", Title:=title)

		If TlBrName = "" Then
			Exit Sub
		End If

		Dim toolbar As CommandBar
		toolbar = GetToolbar(TlBrName)

		If toolbar Is Nothing Then
			MsgBox(Prompt:="That toolbar could not be located.", Title:=title)
			Exit Sub
		End If

		Try

			DigPro.Application.ItemOperations.NewFile("General\Text File")

			Dim textdoc As TextDocument = CType(DigPro.Application.ActiveDocument.Object("TextDocument"), TextDocument)
			With textdoc.Selection
				.Text = "Toolbar: " + toolbar.Name + vbCrLf

				For Each button As CommandBarControl In toolbar.Controls
					.NewLine()
					.Text = "Button: " + button.Caption + vbCrLf
					.Text = "ID: " + CStr(button.Id) + vbCrLf
					.Text = "Caption: " + button.Caption + vbCrLf
					'.Text = "OnAction: " + button.OnAction + vbCrLf        'This doesn't like being called on drop down buttons.
					.Text = "Parameter: " + button.Parameter + vbCrLf
					.Text = "BuiltIn: " + CStr(button.BuiltIn) + vbCrLf
					.Text = "Tooltip: " + button.TooltipText + vbCrLf
					.Text = "Tag: " + button.Tag + vbCrLf
					Dim oleuse As MsoControlOLEUsage
					oleuse = button.OLEUsage
					Select Case button.Type
						Case MsoControlType.msoControlComboBox
							.Text = "Type: ComboBox" + vbCrLf
							Dim cbutton As CommandBarComboBox
							cbutton = CType(button, CommandBarComboBox)
					End Select
				Next button

			End With

		Catch ex As Exception
			MsgBox(Prompt:="An error occurred while trying to read toolbar information.", Title:=title)
		End Try

	End Sub

	Public Sub ListAllCommands()

		'Verify the user wants to continue.
		Dim result As MsgBoxResult
		result = MsgBox(Prompt:="This command may take a while, do you wish to continue?", Title:="List All Commands", _
				Buttons:=CType(MsgBoxStyle.Question + MsgBoxStyle.YesNo, MsgBoxStyle))

		If result = MsgBoxResult.No Then
			Exit Sub
		End If

		DigPro.Application.ItemOperations.NewFile("General\Text File")

		Dim textdoc As TextDocument = CType(DigPro.Application.ActiveDocument.Object("TextDocument"), TextDocument)
		With textdoc.Selection

			.Text = "Listing Commands..." + vbCrLf

			For Each cmd As EnvDTE.Command In DigPro.Application.Commands
				.NewLine()
				.Text = "ID: " + CStr(cmd.ID) + vbCrLf
				.Text = "Guid: " + cmd.Guid + vbCrLf
				.Text = "LocalizedName: " + cmd.LocalizedName + vbCrLf
				.Text = "Name: " + cmd.Name + vbCrLf
			Next

		End With
	End Sub

	Public Function GetToolbar(ByVal name As TlBrName) As CommandBar
		Return GetToolbar(ToolbarNames(name))
	End Function

	Public Function GetToolbar(ByVal Name As String) As CommandBar
		Try
			Dim commandbars As CommandBars = CType(DigPro.Application.CommandBars, CommandBars)
			GetToolbar = CType(commandbars.Item(Name), CommandBar)
		Catch ex As Exception
			GetToolbar = Nothing
		End Try
	End Function

	Public Function CreateNewBuiltInToolbar(ByRef toolbar As CommandBar, ByVal name As String, ByVal overwriteexist As Boolean) As TlBrResult
		'This functions generates toolbars as if they were built into Visual Studio.  I.e. they cannot be deleted, but
		'they maintain there position and controls, et cetera.

		'Try to find the toolbar if it exists.
		Dim commandbars As CommandBars = CType(DigPro.Application.CommandBars, CommandBars)
		For Each appbar As CommandBar In commandbars
			If (appbar.Name = name) Then

				'Found the toolbar to exist.
				If (overwriteexist) Then
					'Over write the existing toolbar.
					Try
						DigPro.Application.Commands.RemoveCommandBar(appbar)
						toolbar = CType(DigPro.Application.Commands.AddCommandBar(name, _
									vsCommandBarType.vsCommandBarTypeToolbar, _
									Nothing, _
									MsoBarPosition.msoBarFloating), CommandBar)
						CreateNewBuiltInToolbar = TlBrResult.CreatedNew
					Catch ex As Exception
						toolbar = Nothing
						CreateNewBuiltInToolbar = TlBrResult.ReturnedNothing
					End Try
					Exit Function
				Else
					'Return the existing toolbar.
					toolbar = appbar
					CreateNewBuiltInToolbar = TlBrResult.ReturnedExisting
					Exit Function
				End If

			End If 'If current toolbar is the one specified on input.

		Next appbar

		'Doesn't exist so create it new.
		Try
			toolbar = CType(DigPro.Application.Commands.AddCommandBar(name, _
						vsCommandBarType.vsCommandBarTypeToolbar, _
						Nothing, _
						MsoBarPosition.msoBarFloating), CommandBar)
			CreateNewBuiltInToolbar = TlBrResult.CreatedNew
		Catch ex As Exception
			toolbar = Nothing
			CreateNewBuiltInToolbar = TlBrResult.ReturnedNothing
		End Try
	End Function

	Public Sub PopulateToolbar(ByRef toolbar As CommandBar, ByVal TBButtons As TlBrControls())
		For Each controldata As TlBrControls In TBButtons

			'Assign the button the value of Nothing to prevent a warning when compiling.
			Dim button As CommandBarButton = Nothing

			'If it is built in add it, otherwise create it new.
			Select Case controldata.Source
				Case TlBrButtonSource.Command
					button = CType(controldata.CommandObj.AddControl(toolbar, toolbar.Controls.Count() + 1), CommandBarButton)

				Case TlBrButtonSource.CreateNew
					button = CType(toolbar.Controls.Add(Type:=controldata.ControlType), CommandBarButton)

					With button
						.DescriptionText = controldata.DescriptionText
						.TooltipText = controldata.TooltipText

						Select Case controldata.ControlType
							Case MsoControlType.msoControlButton
								If Not (controldata.ClickHandler Is Nothing) Then
									Dim buttoncontrol As CommandBarButton
									buttoncontrol = CType(button, CommandBarButton)
									AddHandler buttoncontrol.Click, controldata.ClickHandler
								End If
						End Select

					End With

				Case TlBrButtonSource.CommandBarControl
					Try
						If Not (controldata.CommandBarControlObj Is Nothing) Then
							Dim cbutton As CommandBarControl = controldata.CommandBarControlObj.Copy(Bar:=toolbar)
							cbutton.BeginGroup = controldata.BeginGroup
							cbutton.Enabled = controldata.Enabled
							cbutton.Visible = controldata.Visable
							Continue For
						End If
					Catch ex As Exception
						MsgBox("An error occurred while Visual Studio Tools was trying to copy a built in tool bar button." + vbCrLf + vbCrLf + ex.ToString(), CType(MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, MsgBoxStyle), "Visual Studio Tools")
						Continue For
					End Try

				Case Else
					Throw New Exception("Toolbar button source incorrectly specified.")

			End Select

			'If you don't assign a value to button you get 'warning BC42104'.
			With button
				.BeginGroup = controldata.BeginGroup
				.Enabled = controldata.Enabled
				.Visible = controldata.Visable

				Select Case button.Type

					Case MsoControlType.msoControlButton
						Dim buttoncontrol As CommandBarButton
						buttoncontrol = CType(button, CommandBarButton)
						.Style = controldata.Style

				End Select

				'Allow for the caption to be overridden even on built in types.
				If controldata.Caption <> "" Then
					.Caption = controldata.Caption
				End If

				'Allow for Face ID to be over written even on built in types.
				If controldata.FaceID <> -1 Then
					.FaceId = controldata.FaceID
				End If

			End With

			'If toolbar.Name = "DP Outlining" Then
			'    Dim aStream As System.IO.Stream

			'    ' Call the Assembly GetManifestResourceStream method passing it

			'    ' the namespace of this project and the name of the Bitmap resource.

			'    ' Assign the Stream object returned to the aStream variable.

			'    aStream = Me.GetType().Assembly.GetManifestResourceStream("DPM Visual Studio Tools.CollapsetoDefinitions.bmp")

			'    ' Declare a variable of type Bitmap named aBitmap.

			'    ' Call the Bitmap New constructor passing in the aStream.

			'    ' Assign the address (reference) of the new Bitmap object

			'    ' to the aBitmap variable.

			'    Dim aBitmap As New System.Drawing.Imaging.Bitmap(aStream)

			'    ' Set this form's BackGround image property to aBitmap.


			'End If

		Next controldata


	End Sub

	Private Sub PopulateToolbarButtonArrays()

		TBButtonArrays(TlBrName.Standard) = TBStandardButtons
		TBButtonArrays(TlBrName.TextEditing) = TBTextEditingButtons
		TBButtonArrays(TlBrName.Build) = TBBuildButtons
		TBButtonArrays(TlBrName.SolutionManagement) = TBSolutionManagement
		TBButtonArrays(TlBrName.TextCommands) = TBTextCommands
		TBButtonArrays(TlBrName.Outlining) = TBOutliningButtons

		For i As Integer = 0 To TBStandardButtons.GetUpperBound(0)
			TBStandardButtons(i).Initialize()
		Next
		For i As Integer = 0 To TBTextEditingButtons.GetUpperBound(0)
			TBTextEditingButtons(i).Initialize()
		Next
		For i As Integer = 0 To TBBuildButtons.GetUpperBound(0)
			TBBuildButtons(i).Initialize()
		Next
		For i As Integer = 0 To TBSolutionManagement.GetUpperBound(0)
			TBSolutionManagement(i).Initialize()
		Next
		For i As Integer = 0 To TBTextCommands.GetUpperBound(0)
			TBTextCommands(i).Initialize()
		Next
		For i As Integer = 0 To TBOutliningButtons.GetUpperBound(0)
			TBOutliningButtons(i).Initialize()
		Next

		'Use this to access the arrays.  That way if something changes we don't
		'have to renumber everything because everything is indexed by this.
		Dim Index As Integer

		'Get the commands from the application.  Convert to Commands2 to get the
		'commands that were added with AddNamedCommand2.
		Dim commands As Commands2 = CType(DigPro.Application.Commands, Commands2)

		'Save All.
		Index = 0
		TBStandardButtons(Index).CommandObj = commands.Item("File.SaveAll")

		'Cut.
		Index = Index + 1
		TBStandardButtons(Index).CommandObj = commands.Item("Edit.Cut")
		TBStandardButtons(Index).BeginGroup = True

		'Copy.
		Index = Index + 1
		TBStandardButtons(Index).CommandObj = commands.Item("Edit.Copy")

		'Paste.
		Index = Index + 1
		TBStandardButtons(Index).CommandObj = commands.Item("Edit.Paste")

		'Undo.
		Index = Index + 1
		TBStandardButtons(Index).CommandObj = commands.Item("Edit.Undo")
		TBStandardButtons(Index).BeginGroup = True

		'Redo.
		Index = Index + 1
		TBStandardButtons(Index).CommandObj = commands.Item("Edit.Redo")

		'Start/Continue (Play).
		Index = Index + 1
		TBStandardButtons(Index).CommandObj = commands.Item("Debug.Start")
		TBStandardButtons(Index).BeginGroup = True

		'Break all (pause).
		Index = Index + 1
		TBStandardButtons(Index).CommandObj = commands.Item("Debug.BreakAll")

		'Stop.
		Index = Index + 1
		TBStandardButtons(Index).CommandObj = commands.Item("Debug.StopDebugging")


		'Solution configurations.  I cannot find it in the commands so instead of going to the trouble
		'of building my own I will just try to retrieve it from the "Standard" toolbar.
		Index = Index + 1
		TBStandardButtons(Index).Source = TlBrButtonSource.CommandBarControl
		TBStandardButtons(Index).BeginGroup = True

		Dim commandbars As CommandBars = CType(DigPro.Application.CommandBars, CommandBars)

		Try
			Dim standardbar As CommandBar
			standardbar = commandbars.Item("Standard")

			Dim solconfig As CommandBarControl
			solconfig = standardbar.Controls("Solution Configurations")

			TBStandardButtons(Index).CommandBarControlObj = solconfig

		Catch ex As Exception
			TBStandardButtons(Index).CommandBarControlObj = Nothing
		End Try

		'==============================================
		'TEXT EDITOR TOOLBAR.

		'Decrease line indent.
		Index = 0
		TBTextEditingButtons(Index).CommandObj = commands.Item("Edit.DecreaseLineIndent")

		'Increase line indent.
		Index = Index + 1
		TBTextEditingButtons(Index).CommandObj = commands.Item("Edit.IncreaseLineIndent")

		'Comment.
		Index = Index + 1
		TBTextEditingButtons(Index).CommandObj = commands.Item("Edit.CommentSelection")
		TBTextEditingButtons(Index).BeginGroup = True

		'Uncomment.
		Index = Index + 1
		TBTextEditingButtons(Index).CommandObj = commands.Item("Edit.UncommentSelection")

		'Format Selection
		Index = Index + 1
		TBTextEditingButtons(Index).CommandObj = commands.Item("Edit.FormatSelection")
		'TBTextEditingButtons(Index).Style = MsoButtonStyle.msoButtonCaption
		TBTextEditingButtons(Index).Caption = "Format &Selection"
		TBTextEditingButtons(Index).BeginGroup = True

		'Place book mark.
		Index = Index + 1
		TBTextEditingButtons(Index).CommandObj = commands.Item("Edit.ToggleBookmark")
		TBTextEditingButtons(Index).BeginGroup = True

		'Go to next book mark.
		Index = Index + 1
		TBTextEditingButtons(Index).CommandObj = commands.Item("Edit.NextBookmark")

		'Go to previous book mark.
		Index = Index + 1
		TBTextEditingButtons(Index).CommandObj = commands.Item("Edit.PreviousBookmark")

		'Clear book marks.
		Index = Index + 1
		TBTextEditingButtons(Index).CommandObj = commands.Item("Edit.ClearBookmarks")

		'==============================================
		'BUILD TOOLBAR.

		'Build project.
		Index = 0
		TBBuildButtons(Index).CommandObj = commands.Item("Build.BuildSelection")
		'Build.BuildOnlyProject

		'Compile.
		Index = Index + 1
		TBBuildButtons(Index).CommandObj = commands.Item("Build.Compile")

		'Build solution.
		Index = Index + 1
		TBBuildButtons(Index).CommandObj = commands.Item("Build.BuildSolution")
		TBBuildButtons(Index).BeginGroup = True

		'Batch build.
		Index = Index + 1
		TBBuildButtons(Index).CommandObj = commands.Item("Build.BatchBuild")
		TBBuildButtons(Index).Style = MsoButtonStyle.msoButtonCaption

		'Cancel build.
		Index = Index + 1
		TBBuildButtons(Index).CommandObj = commands.Item("Build.Cancel")
		TBBuildButtons(Index).BeginGroup = True

		'==============================================
		'SOLUTION MANAGEMENT TOOLBAR.

		'Open solution.
		Index = 0
		TBSolutionManagement(Index).CommandObj = commands.Item("File.OpenProject")

		'Recent projects.
		Index = Index + 1
		TBSolutionManagement(Index).Source = TlBrButtonSource.CommandBarControl

		Try
			Dim filemenu As CommandBarControl
			filemenu = commandbars.Item("MenuBar").Controls("File")

			Dim filepopup As CommandBarPopup
			filepopup = CType(filemenu, CommandBarPopup)

			Dim recentproj As CommandBarControl
			recentproj = filepopup.Controls("Recent Projects")

			TBSolutionManagement(Index).CommandBarControlObj = recentproj
			TBSolutionManagement(Index).Style = MsoButtonStyle.msoButtonCaption

		Catch ex As Exception
			TBSolutionManagement(Index).CommandBarControlObj = Nothing
		End Try

		'Close solution
		Index = Index + 1
		TBSolutionManagement(Index).CommandObj = commands.Item("File.CloseSolution")
		TBSolutionManagement(Index).BeginGroup = True

		'==============================================
		'TEXTCOMMANDS TOOLBAR.

		'Open solution.
		Index = 0
		TBTextCommands(Index).CommandObj = commands.Item(DPTextCommands.GetCommandConnectionString(DPTextCommands.CommandName.PrintDebugMessage))
		TBTextCommands(Index).BeginGroup = True
		TBTextCommands(Index).Style = MsoButtonStyle.msoButtonCaption

		Index = Index + 1
		TBTextCommands(Index).CommandObj = commands.Item(DPTextCommands.GetCommandConnectionString(DPTextCommands.CommandName.PrintCheckThrow))
		TBTextCommands(Index).BeginGroup = True
		TBTextCommands(Index).Style = MsoButtonStyle.msoButtonCaption

		'==============================================
		'OUTLINING TOOLBAR.

		Index = 0
		TBOutliningButtons(Index).CommandObj = commands.Item("Edit.StopHidingCurrent")
		TBOutliningButtons(Index).BeginGroup = True
		TBOutliningButtons(Index).Style = MsoButtonStyle.msoButtonCaption

		Index = Index + 1
		TBOutliningButtons(Index).CommandObj = commands.Item("Edit.CollapsetoDefinitions")
		TBOutliningButtons(Index).BeginGroup = False
		TBOutliningButtons(Index).Style = MsoButtonStyle.msoButtonCaption

		'Index = Index + 1
		'TBOutliningButtons(Index).CommandObj = commands.Item("Edit.CollapseAllincurrentblock")
		'TBOutliningButtons(Index).BeginGroup = False
		'TBOutliningButtons(Index).Style = MsoButtonStyle.msoButtonCaption

		'Index = Index + 1
		'TBOutliningButtons(Index).CommandObj = commands.Item("Edit.CollapseBlockcurrentblock")
		'TBOutliningButtons(Index).BeginGroup = False
		'TBOutliningButtons(Index).Style = MsoButtonStyle.msoButtonCaption

	End Sub

End Class

'ID: 256
'Guid: {C9DD4A59-47FB-11D2-83E7-00C04F9902C1}
'LocalizedName: Debug.Breakpoints
'Name: Debug.Breakpoints

'ID: 212
'Guid: {5EFC7975-14BC-11CF-9B2B-00AA00573819}
'LocalizedName: Debug.RunningDocuments
'Name: Debug.RunningDocuments

'ID: 16777984
'Guid: {C9DD4A59-47FB-11D2-83E7-00C04F9902C1}
'LocalizedName: Debug.Watch1
'Name: Debug.Watch1

'ID: 33555200
'Guid: {C9DD4A59-47FB-11D2-83E7-00C04F9902C1}
'LocalizedName: Debug.Watch2
'Name: Debug.Watch2

'ID: 50332416
'Guid: {C9DD4A59-47FB-11D2-83E7-00C04F9902C1}
'LocalizedName: Debug.Watch3
'Name: Debug.Watch3

'ID: 67109632
'Guid: {C9DD4A59-47FB-11D2-83E7-00C04F9902C1}
'LocalizedName: Debug.Watch4
'Name: Debug.Watch4

'ID: 747
'Guid: {5EFC7975-14BC-11CF-9B2B-00AA00573819}
'LocalizedName: Debug.Autos
'Name: Debug.Autos

'ID: 242
'Guid: {5EFC7975-14BC-11CF-9B2B-00AA00573819}
'LocalizedName: Debug.Locals
'Name: Debug.Locals

'ID: 748
'Guid: {5EFC7975-14BC-11CF-9B2B-00AA00573819}
'LocalizedName: Debug.This
'Name: Debug.This

'ID: 240
'Guid: {5EFC7975-14BC-11CF-9B2B-00AA00573819}
'LocalizedName: Debug.Immediate
'Name: Debug.Immediate

'ID: 243
'Guid: {5EFC7975-14BC-11CF-9B2B-00AA00573819}
'LocalizedName: Debug.CallStack
'Name: Debug.CallStack

'ID: 214
'Guid: {5EFC7975-14BC-11CF-9B2B-00AA00573819}
'LocalizedName: Debug.Threads
'Name: Debug.Threads