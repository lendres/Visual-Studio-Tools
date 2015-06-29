Imports System
Imports Microsoft.VisualStudio.CommandBars
Imports Extensibility
Imports EnvDTE
Imports EnvDTE80
Public Class DPTextCommands

	Private _applicationcommands As Commands2
	Private _addininstance As AddIn

	Private Shared _commandnames(CommandName.Size) As String
	Private Shared _keyboardmappings(CommandName.Size) As String
	Private Shared _arraysinitialized As Boolean = False

	Public DebugMessage As Command
	Public CheckThrow As Command

	Public Enum CommandName
		First = 0
		PrintDebugMessage = First
		PrintCheckThrow
		CPointer
		CurlyBraces
		SquareBraces
		ReverseEquals
		Last = ReverseEquals
		Size = Last
	End Enum

	Public Sub New(applicationcommands As Commands2, addininstance As AddIn)
		_applicationcommands = applicationcommands
		_addininstance = addininstance
	End Sub

	Public Shared Function GetCommandName(ByVal commandname As CommandName) As String
		InitializeArrays()
		Return _commandnames(commandname)
	End Function

	Public Shared Function GetCommandConnectionString(ByVal commandname As CommandName) As String
		InitializeArrays()
		Return DigPro.ConnectString + "." + _commandnames(commandname)
	End Function

	Public Shared Function GetKeyboardMapping(ByVal commandname As CommandName) As String
		InitializeArrays()
		Return _keyboardmappings(commandname)
	End Function

	Private Shared Sub InitializeArrays()

		If Not _arraysinitialized Then

			_commandnames(CommandName.PrintDebugMessage) = "PrintDebugMessage"
			_commandnames(CommandName.PrintCheckThrow) = "PrintCheckThrow"
			_commandnames(CommandName.CPointer) = "CPointer"
			_commandnames(CommandName.CurlyBraces) = "CurlyBraces"
			_commandnames(CommandName.SquareBraces) = "SquareBraces"
			_commandnames(CommandName.ReverseEquals) = "ReverseEquals"

			_keyboardmappings(CommandName.PrintDebugMessage) = "Text Editor::Alt+D,Alt+M"
			_keyboardmappings(CommandName.PrintCheckThrow) = "Text Editor::Alt+C,Alt+T"
			_keyboardmappings(CommandName.CPointer) = "Text Editor::Alt+."
			_keyboardmappings(CommandName.CurlyBraces) = "Text Editor::Shift+Alt+["
			_keyboardmappings(CommandName.SquareBraces) = "Text Editor::Alt+["
			_keyboardmappings(CommandName.ReverseEquals) = "Text Editor::Alt+="

			_arraysinitialized = True

		End If

	End Sub

	Public Sub SetShortCutKeys()
	End Sub

	Public Sub CreateCommands()
		DebugMessage = _applicationcommands.AddNamedCommand2(_addininstance, DPTextCommands.GetCommandName(DPTextCommands.CommandName.PrintDebugMessage), "Dbg Msg", "Prints a #pragma message(...) statement.", True, 59, Nothing, CType(vsCommandStatus.vsCommandStatusSupported, Integer) + CType(vsCommandStatus.vsCommandStatusEnabled, Integer), vsCommandStyle.vsCommandStyleText, vsCommandControlType.vsCommandControlTypeButton)


		'Dim command As Command = commands.AddNamedCommand2(_addininstance, "VisualStudioTools", "VisualStudioTools", "Executes the command for Visual Studio Tools", True, 59, Nothing, CType(vsCommandStatus.vsCommandStatusSupported, Integer) + CType(vsCommandStatus.vsCommandStatusEnabled, Integer), vsCommandStyle.vsCommandStylePictAndText, vsCommandControlType.vsCommandControlTypeButton)

		CheckThrow = _applicationcommands.AddNamedCommand2(_addininstance, DPTextCommands.GetCommandName(DPTextCommands.CommandName.PrintCheckThrow), "Chk Thrw", "Prints a CHECK_THROW(..) statement.", True, 59, Nothing, CType(vsCommandStatus.vsCommandStatusSupported, Integer) + CType(vsCommandStatus.vsCommandStatusEnabled, Integer), vsCommandStyle.vsCommandStyleText, vsCommandControlType.vsCommandControlTypeButton)

		_applicationcommands.AddNamedCommand2(_addininstance, DPTextCommands.GetCommandName(DPTextCommands.CommandName.CPointer), "->", "Prints a C pointer (->)", True, , Nothing, 1 + 2)
		_applicationcommands.AddNamedCommand2(_addininstance, DPTextCommands.GetCommandName(DPTextCommands.CommandName.CurlyBraces), "{..}", "Prints a curly braces on separate lines.", True, , Nothing, 1 + 2)
		_applicationcommands.AddNamedCommand2(_addininstance, DPTextCommands.GetCommandName(DPTextCommands.CommandName.SquareBraces), "[..]", "Prints square braces around the last word before the cursor location.", True, , Nothing, 1 + 2)
		_applicationcommands.AddNamedCommand2(_addininstance, DPTextCommands.GetCommandName(DPTextCommands.CommandName.ReverseEquals), "x=y->y=x", "Swaps the left and right hand sides of an equal sign.", True, , Nothing, 1 + 2)
	End Sub

	Public Sub PrintDebugMessage()
		Try
			Dim textdoc As TextDocument = CType(DigPro.Application.ActiveDocument.Object("TextDocument"), TextDocument)

			Dim docname As String = DigPro.Application.ActiveDocument.Name.ToLower()

			Dim lastdot As Integer = docname.LastIndexOf(".")
			Dim extension As String = docname.Substring(lastdot, docname.Length - lastdot)

			Select Case extension
				Case ".cs"
					textdoc.Selection.Text = "// HACK: Debugging Code."

				Case ".cpp", ".c"
					textdoc.Selection.Text = "#pragma message(""DEBUGGING CODE:: File: "" __FILE__ "", Function: "" __FUNCTION__)"

			End Select

		Catch
		End Try

	End Sub

	Public Sub PrintCheckThrow()
		Try
			Dim textdoc As TextDocument = CType(DigPro.Application.ActiveDocument.Object("TextDocument"), TextDocument)
			textdoc.Selection.Text = "CHECK_THROW(false, EXCEPTION_ILLEGAL_USE, (""Reason"", ""Title""));"
		Catch
		End Try
	End Sub

	Public Sub CPointer()
		Try
			Dim textdoc As TextDocument = CType(DigPro.Application.ActiveDocument.Object("TextDocument"), TextDocument)
			textdoc.Selection.Text = "->"
		Catch
		End Try
	End Sub

	Public Sub CurlyBraces()
		Try
			Dim textdoc As TextDocument = CType(DigPro.Application.ActiveDocument.Object("TextDocument"), TextDocument)
			With textdoc.Selection
				'.Backspace()
				.Text = "{"
				.NewLine()
				.Backspace()
				.Text = "}"
				.LineUp(False, 1)
				.NewLine()
			End With
		Catch
		End Try
	End Sub

	Public Sub SquareBraces()

		Dim Cnt1 As Integer
		Cnt1 = 0

		Try
			Dim textdoc As TextDocument = CType(DigPro.Application.ActiveDocument.Object("TextDocument"), TextDocument)
			With textdoc
				.Selection.CharLeft(True, 1)
				Do While (.Selection.Text <> " ") And (.Selection.Text <> ";")
					.Selection.CharLeft(False, 1)
					.Selection.CharLeft(True, 1)
					Cnt1 = Cnt1 + 1

					'Ensure I don't back up over a semi-colon.  Use this as a test if the 
					'routine was called at the end or beginning of a line (perhaps by accident).
					'Should immediately test if there is a semi-colon, new-line or tab when the routine
					'is entered, by I don't know how the tab and new-line is represented; "\t" did not work.
					If (.Selection.Text = ";") Then
						.Selection.CharRight(False, Cnt1)
						.Selection.Text = "[]"
						Exit Sub
					End If
				Loop
				.Selection.Text = "["
				.Selection.CharRight(False, Cnt1)
				.Selection.Text = "]"
			End With
		Catch
		End Try
	End Sub

	Public Sub ReverseEquals()
		'Swap the left hand side and right hand side pieces of code around an equal sign.  Since the left and right side code might contain
		'the same variable or function/property names, et cetera, it is dangerous to try to do a direct find and replace of the strings.
		Try
			Dim textdoc As TextDocument = CType(DigPro.Application.ActiveDocument.Object("TextDocument"), TextDocument)
			With textdoc
				'Select the current line and copy the text to a string.
				.Selection.SelectLine()

				'Copy the line from the selection while remove ending spaces, tabs, carriage returns, and carriage return, line feeds.  The
				'selection of the lines seems to grab the "return" so we need to remove it.
				Dim line As String = .Selection.Text.TrimEnd(New Char() {ChrW(9), ChrW(10), ChrW(13)})
				Dim leftandright As String() = line.Split(New Char() {ChrW(9), "="c, ";"c}, StringSplitOptions.RemoveEmptyEntries)

				'Now that we've split them into the two parts, removing the equal sign and trailing semi-color, we remove all leading and trailing
				'spaces.  We only want the two pieces of code on either side of the equal sign.
				leftandright(0) = leftandright(0).Trim()
				leftandright(1) = leftandright(1).Trim()

				'Split the line at the equal sign, removing the equal sign.
				Dim halves As String() = line.Split("="c)

				'Replace the strings in the two halves.
				halves(0) = halves(0).Replace(leftandright(0), leftandright(1))
				halves(1) = halves(1).Replace(leftandright(1), leftandright(0))

				'Reassemble the string from the two halves which had the strings swapped.
				'We also never have extraneous blank/white space at the end of a line, so we might as well kill that while we are here.
				.Selection.Insert(halves(0) + "=" + halves(1).TrimEnd() + vbCrLf)
			End With
		Catch
		End Try

	End Sub

	Public Sub Exec(ByVal cmmndname As String, ByVal executeOption As vsCommandExecOption, ByRef varIn As Object, ByRef varOut As Object, ByRef handled As Boolean)

		handled = False

		If (executeOption = vsCommandExecOption.vsCommandExecOptionDoDefault) Then

			If cmmndname = DPTextCommands.GetCommandConnectionString(CommandName.PrintDebugMessage) Then
				PrintDebugMessage()
				handled = True
				Exit Sub
			End If

			If cmmndname = DPTextCommands.GetCommandConnectionString(CommandName.PrintCheckThrow) Then
				PrintCheckThrow()
				handled = True
				Exit Sub
			End If

			If cmmndname = DPTextCommands.GetCommandConnectionString(CommandName.CPointer) Then
				CPointer()
				handled = True
				Exit Sub
			End If

			If cmmndname = DPTextCommands.GetCommandConnectionString(CommandName.CurlyBraces) Then
				CurlyBraces()
				handled = True
				Exit Sub
			End If

			If cmmndname = DPTextCommands.GetCommandConnectionString(CommandName.SquareBraces) Then
				SquareBraces()
				handled = True
				Exit Sub
			End If

			If cmmndname = DPTextCommands.GetCommandConnectionString(CommandName.ReverseEquals) Then
				ReverseEquals()
				handled = True
				Exit Sub
			End If

		End If
	End Sub

	Public Sub QueryStatus(ByVal cmdName As String, ByVal neededText As vsCommandStatusTextWanted, ByRef status As vsCommandStatus, ByRef commandText As Object)

		If neededText = EnvDTE.vsCommandStatusTextWanted.vsCommandStatusTextWantedNone Then

			For i As CommandName = CommandName.First To CommandName.Last
				If cmdName = DPTextCommands.GetCommandConnectionString(i) Then
					If DigPro.Application.ActiveDocument Is Nothing Then
						status = vsCommandStatus.vsCommandStatusUnsupported
					Else
						status = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
					End If
					Exit Sub
				End If

			Next

			status = vsCommandStatus.vsCommandStatusUnsupported

		End If
	End Sub

End Class
