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
		FormatVariableDeclarations
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
			_commandnames(CommandName.FormatVariableDeclarations) = "FormatVariableDeclarations"
			_commandnames(CommandName.ReverseEquals) = "ReverseEquals"

			_keyboardmappings(CommandName.PrintDebugMessage) = "Text Editor::Alt+D,Alt+M"
			_keyboardmappings(CommandName.PrintCheckThrow) = "Text Editor::Alt+C,Alt+T"
			_keyboardmappings(CommandName.CPointer) = "Text Editor::Alt+."
			_keyboardmappings(CommandName.CurlyBraces) = "Text Editor::Shift+Alt+["
			_keyboardmappings(CommandName.SquareBraces) = "Text Editor::Alt+["
			_keyboardmappings(CommandName.FormatVariableDeclarations) = "Text Editor::Alt+\"
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
		_applicationcommands.AddNamedCommand2(_addininstance, DPTextCommands.GetCommandName(DPTextCommands.CommandName.FormatVariableDeclarations), "Align Variables", "Aligns the variable names in a selection of variable declarations.", True, , Nothing, 1 + 2)
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
	Public Sub FormatVariableDeclarations()
		'Align the variable names in a selection of variable declarations.
		'Try
		Dim textdoc As TextDocument = CType(DigPro.Application.ActiveDocument.Object("TextDocument"), TextDocument)
		Dim objectSelectedText As TextSelection = textdoc.Selection

		'http://therightstuff.de/2010/01/24/Visual-Studio-Tip-Setting-Indent-Width-And-TabsSpaces-Quickly-Using-Macros.aspx
		'http://www.jamesralexander.com/blog/content/visual-studio-toggle-between-leading-tabs-or-spaces-project
		Dim textEditorProperties As EnvDTE.Properties = GetCurrentLanguageProperties()
		Dim inserttabs As Boolean = CBool(textEditorProperties.Item("InsertTabs").Value)
		Dim tabsize As Integer = CInt(textEditorProperties.Item("TabSize").Value)

		Dim maxlinestart As Integer = 0
		Dim maxlineend As Integer = 0

		'Break the text up into separate lines by splitting the string at the line feed character.
		'Chr(10): [Line Feed Return] (vbLf)
		Dim lines As String() = objectSelectedText.Text.Split(New Char() {ChrW(10)})
		Dim numberoflines As Integer = lines.Length

		'We cannot align things if there is only 1 line.
		If numberoflines < 2 Then
			Return
		End If

		'Extract the prefix we will add in later.
		Dim lineprefix As String = GetInitialLineString(lines, inserttabs, tabsize)

		Dim startsoflines(numberoflines - 1) As String
		Dim variablenames(numberoflines - 1) As String
		Dim endsoflines(numberoflines - 1) As String


		For i As Integer = 0 To numberoflines - 1

			'Select the current line and copy the text to a string.
			'objectSelectedText.GotoLine(i, True)

			'Copy the line from the selection while remove ending spaces, tabs, carriage returns, and carriage return, line feeds.  The
			'selection of the lines seems to grab the "return" so we need to remove it.
			Dim line As String = RemoveTrailingWhiteSpaceAndLineReturns(lines(i))

			If line = "" Then
				'If the line was blank, put in blanks in the output.
				startsoflines(i) = ""
				variablenames(i) = ""
				endsoflines(i) = ""
			Else
				'If the line contains an initialization, we will extract it.
				If line.Contains("=") Then
					Dim equalsignindex As Integer = line.LastIndexOf("=")
					endsoflines(i) = line.Substring(equalsignindex, line.Length - equalsignindex)
					line = line.Substring(0, equalsignindex).TrimEnd()
				Else
					endsoflines(i) = ""
				End If

				'Line wasn't blank, so do the work of parsing the line.
				Dim lastspace As Integer = line.LastIndexOf(ChrW(32))
				Dim lasttab As Integer = line.LastIndexOf(ChrW(9))

				Dim startoflastword As Integer = lastspace
				If lasttab > lastspace Then
					startoflastword = lasttab
				End If

				startsoflines(i) = RemoveLeadingWhiteSpace(line.Substring(0, startoflastword))
				variablenames(i) = line.Substring(startoflastword + 1, line.Length - startoflastword - 1)

				Dim lengthoflineinspaces As Integer = startsoflines(i).Length
				If lengthoflineinspaces > maxlinestart Then
					maxlinestart = lengthoflineinspaces
				End If

				lengthoflineinspaces = NumberOfEquivalentSpaces(variablenames(i), tabsize)
				If lengthoflineinspaces > maxlineend Then
					maxlineend = lengthoflineinspaces
				End If
			End If
		Next

		'This is only used if we are inserting tabs, but we will calculate it here, outside of the for loop so we don't recalculate it on every loop.
		Dim variabletabposition As Integer = CInt(Math.Floor(maxlinestart / tabsize)) + lineprefix.Length + 3
		Dim initializationtabposition As Integer = CInt(Math.Floor(maxlineend / tabsize)) + 3

		'Clear the existing text and move to the start of the line.
		objectSelectedText.Text = ""
		objectSelectedText.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstColumn)

		For i As Integer = 0 To numberoflines - 1
			If startsoflines(i) = "" Then
				'This was a blank line, so we add a blank line back in.
				objectSelectedText.Insert(vbCrLf)
			Else
				If inserttabs Then
					Dim variabletabsneeded As Integer = variabletabposition - CInt(Math.Floor(startsoflines(i).Length / tabsize))
					Dim variabletabstring As String = ""
					For j As Integer = 1 To variabletabsneeded
						variabletabstring = variabletabstring + vbTab
					Next

					Dim initializationtabsneeded As Integer = initializationtabposition - CInt(Math.Floor(variablenames(i).Length / tabsize))
					Dim initializationtabstring As String = ""
					If endsoflines(i) <> "" Then
						For j As Integer = 1 To initializationtabsneeded
							initializationtabstring = initializationtabstring + vbTab
						Next
					End If

					objectSelectedText.Insert(lineprefix + startsoflines(i) + variabletabstring + variablenames(i) + initializationtabstring + endsoflines(i) + vbCrLf)
				Else
					'Insert the line using spaces only.
					If endsoflines(i) = "" Then
						objectSelectedText.Insert(String.Format("{0}{1, -" + (maxlinestart + 10).ToString() + "}{2}" + vbCrLf, lineprefix, startsoflines(i), variablenames(i)))
					Else
						objectSelectedText.Insert(String.Format("{0}{1, -" + (maxlinestart + 10).ToString() + "}{2, -" + (maxlineend + 10).ToString() + "}" + vbCrLf, lineprefix, startsoflines(i), variablenames(i), endsoflines(i)))
					End If
				End If

			End If
		Next


		'Catch
		'End Try
	End Sub
	Private Function GetCurrentLanguageProperties() As EnvDTE.Properties
		If DigPro.Application.ActiveDocument.Language = "CSharp" Then
			GetCurrentLanguageProperties = DigPro.Application.Properties("TextEditor", "CSharp")
		End If

		If DigPro.Application.ActiveDocument.Language = "SQL" Then
			GetCurrentLanguageProperties = DigPro.Application.Properties("TextEditor", "SQL")
		End If

		If DigPro.Application.ActiveDocument.Language = "HTML" Then
			GetCurrentLanguageProperties = DigPro.Application.Properties("TextEditor", "HTML")
		End If

		If DigPro.Application.ActiveDocument.Language = "JScript" Then
			GetCurrentLanguageProperties = DigPro.Application.Properties("TextEditor", "JScript")
		End If

	End Function

	Private Function NumberOfEquivalentSpaces(ByVal line As String, ByVal tabsize As Integer) As Integer
		'Chr(9):  [Tab]              (vbTab)
		Dim tablessstring As String = line.Replace(ChrW(9), "")
		Dim numberoftabs As Integer = line.Length - tablessstring.Length
		NumberOfEquivalentSpaces = numberoftabs * tabsize + line.Length
	End Function

	Private Function GetInitialLineString(ByVal lines As String(), ByVal inserttabs As Boolean, ByVal tabsize As Integer) As String

		'Find the first useable line to extract the indentation from.
		Dim exampleline As String = ""
		For i As Integer = 0 To lines.Length - 1
			If lines(i).Trim() <> "" Then
				If i = 0 Then
					'First line might not be entirely selected.
					exampleline = lines(1)
				Else
					exampleline = lines(i)
				End If
				Exit For
			End If
		Next

		If exampleline = "" Then
			Throw New Exception("No lines contain text.")
		End If

		Dim firstcharacter As String = exampleline.Substring(0, 1)

		Dim newline As String = RemoveTrailingWhiteSpaceAndLineReturns(exampleline)
		Dim totallength As Integer = newline.Length
		newline = RemoveLeadingWhiteSpace(newline)
		Dim whitespacelength As Integer = totallength - newline.Length

		'Convert between tabs and spaces, if required.
		If inserttabs Then
			'If we are inserting TABS, but the line started with spaces, we need to convert the spaces
			'to an equivalent tab size.
			If firstcharacter = " " Then
				whitespacelength = CInt(Math.Ceiling(whitespacelength / tabsize))
				firstcharacter = vbTab
			End If
		Else
			'If we are inserting SPACES, but the line started with spaces, we need to convert the tabs
			'to an equivalent space size.
			If firstcharacter = vbTab Then
				whitespacelength = whitespacelength * tabsize
				firstcharacter = " "
			End If
		End If

		'Add in the number of preceding white space characters required.
		GetInitialLineString = ""
		For i As Integer = 1 To whitespacelength
			GetInitialLineString = GetInitialLineString + firstcharacter
		Next

	End Function
	Private Function RemoveTrailingWhiteSpaceAndLineReturns(ByVal line As String) As String
		'ChrW(9):  [Tab]              (vbTab)
		'ChrW(10): [Line Feed Return] (vbLf)
		'ChrW(13): [Carriage Return]  (vbCr)
		'ChrW(32): [Space]
		RemoveTrailingWhiteSpaceAndLineReturns = line.TrimEnd(New Char() {ChrW(9), ChrW(10), ChrW(13), ChrW(32)})
	End Function
	Private Function RemoveLeadingWhiteSpace(ByVal line As String) As String
		'ChrW(9):  [Tab]              (vbTab)
		'ChrW(32): [Space]
		RemoveLeadingWhiteSpace = line.TrimStart(New Char() {ChrW(9), ChrW(32)})
	End Function

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

			If cmmndname = DPTextCommands.GetCommandConnectionString(CommandName.FormatVariableDeclarations) Then
				FormatVariableDeclarations()
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
