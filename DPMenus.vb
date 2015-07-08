Imports EnvDTE80
Imports EnvDTE
Imports Microsoft.VisualStudio.CommandBars

Public Class DPMenus

    'Visual Studio Application.
	Private _toolbars As DPToolBarsUtils
    Private _properties As DPEnvDTEProperties

    'Fly out menu from "tools".
    Private _dpmmenu As CommandBarPopup

    'Menu items.
	Private WithEvents _buildtoolbars As CommandBarButton
	Private WithEvents _setkeyboardshortcuts As CommandBarButton
    Private WithEvents _gettoolbarids As CommandBarButton
    Private WithEvents _listallcommands As CommandBarButton
    Private WithEvents _applyfontcolors As CommandBarButton
    Private WithEvents _resetkeyboard As CommandBarButton
    Private WithEvents _setrecentcount As CommandBarButton
    Private WithEvents _displayshortcuts As CommandBarButton

	Public Sub New(ByVal toolbarsutils As DPToolBarsUtils, ByVal properties As DPEnvDTEProperties)
		_toolbars = toolbarsutils
		_properties = properties
		CreateMenuItems()
	End Sub

    Private Sub CreateMenuItems()

        Dim commandbars As CommandBars = CType(DigPro.Application.CommandBars, CommandBars)

        Dim menuBarCommandBar As CommandBar = commandbars.Item("MenuBar")

        'Find the Tools command bar on the MenuBar command bar:
        Dim toolscontrol As CommandBarControl = menuBarCommandBar.Controls.Item("Tools")
        Dim toolspopup As CommandBarPopup = CType(toolscontrol, CommandBarPopup)

        'Try to find the DPM menu on the "tools" menu item.
        Dim foundmenu As Boolean = False
        For Each menuitem As CommandBarControl In toolspopup.Controls

            If foundmenu Then
                menuitem.BeginGroup = True
                GoTo FoundMenu
            End If

            If (menuitem.Caption = DigPro.CompanyName) Then
                _dpmmenu = CType(menuitem, CommandBarPopup)
                foundmenu = True
            End If

        Next menuitem

        'If it isn't there then create it.
		_dpmmenu = CType(toolspopup.Controls.Add(Type:=MsoControlType.msoControlPopup, Before:=1, Temporary:=True), CommandBarPopup)

        'We want the next item to have the value begin a new group.
        toolspopup.Controls.Item(2).BeginGroup = True

FoundMenu:
        _dpmmenu.BeginGroup = True

        With _dpmmenu
            .Caption = DigPro.CompanyName
            .BeginGroup = True
        End With

        _buildtoolbars = CType(_dpmmenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
        With _buildtoolbars
            .Caption = "Build Custom Toolbars"
            .Visible = True
            .BeginGroup = True
		End With

		_setkeyboardshortcuts = CType(_dpmmenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
		With _setkeyboardshortcuts
			.Caption = "Set Keyboard Shortcuts"
			.Visible = True
			.BeginGroup = False
		End With

        _gettoolbarids = CType(_dpmmenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
        With _gettoolbarids
            .Caption = "Get Toolbar Button Info"
            '.Visible = True
            .Visible = False
        End With

        _listallcommands = CType(_dpmmenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
        With _listallcommands
            .Caption = "List All Commands in App"
            '.Visible = True
            .Visible = False
        End With

        _applyfontcolors = CType(_dpmmenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
        With _applyfontcolors
            .Caption = "Apply Fonts and Colors"
            .Visible = True
        End With

        '_resetkeyboard = CType(_dpmmenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
        'With _resetkeyboard
        '    .Caption = "(Re)Set the Keyboard Scheme"
        '    .Visible = True
        'End With

        _setrecentcount = CType(_dpmmenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
        With _setrecentcount
            .Caption = "Set Recent Menu Items = 10"
            .Visible = True
        End With

        _displayshortcuts = CType(_dpmmenu.Controls.Add(Type:=MsoControlType.msoControlButton, Temporary:=True), CommandBarButton)
        With _displayshortcuts
            .Caption = "Write Keyboard Shortcuts to an HTML File"
            .Visible = True
        End With

    End Sub

    Private Sub onBuildToolbars_Click(ByVal Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _buildtoolbars.Click
        _toolbars.BuildDPMToolBars()
	End Sub

	Private Sub onSetKeyboarShortCuts_Click(ByVal Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _setkeyboardshortcuts.Click
		_properties.SetKeyboardScheme()
	End Sub

    Private Sub onGetToolbarIDs_Click(ByVal Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _gettoolbarids.Click
        _toolbars.GetToolbarIDs()
    End Sub

    Private Sub onListAllCommands(ByVal Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _listallcommands.Click
        _toolbars.ListAllCommands()
    End Sub

    Private Sub onApplyFontsColors(ByVal Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _applyfontcolors.Click
        _properties.ApplyFontColors()
    End Sub

    Private Sub onSetRecentCount(ByVal Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _setrecentcount.Click
        _properties.SetRecentUsedSize()
    End Sub

    Public Sub onListShortcutsInHTML(ByVal Ctrl As CommandBarButton, ByRef CancelDefault As Boolean) Handles _displayshortcuts.Click

        'Declare a StreamWriter
        Dim sw As System.IO.StreamWriter
		sw = New System.IO.StreamWriter("c:\\temp\\Visual Studio Shortcuts.html")

        'Write the beginning HTML
        WriteHTMLStart(sw)

        ' Add a row for each keyboard shortcut.
        ' Add a row for each keyboard shortcut
        For Each c As Command In DigPro.Application.Commands
            If c.Name <> "" Then
                Dim bindings As System.Array
                bindings = CType(c.Bindings, System.Array)
                For i As Integer = 0 To bindings.Length - 1
                    sw.WriteLine("<tr>")
                    sw.WriteLine("<td>" + c.Name + "</td>")
					sw.WriteLine("<td>" + bindings(i).ToString() + "</td>")
                    sw.WriteLine("</tr>")
                Next

            End If
        Next


        'I want to do something like to following to sort the commands by their name before
        'printing, however, none of the standard collections allow multiple keys so something
        'custom would have to be implement.


        'Dim commandstrings As System.Collections.Generic.SortedList(Of String, String)
        'commandstrings = New System.Collections.Generic.SortedList(Of String, String)

        ''Dim commands As Commands2 = CType(DigPro.Application.Commands, Commands2)
        'For Each c As Command In DigPro.Application.Commands
        '    If c.Name <> "" Then

        '        Dim bindings As System.Array
        '        bindings = CType(c.Bindings, System.Array)

        '        For i As Integer = 0 To bindings.Length - 1
        '            Dim name As String = c.Name
        '            commandstrings.Add(c.Name, bindings(i))
        '        Next

        '    End If
        'Next

        'For Each commandstring As System.Collections.Generic.KeyValuePair(Of String, String) In commandstrings
        '    sw.WriteLine("<tr>")
        '    sw.WriteLine("<td>" + commandstring.Key + "</td>")
        '    sw.WriteLine("<td>" + commandstring.Value + "</td>")
        '    sw.WriteLine("</tr>")

        'Next


        'Write the end HTML
        WriteHTMLEnd(sw)

        'Flush and close the stream
        sw.Flush()
        sw.Close()
    End Sub

    Public Sub WriteHTMLStart(ByVal sw As System.IO.StreamWriter)
        sw.WriteLine("<html>")
        sw.WriteLine("<head>")
        sw.WriteLine("<title>")

        sw.WriteLine("Visual Studio Keyboard Shortcuts")
        sw.WriteLine("</title>")
        sw.WriteLine("</head>")

        sw.WriteLine("<body>")
        sw.WriteLine("<h1>Visual Studio 2005 Keyboard Shortcuts</h1>")
        sw.WriteLine("<font size=""2"" face=""Verdana"">")
        sw.WriteLine("<table border=""1"">")
        sw.WriteLine("<tr BGCOLOR=""#018FFF""><td align=""center""><b>Command</b></td><td align=""center""><b>Shortcut</b></td></tr>")
    End Sub

    Public Sub WriteHTMLEnd(ByVal sw As System.IO.StreamWriter)
        sw.WriteLine("</table>")
        sw.WriteLine("</font>")
        sw.WriteLine("</body>")
        sw.WriteLine("</html>")
    End Sub

End Class
