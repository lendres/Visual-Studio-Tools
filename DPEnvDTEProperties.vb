Imports EnvDTE
Imports EnvDTE80

Public Class DPEnvDTEProperties
	'The code in this class is a little ugly, but it doesn't do too much.  I left
	'the ugly code as examples of what the objects are and how the access them.

	Private _coloritems As EnvDTE.FontsAndColorsItems

	Public Sub New()

		Dim props As EnvDTE.Properties = DigPro.Application.Properties("FontsAndColors", "TextEditor")
		Dim fontprop As EnvDTE.Property = props.Item("FontsAndColorsItems")

		_coloritems = CType(fontprop.Object, EnvDTE.FontsAndColorsItems)

	End Sub

	Public Sub ApplyFontColors()
		_coloritems.Item("String").Foreground = System.Convert.ToUInt32(255)
		_coloritems.Item("String (C# @ Verbatim)").Foreground = System.Convert.ToUInt32(255)
		_coloritems.Item("Number").Foreground = System.Convert.ToUInt32(32768)
		_coloritems.Item("User Types").Foreground = System.Convert.ToUInt32(128)
		_coloritems.Item("User Types (Delegates)").Foreground = System.Convert.ToUInt32(128)
		_coloritems.Item("User Types (Enums)").Foreground = System.Convert.ToUInt32(128)
		_coloritems.Item("User Types (Interfaces)").Foreground = System.Convert.ToUInt32(128)
		_coloritems.Item("User Types (Value types)").Foreground = System.Convert.ToUInt32(128)
	End Sub

	Public Sub SetRecentUsedSize()
		DigPro.Application.Properties("Environment", "General").Item("MRUListContainsNItems").Value = 10
	End Sub

	Public Sub SetKeyboardScheme()

		Dim cmds As Commands = DigPro.Application.Commands
		Dim cmd As Command

		'F7 for build.
		cmd = cmds.Item("Build.BuildSolution")
		cmd.Bindings = "Global::F7"
		'DigPro.Application.Properties("Environment", "Keyboard").Item("SchemeName").Value = "Visual C++ 6"

		'You're getting in my way you dirty dog.
		cmd = cmds.Item("Edit.IncreaseFilterLevel")
		cmd.Bindings = "Global::Ctrl+."

		'Add some additional short cuts.
		cmd = cmds.Item("Debug.DisableAllBreakpoints")
		cmd.Bindings = "Global::Ctrl+E,Ctrl+B"
		cmd = cmds.Item("Debug.EnableAllBreakpoints")
		cmd.Bindings = "Global::Ctrl+D,Ctrl+B"

		For i As DPTextCommands.CommandName = DPTextCommands.CommandName.First To DPTextCommands.CommandName.Last
			Dim s As String = DPTextCommands.GetCommandConnectionString(i)
			cmd = cmds.Item(DPTextCommands.GetCommandConnectionString(i))
			cmd.Bindings = DPTextCommands.GetKeyboardMapping(i)
		Next

	End Sub

End Class

