resgen VisualStudioTools.resx
Al.exe /embed:VisualStudioTools.resources /culture:en-US /out:VisualStudioTools.resources.dll
move VisualStudioTools.resources.dll .\bin\en-US\
del VisualStudioTools.resources
