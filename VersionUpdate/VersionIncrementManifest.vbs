rcfile = "ConnectionCom.manifest"

Set fso = CreateObject("Scripting.FileSystemObject")
Set re  = New RegExp
re.Global = True

Function IncMaint(m, g1, g2, pos, src)
  IncMaint = g1 & (CInt(g2)+1)
End Function

rctext = fso.OpenTextFile(rcfile).ReadAll

re.Pattern = "((?:assemblyIdentity\b.*\bversion=.\b)\d+.\d+.\d+.)(\d+)"
rctext = re.Replace(rctext, GetRef("IncMaint"))

' Both pattern below should work. as version mentioned at few places I would prefer pattern1
' Pattern1 for manifest file ((?:assemblyIdentity\b.*\bversion=.\b)\d+.\d+.\d+.)(\d+)
' Pattern2 for manifest file ((?:version=.\b)\d+.\d+.\d+.)(\d+)


re.Pattern = "((?:PRODUCTVERSION|FILEVERSION) \d+,\d+,\d+,)(\d+)"
rctext = re.Replace(rctext, GetRef("IncMaint"))

re.Pattern = "(""(?:ProductVersion|FileVersion)"", ""\d+, \d+, )(\d+)(, \d+"")"
rctext = re.Replace(rctext, GetRef("IncMaint"))

fso.OpenTextFile(rcfile, 2).Write rctext