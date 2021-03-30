set fso = CreateObject("Scripting.FileSystemObject")

for each file in fso.GetFolder("1").Files
    extName=fso.GetExtensionName(file.name)
    baseName=fso.GetBaseName(file.name)
	baseSplit=split(baseName,"_")
    file.name=baseSplit(0) & "_" & baseSplit(1) & "_" & int((999999 - 100000 + 1) * rnd() + 100000) & "." & extName
next