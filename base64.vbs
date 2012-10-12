Dim oNode, BinaryStream
Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2

Set oNode = CreateObject("Msxml2.DOMDocument.3.0").CreateElement("base64")
oNode.dataType = "bin.base64"
oNode.text = "%%DATA%%"

'Create Stream object
Set BinaryStream = CreateObject("ADODB.Stream")

'Specify stream type - we want To save binary data.
BinaryStream.Type = adTypeBinary

'Open the stream And write binary data To the object
BinaryStream.Open
BinaryStream.Write oNode.nodeTypedValue
BinaryStream.SaveToFile "foo.txt", adSaveCreateOverWrite
