<div align="center">

## Get and Set attributes of a file


</div>

### Description

this is an easy code to get and set any file's attibutes (system, hidden, read-only,...) throught the vb function getatr and setattr
 
### More Info
 
- ofthis (string): the full filename

- tothis (long): the number associated with the attr. the correct values are in the code

ex: 44 = archive (32) + volume (8) + system (4)

getattributes returns a string that contains the attributes

setattributes returns a boolean : true = ok, false = error


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[themaxx](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/themaxx.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/themaxx-get-and-set-attributes-of-a-file__1-31894/archive/master.zip)





### Source Code

```
'*** Get and Set attributes of o file ***********
'*                       *
'************************************************
Public Function GetAttributes(OfThis As String) As String
Dim Tmp As VbFileAttribute
Tmp = GetAttr(OfThis)
If Tmp >= vbAlias Then '64
  GetAttributes = GetAttributes & " Alias"
  Tmp = Tmp - vbAlias
End If
If Tmp >= vbArchive Then ' 32
  GetAttributes = GetAttributes & " Archive"
  Tmp = Tmp - vbArchive
End If
If Tmp >= vbDirectory Then '16
  GetAttributes = GetAttributes & " directory"
  Tmp = Tmp - vbDirectory
End If
If Tmp >= vbVolume Then '8
  GetAttributes = GetAttributes & " volume"
  Tmp = Tmp - vbVolume
  End If
If Tmp >= vbSystem Then '4
  GetAttributes = GetAttributes & " System"
  Tmp = Tmp - vbSystem
End If
If Tmp >= vbHidden Then '2
  GetAttributes = GetAttributes & " Hidden"
  Tmp = Tmp - vbHidden
End If
If Tmp >= vbReadOnly Then '1
  GetAttributes = GetAttributes & " Read Only"
  Tmp = Tmp - vbReadOnly
End If
If Tmp = vbNormal Then '0
  GetAttributes = GetAttributes & " Normal"
  Tmp = Tmp - vbNormal
End If
End Function
Public Function SetAttributes(OfThis As String, ByVal ToThis As VbFileAttribute) As Boolean
 SetAttributes = True
 On Error GoTo errh
 SetAttr OfThis, ToThis
 GoTo fin
errh:
 SetAttributes = False
 Err.Clear
 Exit Function
fin:
 SetAttributes = True
End Function
```

