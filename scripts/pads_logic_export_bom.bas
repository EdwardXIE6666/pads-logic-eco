' =============================================================================
' PADS Logic 原理图元件导出（仅 Logic，不要在 PADS Layout 里运行）
'
' 适用：在已经打开的原理图里，把每个元件实例导出成 CSV，供外部 ECO 使用。
'       Layout 的对象是另一套，本脚本不适用于 Layout。
'
' 加载与运行：
'   1. 打开要导出的原理图，并先保存（未保存时 CSV 会写到临时目录）。
'   2. Tools → Basic Scripts → Basic Script Editor。
'   3. File → Open，选本 .BAS；或把内容贴进编辑器。
'   4. 按 Run（宏名 Main）。编辑器若只载入不自动执行，确认文件最后一行是 Main。
'   也可以 Tools → Basic Scripts，Add 这个 .BAS 后 Run。
'
' 改输出字段：
'   只改下面 EXTRA_ATTRS。固定列不要删、不要调换顺序：
'     RefDes, Part Type, Value, PCB Decal, Part Number, Description, Manufacturer
'   EXTRA_ATTRS 里的属性会追加在 Manufacturer 之后、Qty 之前。
'   外部 ECO 按表头名称取 RefDes / Part Type / Value，不按列号。
'   每个位号一行。Qty 是「相同 Part Type + Value + Part Number + PCB Decal」的个数。
'
' 对接 ECO：
'   把生成的 CSV 交给物料变更工具，作为原理图 BOM。
'   必须保留 RefDes、Part Type、Value 这三列的表头原文。
'   一位号一行，不要合并成位号列表，否则 ECO 无法对上单个位号。
' =============================================================================

Option Explicit

' 留空 = 原理图同目录，文件名 <图名>_BOM_日期时间.csv
' 也可写成完整路径，例如 C:\Temp\logic_bom.csv
Const OUTPUT_PATH = ""

' True 时在 CSV 之外再另存一份同名 .xlsx（本机需有 Excel）。CSV 失败则整次失败；Excel 失败只提示。
Const ALSO_EXCEL = False

' 追加属性，竖线分隔。脚本按这个名字在元件上找，找不到再找 Part Type。没有就留空。
Const EXTRA_ATTRS = "Tolerance|Package"

' 位号前缀，逗号分隔。例：TP,MH,FID 会跳过 TP1、MH2。留空表示不按前缀过滤。
Const SKIP_PREFIXES = ""

' Part Type 名称里包含这些文字就跳过，逗号分隔。例：GND,VCC,+5V。留空表示不过滤。
Const SKIP_PART_CONTAINS = ""

' True：只保留「字母开头、后面带数字」的位号，从而去掉 GND、VCC、+3V3 这类电源符号。
Const SKIP_POWER_SYMBOLS = False

Sub Main()
    On Error GoTo Fail

    Dim doc, folder, schName, csvPath, n
    Set doc = Nothing
    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo Fail
    If doc Is Nothing Then
        MsgBox "没有打开的原理图。请先打开 PADS Logic 文档再运行。", vbCritical, "PADS Logic BOM"
        Exit Sub
    End If

    folder = DocFolder(doc)
    schName = BaseName(doc)
    csvPath = ResolvePath(folder, schName)
    If csvPath = "" Then
        MsgBox "无法确定输出路径。请先保存原理图，或在脚本里填写 OUTPUT_PATH。", vbCritical, "PADS Logic BOM"
        Exit Sub
    End If

    n = ExportComponents(doc, csvPath)
    If n < 0 Then Exit Sub

    Dim excelNote
    excelNote = ""
    If ALSO_EXCEL Then
        If SaveExcel(csvPath) Then
            excelNote = vbCrLf & "Excel: " & XlsxPath(csvPath)
        Else
            excelNote = vbCrLf & "Excel 未生成。CSV 已写出。"
        End If
    End If

    MsgBox "导出完成，共 " & n & " 个元件。" & vbCrLf & csvPath & excelNote, vbInformation, "PADS Logic BOM"
    Exit Sub

Fail:
    MsgBox "导出失败：" & Err.Number & " " & Err.Description, vbCritical, "PADS Logic BOM"
End Sub

Function ExportComponents(doc, csvPath)
    ExportComponents = -1

    Dim comp, list(), qtyMap, i, n, kept
    n = 0
    On Error Resume Next
    For Each comp In doc.Components
        n = n + 1
    Next
    If Err.Number <> 0 Then
        MsgBox "读不到 ActiveDocument.Components。" & vbCrLf & Err.Description, vbCritical, "PADS Logic BOM"
        Exit Function
    End If
    On Error GoTo 0

    If n = 0 Then
        MsgBox "当前原理图没有元件。", vbExclamation, "PADS Logic BOM"
        Exit Function
    End If

    ReDim list(n - 1)
    Set qtyMap = CreateObject("Scripting.Dictionary")
    qtyMap.CompareMode = 1

    kept = 0
    On Error Resume Next
    For Each comp In doc.Components
        If Err.Number <> 0 Then Exit For
        Dim row
        Set row = ReadComponent(doc, comp)
        If Not row Is Nothing Then
            If Wanted(row("RefDes"), row("Part Type")) Then
                Set list(kept) = row
                Bump qtyMap, row("Key")
                kept = kept + 1
            End If
        End If
    Next
    On Error GoTo 0

    If kept = 0 Then
        MsgBox "过滤之后没有可导出的元件。请检查 SKIP_PREFIXES / SKIP_PART_CONTAINS / SKIP_POWER_SYMBOLS。", vbExclamation, "PADS Logic BOM"
        Exit Function
    End If

    Dim extras, headers, body, r
    extras = Split(EXTRA_ATTRS, "|")
    headers = HeaderLine(extras)
    body = headers & vbCrLf
    For i = 0 To kept - 1
        Set r = list(i)
        r("Qty") = CStr(qtyMap(r("Key")))
        body = body & RowLine(r, extras) & vbCrLf
    Next

    If Not WriteText(csvPath, body) Then
        MsgBox "CSV 没有写成功：" & csvPath, vbCritical, "PADS Logic BOM"
        Exit Function
    End If
    ExportComponents = kept
End Function

Function ReadComponent(doc, comp)
    Dim row, refdes, partType, ptObj
    Set ReadComponent = Nothing
    Set row = CreateObject("Scripting.Dictionary")
    row.CompareMode = 1

    refdes = CompName(comp)
    partType = PartTypeText(comp)
    Set ptObj = FindPartType(doc, partType)

    row("RefDes") = refdes
    row("Part Type") = partType
    row("Value") = AttrEither(comp, ptObj, Array("Value", "Val", "值"))
    row("PCB Decal") = PcbDecal(comp, ptObj)
    row("Part Number") = AttrEither(comp, ptObj, Array("Part Number", "PartNumber", "PART NUMBER", "PN", "Material Code", "物料编码"))
    row("Description") = AttrEither(comp, ptObj, Array("Description", "Desc", "描述"))
    row("Manufacturer") = AttrEither(comp, ptObj, Array("Manufacturer", "Mfr", "Vendor", "厂家"))
    row("Qty") = "1"
    row("Key") = partType & Chr(1) & row("Value") & Chr(1) & row("Part Number") & Chr(1) & row("PCB Decal")

    Dim extras, i, name
    extras = Split(EXTRA_ATTRS, "|")
    For i = 0 To UBound(extras)
        name = Trim(extras(i))
        If name <> "" Then row(name) = AttrEither(comp, ptObj, Array(name))
    Next
    Set ReadComponent = row
End Function

Function Wanted(refdes, partType)
    Wanted = False
    If Trim(refdes) = "" And Trim(partType) = "" Then Exit Function
    If HasPrefix(refdes, SKIP_PREFIXES) Then Exit Function
    If ContainsAny(partType, SKIP_PART_CONTAINS) Then Exit Function
    If SKIP_POWER_SYMBOLS Then
        If Not LooksLikeRefdes(refdes) Then Exit Function
    End If
    Wanted = True
End Function

Function HeaderLine(extras)
    Dim s, i, name
    s = "RefDes,Part Type,Value,PCB Decal,Part Number,Description,Manufacturer"
    For i = 0 To UBound(extras)
        name = Trim(extras(i))
        If name <> "" Then s = s & "," & Csv(name)
    Next
    HeaderLine = s & ",Qty"
End Function

Function RowLine(row, extras)
    Dim s, i, name
    s = Csv(row("RefDes")) & "," & Csv(row("Part Type")) & "," & Csv(row("Value")) & "," & _
        Csv(row("PCB Decal")) & "," & Csv(row("Part Number")) & "," & Csv(row("Description")) & "," & _
        Csv(row("Manufacturer"))
    For i = 0 To UBound(extras)
        name = Trim(extras(i))
        If name <> "" Then s = s & "," & Csv(DictGet(row, name))
    Next
    RowLine = s & "," & Csv(row("Qty"))
End Function

Function Csv(value)
    Dim s
    s = value & ""
    s = Replace(s, """", """""")
    If InStr(s, ",") > 0 Or InStr(s, """") > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then
        Csv = """" & s & """"
    Else
        Csv = s
    End If
End Function

Function CompName(comp)
    CompName = ""
    If comp Is Nothing Then Exit Function
    On Error Resume Next
    CompName = comp.Name & ""
    If CompName = "" Then CompName = comp.RefDes & ""
    Err.Clear
End Function

Function PartTypeText(comp)
    Dim s, obj
    PartTypeText = ""
    If comp Is Nothing Then Exit Function
    On Error Resume Next
    s = comp.PartType & ""
    If Err.Number = 0 And s <> "" Then
        PartTypeText = s
        Exit Function
    End If
    Err.Clear
    s = comp.PartTypeName & ""
    If Err.Number = 0 And s <> "" Then
        PartTypeText = s
        Exit Function
    End If
    Err.Clear
    Set obj = comp.PartType
    If Err.Number = 0 Then
        If Not obj Is Nothing Then PartTypeText = obj.Name & ""
    End If
    Err.Clear
End Function

Function Prop(obj, propName)
    Prop = ""
    If obj Is Nothing Or propName = "" Then Exit Function
    On Error Resume Next
    If propName = "PCBDecal" Then Prop = obj.PCBDecal & ""
    If Prop = "" And propName = "Decal" Then Prop = obj.Decal & ""
    Err.Clear
End Function

Function AttrEither(comp, ptObj, names)
    Dim s
    s = AttrOn(comp, names)
    If s = "" Then s = AttrOn(ptObj, names)
    AttrEither = s
End Function

Function AttrOn(obj, names)
    Dim attrs, attr, i, n
    AttrOn = ""
    If obj Is Nothing Then Exit Function
    On Error Resume Next
    Set attrs = obj.Attributes
    If Err.Number <> 0 Or attrs Is Nothing Then
        Err.Clear
        Set attrs = obj.Attrs
    End If
    If Err.Number <> 0 Or attrs Is Nothing Then
        Err.Clear
        Exit Function
    End If
    For i = 0 To UBound(names)
        n = names(i)
        Set attr = Nothing
        Set attr = attrs.Item(n)
        If Err.Number = 0 Then
            If Not attr Is Nothing Then
                AttrOn = attr.Value & ""
                If AttrOn = "" Then AttrOn = attr.Text & ""
                If AttrOn <> "" Then Exit Function
            End If
        End If
        Err.Clear
    Next
    For Each attr In attrs
        If Err.Number <> 0 Then Exit For
        For i = 0 To UBound(names)
            If StrComp(attr.Name & "", names(i), vbTextCompare) = 0 Then
                AttrOn = attr.Value & ""
                If AttrOn = "" Then AttrOn = attr.Text & ""
                Exit Function
            End If
        Next
    Next
    Err.Clear
End Function

Function PcbDecal(comp, ptObj)
    Dim s, d
    s = AttrEither(comp, ptObj, Array("PCB Decal", "PCBDecal", "Decal", "封装"))
    If s <> "" Then
        PcbDecal = s
        Exit Function
    End If
    s = Prop(ptObj, "PCBDecal")
    If s = "" Then s = Prop(ptObj, "Decal")
    If s = "" Then s = Prop(comp, "PCBDecal")
    If s = "" Then s = Prop(comp, "Decal")
    If s = "" And Not ptObj Is Nothing Then
        On Error Resume Next
        For Each d In ptObj.Decals
            s = d.Name & ""
            If s = "" Then s = d & ""
            If s <> "" Then Exit For
        Next
        Err.Clear
    End If
    PcbDecal = s
End Function

Function FindPartType(doc, name)
    Dim pt
    Set FindPartType = Nothing
    If Trim(name) = "" Then Exit Function
    On Error Resume Next
    For Each pt In doc.PartTypes
        If Err.Number <> 0 Then Exit For
        If StrComp(pt.Name & "", name, vbTextCompare) = 0 Then
            Set FindPartType = pt
            Exit Function
        End If
    Next
    Err.Clear
End Function

Function HasPrefix(refdes, prefixes)
    Dim parts, i, p, u, nxt
    HasPrefix = False
    If Trim(prefixes) = "" Or Trim(refdes) = "" Then Exit Function
    u = UCase(Trim(refdes))
    parts = Split(prefixes, ",")
    For i = 0 To UBound(parts)
        p = UCase(Trim(parts(i)))
        If p <> "" Then
            If Left(u, Len(p)) = p Then
                If Len(u) = Len(p) Then
                    HasPrefix = True
                    Exit Function
                End If
                nxt = Mid(u, Len(p) + 1, 1)
                If nxt >= "0" And nxt <= "9" Then
                    HasPrefix = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Function ContainsAny(text, tokens)
    Dim parts, i, p
    ContainsAny = False
    If Trim(tokens) = "" Or Trim(text) = "" Then Exit Function
    parts = Split(tokens, ",")
    For i = 0 To UBound(parts)
        p = UCase(Trim(parts(i)))
        If p <> "" Then
            If InStr(UCase(text), p) > 0 Then
                ContainsAny = True
                Exit Function
            End If
        End If
    Next
End Function

Function LooksLikeRefdes(text)
    Dim s, i, ch, sawLetter, sawDigit
    LooksLikeRefdes = False
    s = UCase(Trim(text))
    If s = "" Then Exit Function
    sawLetter = False
    sawDigit = False
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch >= "A" And ch <= "Z" Then
            If sawDigit Then Exit Function
            sawLetter = True
        ElseIf ch >= "0" And ch <= "9" Then
            If Not sawLetter Then Exit Function
            sawDigit = True
        Else
            Exit Function
        End If
    Next
    LooksLikeRefdes = sawLetter And sawDigit
End Function

Sub Bump(qtyMap, key)
    If qtyMap.Exists(key) Then
        qtyMap(key) = qtyMap(key) + 1
    Else
        qtyMap.Add key, 1
    End If
End Sub

Function DictGet(row, key)
    DictGet = ""
    If row.Exists(key) Then DictGet = row(key) & ""
End Function

Function DocFolder(doc)
    Dim p, i
    DocFolder = ""
    On Error Resume Next
    p = doc.Path & ""
    If p = "" Then p = doc.FullName & ""
    If p = "" Then
        DocFolder = Environ("TEMP")
        Exit Function
    End If
    If Right(p, 1) = "\" Then
        DocFolder = p
        Exit Function
    End If
    i = InStrRev(p, "\")
    If i > 1 And InStr(p, ".") > 0 Then
        DocFolder = Left(p, i - 1)
    Else
        DocFolder = p
    End If
End Function

Function BaseName(doc)
    Dim n, i
    BaseName = "schematic"
    On Error Resume Next
    n = doc.Name & ""
    If n = "" Then Exit Function
    i = InStrRev(n, "\")
    If i > 0 Then n = Mid(n, i + 1)
    i = InStrRev(n, ".")
    If i > 1 Then n = Left(n, i - 1)
    If n <> "" Then BaseName = n
End Function

Function ResolvePath(folder, schName)
    Dim p
    ResolvePath = ""
    p = Trim(OUTPUT_PATH)
    If p = "" Then
        If folder = "" Then Exit Function
        If Right(folder, 1) <> "\" Then folder = folder & "\"
        ResolvePath = folder & schName & "_BOM_" & Stamp() & ".csv"
        Exit Function
    End If
    If Right(p, 1) = "\" Then
        ResolvePath = p & schName & "_BOM_" & Stamp() & ".csv"
    Else
        ResolvePath = p
    End If
End Function

Function Stamp()
    Stamp = Year(Now) & Two(Month(Now)) & Two(Day(Now)) & "_" & Two(Hour(Now)) & Two(Minute(Now)) & Two(Second(Now))
End Function

Function Two(n)
    If n < 10 Then Two = "0" & CStr(n) Else Two = CStr(n)
End Function

Function WriteText(path, text)
    Dim stm
    WriteText = False
    On Error Resume Next
    Set stm = CreateObject("ADODB.Stream")
    If Err.Number = 0 Then
        stm.Type = 2
        stm.Charset = "utf-8"
        stm.Open
        stm.WriteText text
        stm.SaveToFile path, 2
        stm.Close
        If Err.Number = 0 Then
            WriteText = True
            Exit Function
        End If
    End If
    Err.Clear
    Dim f
    Set f = CreateObject("Scripting.FileSystemObject").CreateTextFile(path, True, False)
    If Err.Number <> 0 Then Exit Function
    f.Write text
    f.Close
    WriteText = (Err.Number = 0)
End Function

Function XlsxPath(csvPath)
    If Len(csvPath) > 4 Then
        If LCase(Right(csvPath, 4)) = ".csv" Then
            XlsxPath = Left(csvPath, Len(csvPath) - 4) & ".xlsx"
            Exit Function
        End If
    End If
    XlsxPath = csvPath & ".xlsx"
End Function

Function SaveExcel(csvPath)
    Dim xl, wb, out
    SaveExcel = False
    On Error Resume Next
    Set xl = CreateObject("Excel.Application")
    If Err.Number <> 0 Then Exit Function
    xl.DisplayAlerts = False
    xl.Visible = False
    Set wb = xl.Workbooks.Open(csvPath)
    If Err.Number <> 0 Then
        xl.Quit
        Exit Function
    End If
    out = XlsxPath(csvPath)
    wb.SaveAs out, 51
    wb.Close False
    xl.Quit
    SaveExcel = (Err.Number = 0)
End Function

Main
