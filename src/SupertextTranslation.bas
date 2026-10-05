Attribute VB_Name = "SupertextTranslation"
'
' SupertextTranslation.bas - CorelDRAW VBA macros (XLIFF translation round trip)
' Supertext - CorelDRAW Supertext Translation
' https://github.com/Supertext/CorelDRAW-Supertext-Translation
'
' TranslationExport  exports all text of the active document to XLIFF 1.2.
'                    Every exported text object gets a hidden, persistent ID
'                    (shape property "SupertextID", saved in the .cdr file).
' TranslationImport  reads translated XLIFF files and writes the translations
'                    back into the same text objects, keeping formatting.
'
' - One trans-unit per paragraph; linked text frames are exported once per story.
' - Formatting changes inside a paragraph become <g id="n"> inline tags.
' - Line breaks become <x ctype="lb"/>, tabs <x ctype="x-tab"/>, embedded
'   objects <x ctype="x-coreldraw-object"/> (must stay in place).
' - Each text object carries a <note> with page, layer, font and text type.
'
' The same XLIFF format is used by the Supertext Illustrator and InDesign
' scripts. Windows only (CorelDRAW for Mac has no VBA).
' Install: see docs/installation-guide.md
'
Option Explicit

Private Const VERSION As String = "1.0.0"
Private Const PROP_NAME As String = "SupertextID"
Private Const PROP_INDEX As Long = 1
' "fast" reads formatting runs via EnumRanges. If inline formatting ever goes
' missing after import, switch to "exact" (slower). Export and import share it.
Private Const RUN_DETECTION As String = "fast"
' Character written for a line break the source paragraph didn't have
' (10 = LF, 11 = VT). Breaks already in a paragraph reuse that character.
Private Const DEFAULT_LINE_BREAK As Long = 10

Private g_RangeMode As Long      ' 0 unknown, 1 = 0-based end-exclusive, 2 = 1-based end-inclusive
Private g_Doc As Document
Private g_Fso As Object

' =====================================================================
'  EXPORT
' =====================================================================

Public Sub TranslationExport()
    If Documents.Count = 0 Then MsgBox "Open a document first.", vbExclamation: Exit Sub
    Set g_Doc = ActiveDocument
    Set g_Fso = CreateObject("Scripting.FileSystemObject")
    g_RangeMode = 0

    Dim src As String, tgtText As String, targets As Variant, includeHidden As Boolean
    src = AskLangs("Source language (one code, e.g. de-CH):", "de-CH", True)
    If src = "" Then Exit Sub
    tgtText = AskLangs("Target languages, comma separated (one XLIFF file per language)." & vbCr & _
        "Leave empty for a single file without target language.", "fr-CH, it-CH, en-GB", False)
    If tgtText = vbNullChar Then Exit Sub
    includeHidden = (MsgBox("Include text on hidden layers?", vbYesNo + vbQuestion + vbDefaultButton2, _
        "Export for Translation") = vbYes)

    Dim shapesList As Collection, items As New Collection, used As Object, skipped As Long
    Set used = CreateObject("Scripting.Dictionary")
    Set shapesList = CollectTextShapes(g_Doc, includeHidden)

    Dim entry As Variant, s As Shape, a As Object, id As String, item As Object
    For Each entry In shapesList
        Set s = entry(0)
        If IsFirstFrame(s) Then
            Set a = Analyze(s)
            If Not a Is Nothing Then
                If HasText(a) Then
                    id = WithUnlocked(s, entry(1), "id", used)
                    If id = "" Then
                        skipped = skipped + 1
                    Else
                        Set item = CreateObject("Scripting.Dictionary")
                        item("id") = id
                        Set item("analysis") = a
                        item("order") = entry(2)
                        item("top") = SafeNum(s, "TopY")
                        item("left") = SafeNum(s, "LeftX")
                        item("note") = Describe(s, entry(1), entry(3))
                        items.Add item
                    End If
                End If
            Else
                skipped = skipped + 1
            End If
        End If
    Next

    If items.Count = 0 Then MsgBox "No translatable text found in " & g_Doc.FileName & ".", vbInformation: Exit Sub

    Dim sorted() As Object: sorted = SortItems(items)

    Dim folder As String
    folder = DocFolder(g_Doc)
    If folder = "" Then folder = BrowseFolder("Choose where to save the XLIFF files")
    If folder = "" Then Exit Sub

    Dim base As String, f As String, written As String, units As Long, i As Long
    base = BaseName(IIf(g_Doc.FileName = "", "Untitled", g_Doc.FileName))
    targets = SplitLangs(tgtText)
    If UBound(targets) < 0 Then targets = Array("")
    For i = 0 To UBound(targets)
        f = folder & "\" & base & IIf(targets(i) = "", "", "_" & targets(i)) & ".xlf"
        WriteUtf8 f, BuildXliff(sorted, src, CStr(targets(i)), units)
        written = written & f & vbCr
    Next

    Dim msg As String
    If DocFolder(g_Doc) <> "" Then
        g_Doc.Save
        msg = "The document was saved with the text IDs."
    Else
        msg = "IMPORTANT: save this document. The text IDs needed for the import are stored in it."
    End If
    MsgBox "Exported " & items.Count & " text objects (" & units & " paragraphs) to:" & vbCr & vbCr & _
        written & vbCr & msg & IIf(skipped > 0, vbCr & vbCr & skipped & " text object(s) could not be tagged or read and were skipped.", ""), _
        vbInformation, "Export for Translation"
    Shell "explorer.exe """ & folder & """", vbNormalFocus
End Sub

Private Function BuildXliff(items() As Object, ByVal src As String, ByVal tgt As String, ByRef units As Long) As String
    Dim x As String, i As Long, p As Long, it As Object, paras As Collection, para As Object
    units = 0
    x = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbLf & _
        "<xliff version=""1.2"" xmlns=""urn:oasis:names:tc:xliff:document:1.2"">" & vbLf & _
        "  <file original=""" & Esc(g_Doc.FileName) & """ source-language=""" & Esc(src) & """" & _
        IIf(tgt <> "", " target-language=""" & Esc(tgt) & """", "") & " datatype=""x-coreldraw"">" & vbLf & _
        "    <header>" & vbLf & _
        "      <tool tool-id=""supertext-coreldraw"" tool-name=""Supertext CorelDRAW Translation"" tool-version=""" & VERSION & """/>" & vbLf & _
        "    </header>" & vbLf & "    <body>" & vbLf
    For i = LBound(items) To UBound(items)
        Set it = items(i)
        Set paras = it("analysis")("paras")
        x = x & "      <group id=""" & Esc(it("id")) & """>" & vbLf & _
            "        <note>" & Esc(it("note")) & "</note>" & vbLf
        For p = 1 To paras.Count
            Set para = paras(p)
            If Translatable(para("text")) Then
                x = x & "        <trans-unit id=""" & Esc(it("id")) & "_p" & (p - 1) & """ xml:space=""preserve"">" & vbLf & _
                    "          <source>" & InlineXml(para) & "</source>" & vbLf & _
                    "        </trans-unit>" & vbLf
                units = units + 1
            End If
        Next
        x = x & "      </group>" & vbLf
    Next
    BuildXliff = x & "    </body>" & vbLf & "  </file>" & vbLf & "</xliff>" & vbLf
End Function

' Run 0 is the paragraph's base style; runs with another style become <g id="run index">.
Private Function InlineXml(ByVal para As Object) As String
    Dim out As String, k As Long, r As Variant, baseKey As String, inner As String
    Dim nLb As Long, nTab As Long, nObj As Long, runs As Collection
    Set runs = para("runs")
    If runs.Count > 0 Then r = runs(1): baseKey = r(2)
    For k = 1 To runs.Count
        r = runs(k)
        inner = EscInline(Mid$(para("text"), r(0) - para("start") + 1, r(1)), nLb, nTab, nObj)
        If k = 1 Or r(2) = baseKey Then
            out = out & inner
        Else
            out = out & "<g id=""" & (k - 1) & """>" & inner & "</g>"
        End If
    Next
    InlineXml = out
End Function

Private Function EscInline(ByVal t As String, nLb As Long, nTab As Long, nObj As Long) As String
    Dim s As String, buf As String, i As Long, ch As String
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If IsLineBreak(ch) Then
            s = s & Esc(buf): buf = "": nLb = nLb + 1
            s = s & "<x id=""lb" & nLb & """ ctype=""lb""/>"
        ElseIf ch = vbTab Then
            s = s & Esc(buf): buf = "": nTab = nTab + 1
            s = s & "<x id=""tab" & nTab & """ ctype=""x-tab""/>"
        ElseIf IsHard(ch) Then
            s = s & Esc(buf): buf = "": nObj = nObj + 1
            s = s & "<x id=""o" & nObj & """ ctype=""x-coreldraw-object""/>"
        Else
            buf = buf & ch
        End If
    Next
    EscInline = s & Esc(buf)
End Function

Private Function Esc(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    Esc = Replace(s, """", "&quot;")
End Function

Private Function Describe(ByVal s As Shape, ByVal lyr As Layer, ByVal place As String) As String
    Dim p As String, first As TextRange, n As String
    p = place & " | Layer: " & lyr.Name
    If IsParagraphText(s) Then p = p & " | paragraph text" Else p = p & " | artistic text"
    On Error Resume Next
    Set first = Span(s, 0, 1)
    If Not first Is Nothing Then p = p & " | " & first.Font & " " & Round(first.Size, 1) & " pt"
    n = s.Name
    If n <> "" Then p = p & " | Name: " & n
    If IsParagraphText(s) Then
        p = p & " | fixed text frame, watch the length"
        If Overflows(s) Then p = p & " | already overflowing in the source"
    End If
    If Not lyr.Visible Then p = p & " | hidden"
    On Error GoTo 0
    Describe = p
End Function

Private Function SortItems(ByVal items As Collection) As Object()
    Dim arr() As Object, i As Long, j As Long, tmp As Object
    ReDim arr(1 To items.Count)
    For i = 1 To items.Count: Set arr(i) = items(i): Next
    ' Reading order: page, then top to bottom, then left to right (insertion sort).
    For i = 2 To items.Count
        Set tmp = arr(i)
        j = i - 1
        Do While j >= 1
            If Not ItemBefore(tmp, arr(j)) Then Exit Do
            Set arr(j + 1) = arr(j)
            j = j - 1
        Loop
        Set arr(j + 1) = tmp
    Next
    SortItems = arr
End Function

Private Function ItemBefore(ByVal a As Object, ByVal b As Object) As Boolean
    If a("order") <> b("order") Then ItemBefore = a("order") < b("order"): Exit Function
    If Abs(a("top") - b("top")) > 0.03 Then ItemBefore = a("top") > b("top"): Exit Function
    ItemBefore = a("left") < b("left")
End Function

' =====================================================================
'  IMPORT
' =====================================================================

Public Sub TranslationImport()
    If Documents.Count = 0 Then MsgBox "Open the original (exported) document first.", vbExclamation: Exit Sub
    Set g_Doc = ActiveDocument
    Set g_Fso = CreateObject("Scripting.FileSystemObject")
    g_RangeMode = 0

    Dim files As Collection, packs As New Collection, f As Variant, pack As Object
    Set files = PickXliffFiles()
    If files Is Nothing Then Exit Sub
    For Each f In files
        Set pack = ReadXliff(CStr(f))
        If pack Is Nothing Then Exit Sub
        packs.Add pack
    Next

    Dim summary As String, copies As Boolean, ans As VbMsgBoxResult
    For Each pack In packs
        summary = summary & IIf(pack("lang") = "", "?", pack("lang")) & "  -  " & pack("count") & " paragraphs  -  " & _
            g_Fso.GetFileName(pack("file")) & _
            IIf(pack("original") <> "" And pack("original") <> g_Doc.FileName, "   (exported from " & pack("original") & "!)", "") & vbCr
    Next
    copies = True
    If packs.Count = 1 Then
        ans = MsgBox(summary & vbCr & "Create a translated copy (recommended)?" & vbCr & vbCr & _
            "Yes = new file <name>_<language>.cdr, the original stays untouched" & vbCr & _
            "No = apply to the open document (one undo step)", vbYesNoCancel + vbQuestion, "Import Translation")
        If ans = vbCancel Then Exit Sub
        copies = (ans = vbYes)
    Else
        If MsgBox(summary & vbCr & "A translated copy is created per language. The original stays untouched.", _
            vbOKCancel + vbInformation, "Import Translation") = vbCancel Then Exit Sub
    End If

    Dim out As String, d As Document, lang As String, dest As String, overwrite As Long, n As Long
    If copies Then
        If DocFolder(g_Doc) = "" Then MsgBox "Save the original document first.", vbExclamation: Exit Sub
        If g_Doc.Dirty Then
            If MsgBox("Save changes to " & g_Doc.FileName & " first?", vbOKCancel + vbQuestion) = vbCancel Then Exit Sub
            g_Doc.Save
        End If
        For Each pack In packs
            n = n + 1
            lang = pack("lang"): If lang = "" Then lang = "translated" & n
            dest = DocFolder(g_Doc) & "\" & BaseName(g_Doc.FileName) & "_" & lang & "." & g_Fso.GetExtensionName(g_Doc.FileName)
            If IsOpen(dest) Then
                out = out & lang & ": skipped, close " & g_Fso.GetFileName(dest) & " first." & vbCr & vbCr
            ElseIf g_Fso.FileExists(dest) And overwrite = 2 Then
                out = out & lang & ": skipped (file exists)." & vbCr & vbCr
            Else
                If g_Fso.FileExists(dest) And overwrite = 0 Then
                    overwrite = IIf(MsgBox("Some translated copies already exist. Overwrite them?", vbYesNo + vbQuestion) = vbYes, 1, 2)
                End If
                If g_Fso.FileExists(dest) And overwrite = 2 Then
                    out = out & lang & ": skipped (file exists)." & vbCr & vbCr
                Else
                    g_Fso.CopyFile g_Doc.FullFileName, dest, True
                    Set d = OpenDocument(dest)
                    out = out & Report(lang, d.FileName, ApplyPackUndoable(d, pack)) & vbCr & vbCr
                    d.Save
                End If
            End If
        Next
    Else
        Set pack = packs(1)
        out = Report(pack("lang"), g_Doc.FileName, ApplyPackUndoable(g_Doc, pack))
    End If
    MsgBox "Import finished" & vbCr & vbCr & out, vbInformation, "Import Translation"
End Sub

Private Function ApplyPackUndoable(ByVal d As Document, ByVal pack As Object) As Object
    Dim opt As Boolean
    d.Activate
    opt = Optimization
    d.BeginCommandGroup "Import translation"
    Optimization = True
    Dim e As Long, desc As String
    On Error GoTo failed
    Set ApplyPackUndoable = ApplyPack(d, pack)
    GoTo cleanup
failed:
    e = Err.Number: desc = Err.Description
    Resume cleanup
cleanup:
    On Error Resume Next
    Optimization = opt
    d.EndCommandGroup
    ActiveWindow.Refresh
    On Error GoTo 0
    If e <> 0 Then Err.Raise e, "SupertextTranslation", desc
End Function

Private Function ApplyPack(ByVal d As Document, ByVal pack As Object) As Object
    Dim rep As Object, index As Object, shapesList As Collection, entry As Variant, id As String
    Set rep = NewReport()
    Set index = CreateObject("Scripting.Dictionary")
    Set shapesList = CollectTextShapes(d, True)
    For Each entry In shapesList
        If IsFirstFrame(entry(0)) Then
            id = GetId(entry(0))
            If id <> "" Then
                If index.Exists(id) Then
                    rep("dupes") = rep("dupes") + 1
                Else
                    index.Add id, entry
                End If
            End If
        End If
    Next

    Dim gid As Variant, res As String, s As Shape
    For Each gid In pack("units").Keys
        If Not index.Exists(gid) Then
            rep("missing") = rep("missing") + 1
        Else
            entry = index(gid)
            Set s = entry(0)
            res = TryApply(s, entry(1), pack("units")(gid), rep)
            If Left$(res, 6) = "error:" Then
                rep("errors").Add Snippet(StoryText(s)) & " " & Mid$(res, 8)
            ElseIf res = "changed" Then
                rep("changed").Add Snippet(StoryText(s))
            ElseIf res = "ok" Then
                If Overflows(s) Then rep("overflow").Add entry(3) & " " & Snippet(StoryText(s))
            End If
        End If
    Next
    Set ApplyPack = rep
End Function

Private Function TryApply(ByVal s As Shape, ByVal lyr As Layer, ByVal unitMap As Object, ByVal rep As Object) As String
    On Error GoTo failed
    TryApply = WithUnlocked(s, lyr, "apply", unitMap, rep)
    Exit Function
failed:
    TryApply = "error: " & Err.Description
End Function

Private Function ApplyShape(ByVal s As Shape, ByVal unitMap As Object, ByVal rep As Object) As String
    Dim a As Object, paras As Collection, key As Variant, p As Long, u As Variant, anyDone As Boolean
    Set a = Analyze(s)
    If a Is Nothing Then ApplyShape = "changed": Exit Function
    Set paras = a("paras")

    ' Safety: the source text must still match what was exported.
    Dim para As Object
    For Each key In unitMap.Keys
        p = CLng(key) + 1
        If p > paras.Count Then ApplyShape = "changed": Exit Function
        u = unitMap(key)
        Set para = paras(p)
        If PlainSegs(u(0)) <> Canon(para("text")) Then ApplyShape = "changed": Exit Function
    Next

    ' Last paragraph first, so earlier character positions stay valid.
    For p = paras.Count To 1 Step -1
        If unitMap.Exists(CStr(p - 1)) Then
            u = unitMap(CStr(p - 1))
            Set para = paras(p)
            If u(1) Is Nothing Then
                rep("untranslated") = rep("untranslated") + 1
            ElseIf ApplyParagraph(s, para, u(1)) Then
                rep("translated") = rep("translated") + 1
                anyDone = True
            Else
                rep("tags").Add Snippet(para("text"))
            End If
        End If
    Next
    ApplyShape = IIf(anyDone, "ok", "unchanged")
End Function

' Replaces the text between the object characters of one paragraph. Each
' chunk's translation is inserted before its old text and gets the formatting
' of the matching original style run (CopyAttributes); then the old text is
' deleted. Returns False (and changes nothing) if the placeholders don't match.
Private Function ApplyParagraph(ByVal s As Shape, ByVal para As Object, ByVal segs As Collection) As Boolean
    Dim txt As String, pStart As Long, i As Long, k As Long, n As Long
    txt = para("text"): pStart = para("start")

    Dim hard As New Collection
    For i = 1 To Len(txt)
        If IsHard(Mid$(txt, i, 1)) Then hard.Add pStart + i - 1
    Next

    ' Split the target at its placeholders, which must be o1..oN in order.
    Dim chunks As New Collection, cur As Collection, seg As Variant
    Set cur = New Collection: chunks.Add cur
    For Each seg In segs
        If seg(0) Then
            n = n + 1
            If seg(3) <> "o" & n Then Exit Function
            Set cur = New Collection: chunks.Add cur
        ElseIf Len(seg(1)) > 0 Then
            cur.Add seg
        End If
    Next
    If n <> hard.Count Then Exit Function

    ' Line breaks: reuse the character the paragraph already has.
    Dim lb As String
    lb = ChrW(DEFAULT_LINE_BREAK)
    If InStr(txt, vbLf) > 0 Then lb = vbLf Else If InStr(txt, Chr(11)) > 0 Then lb = Chr(11)

    ' loc(k): current position of a character that carries style run k's formatting.
    Dim runs As Collection, loc() As Long, rn As Variant
    Set runs = para("runs")
    If runs.Count = 0 Then ReDim loc(0) Else ReDim loc(runs.Count - 1)
    For k = 1 To runs.Count
        rn = runs(k)
        loc(k - 1) = rn(0)
    Next
    If runs.Count = 0 Then loc(0) = -1

    Dim bounds() As Long
    ReDim bounds(0 To 2 * hard.Count + 1)
    bounds(0) = pStart
    For i = 1 To hard.Count
        bounds(2 * i - 1) = hard(i)
        bounds(2 * i) = hard(i) + 1
    Next
    bounds(2 * hard.Count + 1) = pStart + Len(txt)

    ' Pass 1, last chunk first: insert and format the new text while every
    ' original character still exists as a formatting source.
    Dim c As Long, cs As Long, oldLen As Long, newText As String, L As Long, off As Long, src As Long
    Dim segText As String, st As Long, ins() As Long, shift As Long
    ReDim ins(1 To chunks.Count)
    For c = chunks.Count To 1 Step -1
        cs = bounds(2 * (c - 1))
        newText = ""
        For Each seg In chunks(c): newText = newText & Replace(seg(1), vbLf, lb): Next
        L = Len(newText)
        ins(c) = L
        shift = shift + L
        If L > 0 Then
            InsertAt s, cs, newText
            For k = 0 To UBound(loc)
                If loc(k) >= cs Then loc(k) = loc(k) + L
            Next
            off = 0
            For Each seg In chunks(c)
                segText = Replace(seg(1), vbLf, lb)
                st = seg(2): If st < 0 Or st > UBound(loc) Then st = 0
                src = loc(st)
                If src < 0 Then src = FirstValid(loc)
                If src >= 0 Then Span(s, cs + off, Len(segText)).CopyAttributes Span(s, src, 1)
                off = off + Len(segText)
            Next
        End If
    Next

    ' Pass 2, last chunk first: delete the old text. The old text of chunk c
    ' now starts after the new text of chunks 1..c.
    For c = chunks.Count To 1 Step -1
        cs = bounds(2 * (c - 1)): oldLen = bounds(2 * (c - 1) + 1) - cs
        If oldLen > 0 Then Span(s, cs + shift, oldLen).Delete
        shift = shift - ins(c)
    Next
    ApplyParagraph = True
End Function

Private Function FirstValid(loc() As Long) As Long
    Dim k As Long
    FirstValid = -1
    For k = 0 To UBound(loc)
        If loc(k) >= 0 Then FirstValid = loc(k): Exit Function
    Next
End Function

Private Sub InsertAt(ByVal s As Shape, ByVal pos As Long, ByVal text As String)
    If pos < s.Text.Story.Length Then
        Span(s, pos, 1).InsertBeforeWide text
    Else
        Span(s, pos - 1, 1).InsertAfterWide text
    End If
End Sub

Private Function Overflows(ByVal s As Shape) As Boolean
    On Error Resume Next
    If IsParagraphText(s) Then Overflows = s.Text.Overflow
End Function

' ---------------------------------------------------------------- report

Private Function NewReport() As Object
    Dim r As Object
    Set r = CreateObject("Scripting.Dictionary")
    r("translated") = 0: r("untranslated") = 0: r("missing") = 0: r("dupes") = 0
    Set r("changed") = New Collection: Set r("tags") = New Collection: Set r("overflow") = New Collection
    Set r("errors") = New Collection
    Set NewReport = r
End Function

Private Function Report(ByVal lang As String, ByVal name As String, ByVal r As Object) As String
    Dim l As String
    l = IIf(lang <> "", lang & " -> ", "") & name & vbCr & "  Translated paragraphs: " & r("translated")
    If r("untranslated") > 0 Then l = l & vbCr & "  Not translated (original kept): " & r("untranslated")
    If r("tags").Count > 0 Then l = l & vbCr & "  Kept original, object placeholders missing or moved: " & r("tags").Count & " (" & Join5(r("tags")) & ")"
    If r("changed").Count > 0 Then l = l & vbCr & "  Skipped, text changed since export: " & Join5(r("changed"))
    If r("missing") > 0 Then l = l & vbCr & "  Text objects not found (deleted?): " & r("missing")
    If r("dupes") > 0 Then l = l & vbCr & "  Text objects with duplicate IDs (copied after export): " & r("dupes")
    If r("overflow").Count > 0 Then l = l & vbCr & "  OVERFLOWING TEXT in " & r("overflow").Count & " frame(s): " & Join5(r("overflow"))
    If r("errors").Count > 0 Then l = l & vbCr & "  ERRORS, original kept: " & r("errors").Count & " (" & Join5(r("errors")) & ")"
    Report = l
End Function

Private Function Join5(ByVal c As Collection) As String
    Dim i As Long, s As String
    For i = 1 To c.Count
        If i > 5 Then s = s & "; ...": Exit For
        s = s & IIf(i > 1, "; ", "") & c(i)
    Next
    Join5 = s
End Function

Private Function Snippet(ByVal t As String) As String
    Dim i As Long, ch As String, s As String
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If AscW(ch) >= 0 And AscW(ch) < 32 Then s = s & " " Else s = s & ch
    Next
    s = Trim$(s)
    If Len(s) > 30 Then s = Left$(s, 30) & "..."
    Snippet = """" & s & """"
End Function

' =====================================================================
'  XLIFF READING
' =====================================================================

Private Function PickXliffFiles() As Collection
    Dim path As String, folder As String, base As String, f As String, v As Variant, c As New Collection
    On Error Resume Next
    path = CorelScriptTools.GetFileBox("XLIFF files (*.xlf;*.xliff)|*.xlf;*.xliff|All files (*.*)|*.*", _
        "Select a translated XLIFF file", 0)
    If Err.Number <> 0 Then
        Err.Clear
        path = InputBox("Full path of the translated XLIFF file:", "Import Translation")
    End If
    On Error GoTo 0
    If path = "" Then Exit Function
    If Not g_Fso.FileExists(path) Then MsgBox "File not found: " & path, vbExclamation: Exit Function
    c.Add path

    ' Offer the other languages exported from the same document.
    folder = g_Fso.GetParentFolderName(path)
    base = BaseName(g_Doc.FileName)
    Dim others As New Collection
    f = Dir(folder & "\" & base & "_*.xlf")
    Do While f <> ""
        If LCase$(folder & "\" & f) <> LCase$(path) Then others.Add folder & "\" & f
        f = Dir()
    Loop
    If others.Count > 0 Then
        If MsgBox(others.Count & " more XLIFF file(s) for " & g_Doc.FileName & " found in the same folder." & vbCr & _
            "Import them too?", vbYesNo + vbQuestion, "Import Translation") = vbYes Then
            For Each v In others: c.Add v: Next
        End If
    End If
    Set PickXliffFiles = c
End Function

Private Function ReadXliff(ByVal path As String) As Object
    Dim x As Object, pack As Object, fileEl As Object, m As Object
    Set x = CreateObject("MSXML2.DOMDocument.6.0")
    x.async = False
    x.preserveWhiteSpace = True
    x.validateOnParse = False
    x.resolveExternals = False
    If Not x.Load(path) Then
        MsgBox "Could not read " & g_Fso.GetFileName(path) & ":" & vbCr & x.parseError.reason, vbExclamation
        Exit Function
    End If
    Set pack = CreateObject("Scripting.Dictionary")
    pack("file") = path: pack("lang") = "": pack("original") = "": pack("count") = 0
    Set pack("units") = CreateObject("Scripting.Dictionary")
    WalkXliff x.documentElement, pack
    If pack("lang") = "" Then
        Set m = CreateObject("VBScript.RegExp")
        m.Pattern = "_([A-Za-z]{2,3}(-[A-Za-z0-9]+)*)\.(xlf|xliff|xml)$"
        m.IgnoreCase = True
        If m.Test(g_Fso.GetFileName(path)) Then pack("lang") = m.Execute(g_Fso.GetFileName(path))(0).SubMatches(0)
    End If
    If pack("count") = 0 Then MsgBox "No translation units found in " & g_Fso.GetFileName(path) & ".", vbExclamation: Exit Function
    Set ReadXliff = pack
End Function

Private Sub WalkXliff(ByVal node As Object, ByVal pack As Object)
    Dim ch As Object, id As String, pos As Long, gid As String, pIdx As String
    Dim srcSegs As Collection, tgtSegs As Collection, k As Object, unitMap As Object
    For Each ch In node.childNodes
        If ch.nodeType = 1 Then
            Select Case ch.baseName
            Case "file"
                If pack("lang") = "" Then pack("lang") = AttrOf(ch, "target-language")
                If pack("original") = "" Then pack("original") = AttrOf(ch, "original")
                WalkXliff ch, pack
            Case "trans-unit"
                id = AttrOf(ch, "id")
                pos = InStrRev(id, "_p")
                If pos > 1 Then
                    gid = Left$(id, pos - 1): pIdx = Mid$(id, pos + 2)
                    If IsNumeric(pIdx) Then
                        Set srcSegs = New Collection: Set tgtSegs = Nothing
                        For Each k In ch.childNodes
                            If k.nodeType = 1 Then
                                If k.baseName = "source" Then Set srcSegs = InlineSegs(k, 0, New Collection)
                                If k.baseName = "target" Then Set tgtSegs = InlineSegs(k, 0, New Collection)
                            End If
                        Next
                        If Not tgtSegs Is Nothing Then
                            If Not Translatable(PlainSegs(tgtSegs)) Then Set tgtSegs = Nothing
                        End If
                        If Not pack("units").Exists(gid) Then pack("units").Add gid, CreateObject("Scripting.Dictionary")
                        Set unitMap = pack("units")(gid)
                        unitMap(CStr(CLng(pIdx))) = Array(srcSegs, tgtSegs)
                        pack("count") = pack("count") + 1
                    End If
                End If
            Case Else
                WalkXliff ch, pack
            End Select
        End If
    Next
End Sub

Private Function AttrOf(ByVal el As Object, ByVal name As String) As String
    Dim v As Variant
    v = el.getAttribute(name)
    If Not IsNull(v) Then AttrOf = CStr(v)
End Function

' Segment = Array(isHard, text, style, hardId). style 0 = paragraph base, n = run n.
Private Function InlineSegs(ByVal node As Object, ByVal style As Long, ByVal segs As Collection) As Collection
    Dim ch As Object, t As String, ct As String, gid As String
    For Each ch In node.childNodes
        Select Case ch.nodeType
        Case 3, 4
            t = Replace(Replace(Replace(ch.nodeValue, vbCrLf, vbLf), vbCr, vbLf), Chr(11), vbLf)
            segs.Add Array(False, t, style, "")
        Case 1
            Select Case ch.baseName
            Case "g"
                gid = DigitsOf(AttrOf(ch, "id"))
                If gid = "" Then InlineSegs ch, style, segs Else InlineSegs ch, CLng(gid), segs
            Case "x"
                ct = AttrOf(ch, "ctype")
                If ct = "lb" Then
                    segs.Add Array(False, vbLf, style, "")
                ElseIf ct = "x-tab" Then
                    segs.Add Array(False, vbTab, style, "")
                ElseIf ct = "x-coreldraw-object" Then
                    segs.Add Array(True, "", style, AttrOf(ch, "id"))
                End If
            Case "mrk", "sub"
                InlineSegs ch, style, segs
            End Select
        End Select
    Next
    Set InlineSegs = segs
End Function

Private Function DigitsOf(ByVal s As String) As String
    Dim i As Long
    For i = 1 To Len(s)
        If Mid$(s, i, 1) Like "#" Then DigitsOf = DigitsOf & Mid$(s, i, 1)
    Next
    If Len(DigitsOf) > 9 Then DigitsOf = ""
End Function

' Plain text with line breaks as LF and object placeholders as U+E000.
Private Function PlainSegs(ByVal segs As Collection) As String
    Dim seg As Variant, s As String
    For Each seg In segs
        If seg(0) Then s = s & ChrW(&HE000) Else s = s & seg(1)
    Next
    PlainSegs = s
End Function

Private Function Canon(ByVal text As String) As String
    Dim i As Long, ch As String, s As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If IsLineBreak(ch) Then
            s = s & vbLf
        ElseIf IsHard(ch) Then
            s = s & ChrW(&HE000)
        Else
            s = s & ch
        End If
    Next
    Canon = s
End Function

' =====================================================================
'  SHARED: shapes, IDs, text analysis
' =====================================================================

' Returns a Collection of Array(shape, layer, sortOrder, placeLabel)
' for every text shape in the document (pages, then master page/desktop).
Private Function CollectTextShapes(ByVal d As Document, ByVal includeHidden As Boolean) As Collection
    Dim c As New Collection, seen As Object, p As Long, lyr As Layer
    Set seen = CreateObject("Scripting.Dictionary")
    For p = 1 To d.Pages.Count
        For Each lyr In d.Pages(p).Layers
            If lyr.Visible Or includeHidden Then AddShapes lyr.Shapes, lyr, p, "Page " & p, c, seen
        Next
    Next
    For Each lyr In d.MasterPage.Layers
        If lyr.Visible Or includeHidden Then
            AddShapes lyr.Shapes, lyr, 100000, IIf(LayerIsDesktop(lyr), "Desktop (pasteboard)", "Master layer, all pages"), c, seen
        End If
    Next
    Set CollectTextShapes = c
End Function

Private Sub AddShapes(ByVal shs As Shapes, ByVal lyr As Layer, ByVal order As Long, ByVal place As String, ByVal c As Collection, ByVal seen As Object)
    Dim s As Shape, key As String
    For Each s In shs
        key = CStr(s.StaticID)
        If Not seen.Exists(key) Then
            seen.Add key, True
            If s.Type = cdrTextShape Then
                c.Add Array(s, lyr, order, place)
            ElseIf s.Type = cdrGroupShape Then
                AddShapes s.Shapes, lyr, order, place, c, seen
            End If
            On Error Resume Next
            Dim pc As PowerClip
            Set pc = Nothing
            Set pc = s.PowerClip
            On Error GoTo 0
            If Not pc Is Nothing Then AddShapes pc.Shapes, lyr, order, place, c, seen
        End If
    Next
End Sub

Private Function LayerIsDesktop(ByVal lyr As Layer) As Boolean
    On Error Resume Next
    LayerIsDesktop = lyr.IsDesktopLayer
End Function

Private Function IsParagraphText(ByVal s As Shape) As Boolean
    On Error Resume Next
    IsParagraphText = (s.Text.Type = cdrParagraphText)
End Function

' Linked paragraph frames share one story: only the first frame is processed.
Private Function IsFirstFrame(ByVal s As Shape) As Boolean
    IsFirstFrame = True
    On Error GoTo done
    If IsParagraphText(s) Then IsFirstFrame = (s.Text.Frame.Range.Start <= s.Text.Story.Start)
done:
End Function

Private Function GetId(ByVal s As Shape) As String
    Dim v As Variant
    On Error Resume Next
    v = s.Properties(PROP_NAME, PROP_INDEX)
    If Err.Number = 0 And Not IsEmpty(v) And Not IsNull(v) Then GetId = CStr(v)
End Function

Private Function EnsureId(ByVal s As Shape, ByVal used As Object) As String
    Dim id As String
    id = GetId(s)
    ' Keep an existing ID unless another object already has it (copy/paste).
    If id <> "" And Not used.Exists(id) Then used.Add id, True: EnsureId = id: Exit Function
    Do
        Randomize
        id = "cd" & LCase$(Hex$(CLng(Timer * 100))) & LCase$(Hex$(CLng(Rnd * 2147483647)))
    Loop While used.Exists(id)
    On Error GoTo fail
    s.Properties(PROP_NAME, PROP_INDEX) = id
    used.Add id, True
    EnsureId = id
    Exit Function
fail:
    EnsureId = ""
End Function

' Temporarily makes the layer editable and the shape unlocked, runs the
' action ("id" or "apply"), then restores the state.
Private Function WithUnlocked(ByVal s As Shape, ByVal lyr As Layer, ByVal action As String, ByVal arg1 As Object, Optional ByVal arg2 As Object) As String
    Dim wasLocked As Boolean, wasReadOnly As Boolean, result As String, e As Long, desc As String
    On Error Resume Next
    wasReadOnly = Not lyr.Editable
    If wasReadOnly Then lyr.Editable = True
    wasLocked = s.Locked
    If wasLocked Then s.Locked = False
    Err.Clear
    On Error GoTo failed
    If action = "id" Then result = EnsureId(s, arg1) Else result = ApplyShape(s, arg1, arg2)
    GoTo cleanup
failed:
    e = Err.Number: desc = Err.Description
    Resume cleanup
cleanup:
    On Error Resume Next
    If wasLocked Then s.Locked = True
    If wasReadOnly Then lyr.Editable = False
    On Error GoTo 0
    If e <> 0 Then Err.Raise e, "SupertextTranslation", desc
    WithUnlocked = result
End Function

' Story text with one character per text position, or vbNullChar if the
' text and the character positions can't be matched.
Private Function StoryText(ByVal s As Shape) As String
    Dim t As String, n As Long
    t = s.Text.Story.WideText
    n = s.Text.Story.Length
    If Len(t) <> n Then t = Replace(t, vbCrLf, vbCr)
    If Len(t) <> n Then t = vbNullChar
    StoryText = t
End Function

' Returns a Dictionary: "text", "paras" = Collection of Dictionary
' ("start" 0-based position, "text", "runs" = Collection of Array(start, len, key)).
Private Function Analyze(ByVal s As Shape) As Object
    Dim txt As String, runs As Collection, parts() As String, paras As New Collection
    Dim p As Long, pos As Long, para As Object, k As Long, r As Variant, a As Long, b As Long, e As Long
    txt = StoryText(s)
    If txt = vbNullChar Then Exit Function
    Set runs = GetRuns(s, Len(txt))
    parts = Split(txt, vbCr)
    pos = 0
    For p = 0 To UBound(parts)
        e = pos + Len(parts(p))
        Set para = CreateObject("Scripting.Dictionary")
        para("start") = pos: para("text") = parts(p)
        Set para("runs") = New Collection
        For k = 1 To runs.Count
            r = runs(k)
            a = IIf(r(0) > pos, r(0), pos)
            b = IIf(r(0) + r(1) < e, r(0) + r(1), e)
            If b > a Then PushRun para("runs"), a, b - a, CStr(r(2))
        Next
        paras.Add para
        pos = e + 1
    Next
    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    res("text") = txt
    Set res("paras") = paras
    Set Analyze = res
End Function

Private Function GetRuns(ByVal s As Shape, ByVal n As Long) As Collection
    Dim runs As Collection, i As Long
    If RUN_DETECTION = "fast" Then
        Set runs = FastRuns(s, n)
        If Not runs Is Nothing Then Set GetRuns = runs: Exit Function
    End If
    Set runs = New Collection
    For i = 0 To n - 1
        PushRun runs, i, 1, StyleKey(Span(s, i, 1))
    Next
    Set GetRuns = runs
End Function

' Runs from EnumRanges (one call per formatting change), or Nothing if their
' lengths don't add up to the story length.
Private Function FastRuns(ByVal s As Shape, ByVal n As Long) As Collection
    Dim runs As New Collection, rs As Object, r As TextRange, i As Long, pos As Long, l As Long
    On Error GoTo fail
    Set rs = s.Text.Story.EnumRanges()
    For i = 1 To rs.Count
        Set r = rs(i)
        l = r.Length
        If l > 0 Then
            PushRun runs, pos, l, StyleKey(r)
            pos = pos + l
        End If
    Next
    If pos = n Then Set FastRuns = runs
fail:
End Function

Private Sub PushRun(ByVal runs As Collection, ByVal start As Long, ByVal length As Long, ByVal key As String)
    Dim last As Variant
    If runs.Count > 0 Then
        last = runs(runs.Count)
        If last(2) = key And last(0) + last(1) = start Then
            runs.Remove runs.Count
            runs.Add Array(last(0), last(1) + length, key)
            Exit Sub
        End If
    End If
    runs.Add Array(start, length, key)
End Sub

Private Function StyleKey(ByVal r As TextRange) As String
    StyleKey = Prop(r, "Font") & "|" & Prop(r, "Size") & "|" & Prop(r, "Bold") & "|" & Prop(r, "Italic") & "|" & _
        Prop(r, "Underline") & "|" & Prop(r, "Strikethru") & "|" & Prop(r, "Position") & "|" & Prop(r, "Case") & "|" & _
        Prop(r, "CharSpacing") & "|" & FillKey(r)
End Function

Private Function FillKey(ByVal r As TextRange) As String
    On Error Resume Next
    FillKey = CStr(r.Fill.Type)
    FillKey = FillKey & ":" & r.Fill.UniformColor.HexValue
End Function

Private Function Prop(ByVal o As Object, ByVal name As String) As String
    On Error Resume Next
    Prop = CStr(CallByName(o, name, VbGet))
End Function

' The TextRange covering [pos, pos + length) of the story (0-based positions).
' CorelDRAW's documentation doesn't state whether TextRange.Range is 0-based
' and end-exclusive, so the convention is detected once per run.
Private Function Span(ByVal s As Shape, ByVal pos As Long, ByVal length As Long) As TextRange
    Dim st As TextRange
    Set st = s.Text.Story
    If g_RangeMode = 0 Then CalibrateRange st
    If g_RangeMode = 1 Then
        Set Span = st.Range(pos, pos + length)
    Else
        Set Span = st.Range(pos + 1, pos + length)
    End If
End Function

Private Sub CalibrateRange(ByVal st As TextRange)
    Dim t As String
    t = st.WideText
    g_RangeMode = 1
    If Len(t) >= 2 Then
        If RangeText(st, 0, 1) = Left$(t, 1) And RangeText(st, 1, 2) = Mid$(t, 2, 1) Then Exit Sub
        If RangeText(st, 1, 1) = Left$(t, 1) And RangeText(st, 2, 2) = Mid$(t, 2, 1) Then g_RangeMode = 2
    End If
End Sub

Private Function RangeText(ByVal st As TextRange, ByVal a As Long, ByVal b As Long) As String
    RangeText = vbNullChar
    On Error Resume Next
    RangeText = st.Range(a, b).WideText
End Function

Private Function IsLineBreak(ByVal ch As String) As Boolean
    IsLineBreak = (ch = vbLf Or ch = Chr(11))
End Function

' Characters that stand for embedded objects or fields. They must survive the import.
Private Function IsHard(ByVal ch As String) As Boolean
    Dim c As Long
    c = AscW(ch) And &HFFFF&
    IsHard = (c < 32 And c <> 9 And c <> 10 And c <> 11 And c <> 13) Or c = &HFFFC& Or c = &HFEFF&
End Function

Private Function Translatable(ByVal t As String) As Boolean
    Dim i As Long, c As Long
    For i = 1 To Len(t)
        c = AscW(Mid$(t, i, 1)) And &HFFFF&
        If c > 32 And c <> &HA0& And c <> &HFFFC& And c <> &HFEFF& And c <> &HE000& And c <> &H2028& And c <> &H2029& Then
            Translatable = True: Exit Function
        End If
    Next
End Function

Private Function HasText(ByVal a As Object) As Boolean
    Dim para As Variant
    For Each para In a("paras")
        If Translatable(para("text")) Then HasText = True: Exit Function
    Next
End Function

' =====================================================================
'  HELPERS
' =====================================================================

Private Function AskLangs(ByVal prompt As String, ByVal def As String, ByVal singleOnly As Boolean) As String
    Dim v As String, langs As Variant
    Do
        v = InputBox(prompt, "Export for Translation", def)
        If StrPtr(v) = 0 Then AskLangs = IIf(singleOnly, "", vbNullChar): Exit Function
        langs = SplitLangs(v)
        If IsNull(langs) Then
            MsgBox "Use language codes like de-CH, fr-CH, en-GB.", vbExclamation
        ElseIf singleOnly And UBound(langs) <> 0 Then
            MsgBox "Enter exactly one source language code, e.g. de-CH.", vbExclamation
        Else
            If singleOnly Then AskLangs = langs(0) Else AskLangs = v
            Exit Function
        End If
        def = v
    Loop
End Function

' Returns an array of language codes, or Null if one is invalid.
Private Function SplitLangs(ByVal v As String) As Variant
    Dim re As Object, parts() As String, out() As String, i As Long, n As Long
    v = Replace(Replace(v, ";", ","), " ", ",")
    parts = Split(v, ",")
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*$"
    ReDim out(0 To UBound(parts) + 1)
    n = 0
    For i = 0 To UBound(parts)
        If parts(i) <> "" Then
            If Not re.Test(parts(i)) Then SplitLangs = Null: Exit Function
            out(n) = parts(i): n = n + 1
        End If
    Next
    If n = 0 Then SplitLangs = Array(): Exit Function
    ReDim Preserve out(0 To n - 1)
    SplitLangs = out
End Function

Private Function DocFolder(ByVal d As Document) As String
    On Error Resume Next
    If d.FullFileName <> "" Then
        If g_Fso.FileExists(d.FullFileName) Then DocFolder = g_Fso.GetParentFolderName(d.FullFileName)
    End If
End Function

Private Function BaseName(ByVal name As String) As String
    Dim p As Long
    p = InStrRev(name, ".")
    If p > 1 Then BaseName = Left$(name, p - 1) Else BaseName = name
End Function

Private Function BrowseFolder(ByVal title As String) As String
    Dim f As Object
    On Error Resume Next
    Set f = CreateObject("Shell.Application").BrowseForFolder(0, title, 0)
    If Not f Is Nothing Then BrowseFolder = f.Self.Path
End Function

Private Function IsOpen(ByVal path As String) As Boolean
    Dim d As Document
    For Each d In Documents
        If LCase$(d.FullFileName) = LCase$(path) Then IsOpen = True: Exit Function
    Next
End Function

Private Sub WriteUtf8(ByVal path As String, ByVal text As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText text
    st.SaveToFile path, 2
    st.Close
End Sub

Private Function SafeNum(ByVal o As Object, ByVal name As String) As Double
    On Error Resume Next
    SafeNum = CDbl(CallByName(o, name, VbGet))
End Function
