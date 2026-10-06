Sub Main()
    Dim femap As Object
    Dim curView As Object
    Dim rc As Long
    Dim v As Long
    Dim bc As Variant

    ' On Error GoTo ExitHere

    Set femap = GetObject("", "femap.model")

    v = femap.Info_Window(0)

    Set curView = femap.feView
    rc = curView.Get(v)

    bc = curView.WindowBackColor
    curView.WindowTitleBar = False

    rc = curView.Put(v)

    rc = femap.feAppMessage(2, "Active View ID: " & CStr(v))

    ExitHere:
    Set femap = Nothing
    End Sub
