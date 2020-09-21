Rem Attribute VBA_ModuleType=VBAFormModule
Option VBASupport 1
Option Explicit

Private Sub ButtonOk_Click()
    icol = CInt(TBinputcol.Value)
    ocol = CInt(TBoutputcol.Value)
    Unload MessBox
End Sub
