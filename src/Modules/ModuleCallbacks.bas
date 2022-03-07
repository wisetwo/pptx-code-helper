Public PptxCodeHelperRibbon As IRibbonUI

Sub PptxCodeHelperInitialize(Ribbon As IRibbonUI)
    Set PptxCodeHelperRibbon = Ribbon
    PptxCodeHelperRibbon.Invalidate
End Sub
