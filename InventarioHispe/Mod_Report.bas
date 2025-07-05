Attribute VB_Name = "ModReport"
Option Explicit
Public sw_Report As Boolean
Public fecha As String
Public Sub ShowReport(rpt As Object, strlink As String, strSQL As String)
   
    rpt.Printer.DeviceName = ""
    
    rpt.Printer.PaperSize = 256
    
    'rpt.Printer.PaperWidth = 1440 * 6.5
    rpt.fldempresa.Text = wnomcia
    rpt.TxtFecha.Text = Format(Date, "dd/mm/yyyy")
    rpt.fldtitulo.Text = "Del  " + Format(fecha, "dd/mm/yyyy")
    rpt.Width = 1440 * 7.2
    
    rpt.Height = 1440 * 3.5
    
    rpt.Tag = strlink
    
    rpt.DataControl1.Source = strSQL
    
    rpt.ToolbarVisible = False
    
    rpt.Show vbModal
End Sub

