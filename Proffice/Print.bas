Attribute VB_Name = "Print"
Public Sub PrintReport()
    Printer.FontSize = 10
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print "                                                                                         " & Form1.txtBorja.Text & "      " & Form1.txtSluta.Text & "        " & Form1.txtLunch.Text & "       " & Form1.txtTotalt.Text
    Printer.EndDoc
End Sub
