VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Mttos 
   Caption         =   "Mantenimientos"
   ClientHeight    =   8760.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   15264
   OleObjectBlob   =   "Mttos.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Mttos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Guardar_Click()

    Dim fila As Object
    Dim linea As Integer
    Dim NombreArchivo, RutaArchivo As String
    
    valor_buscado = Me.CodigoTxt
    
    Set fila = Sheets("INVENTARIO GENERAL").Range("A:A").Find(valor_buscado, lookat:=xlWhole)
    
    linea = fila.Row
    
    
    If Me.Preventivo = True Then
        Sheets("PREVENTIVO").Select
        
        NombreArchivo = ActiveSheet.Name & " " & Sheets("INVENTARIO GENERAL").Range("B" & linea) & " " & Sheets("INVENTARIO GENERAL").Range("H" & linea)
        RutaArchivo = ActiveWorkbook.Path & "\MTTOS DIC\" & NombreArchivo & ".pdf"
        
        Range("B20:B24").Value = ""
        Range("D20:D24").Value = ""
        Range("F20:F24").Value = ""
        
        Range("A18").Value = Me.Diagnostico.Value
        Range("F12").Value = "X"
        Range("D26").Value = Me.Voltaje.Value
        Range("D27").Value = Me.Amperaje.Value
        Range("F26").Value = Me.Presion.Value
        Range("F27").Value = Me.Temperatura.Value
        Range("D29").Value = Me.Horas.Value
        
               
        Range("D12").Value = Sheets("INVENTARIO GENERAL").Range("B" & linea)
        Range("D13").Value = Sheets("INVENTARIO GENERAL").Range("C" & linea)
        Range("D14").Value = Sheets("INVENTARIO GENERAL").Range("D" & linea)
        Range("D15").Value = Sheets("INVENTARIO GENERAL").Range("E" & linea)
        Range("D16").Value = Sheets("INVENTARIO GENERAL").Range("K" & linea)
        Range("F9").Value = Sheets("INVENTARIO GENERAL").Range("H" & linea)
        Range("C10").Value = Sheets("INVENTARIO GENERAL").Range("H" & linea)
        
        Range("F10").Value = Sheets("CRONOGRAMA MTTO").Range("U" & linea)
        Range("C7").Value = Sheets("CRONOGRAMA MTTO").Range("H" & linea)
        Range("F7").Value = Sheets("CRONOGRAMA MTTO").Range("G" & linea)
        
        Range("D20").Value = Sheets("CRONOGRAMA MTTO").Range("J" & linea)
        Range("D21").Value = Sheets("CRONOGRAMA MTTO").Range("K" & linea)
        Range("D22").Value = Sheets("CRONOGRAMA MTTO").Range("L" & linea)
        Range("D23").Value = Sheets("CRONOGRAMA MTTO").Range("M" & linea)
        Range("D24").Value = Sheets("CRONOGRAMA MTTO").Range("N" & linea)
        
        Range("F20").Value = Sheets("CRONOGRAMA MTTO").Range("O" & linea)
        Range("F21").Value = Sheets("CRONOGRAMA MTTO").Range("P" & linea)
        Range("F22").Value = Sheets("CRONOGRAMA MTTO").Range("Q" & linea)
        Range("F23").Value = Sheets("CRONOGRAMA MTTO").Range("R" & linea)
        Range("F24").Value = Sheets("CRONOGRAMA MTTO").Range("S" & linea)
        
        
        If Me.LimpiezaGeneral = True Then
            Range("B20").Value = "X"
        End If
        If Me.Lubricacion = True Then
            Range("B21").Value = "X"
        End If
        If Me.RevisionElectrica = True Then
            Range("B22").Value = "X"
        End If
        If Me.RevisionElectronica = True Then
            Range("B23").Value = "X"
        End If
        If Me.RevisionSensores = True Then
            Range("B24").Value = "X"
        End If
        
        Sheets("PREVENTIVO").ExportAsFixedFormat Type:=xlTypePDF, Filename:=RutaArchivo, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False
        
    ElseIf Me.Correctivo = True Then
        Sheets("CORRECTIVO").Select
        NombreArchivo = ActiveSheet.Name & " " & Sheets("INVENTARIO GENERAL").Range("B" & linea) & " " & Sheets("INVENTARIO GENERAL").Range("H" & linea)
        RutaArchivo = ActiveWorkbook.Path & "\" & NombreArchivo & ".pdf"
        
        Range("B20:B24").Value = ""
        Range("D20:D24").Value = ""
        Range("F20:F24").Value = ""
                
        Range("A18").Value = Me.Diagnostico.Value
        Range("F12").Value = "X"
        Range("D26").Value = Me.Voltaje.Value
        Range("D27").Value = Me.Amperaje.Value
        Range("F26").Value = Me.Presion.Value
        Range("F27").Value = Me.Temperatura.Value
        Range("D29").Value = Me.Horas.Value
        
               
        Range("D12").Value = Sheets("INVENTARIO GENERAL").Range("B" & linea)
        Range("D13").Value = Sheets("INVENTARIO GENERAL").Range("C" & linea)
        Range("D14").Value = Sheets("INVENTARIO GENERAL").Range("D" & linea)
        Range("D15").Value = Sheets("INVENTARIO GENERAL").Range("E" & linea)
        Range("D16").Value = Sheets("INVENTARIO GENERAL").Range("K" & linea)
        Range("F9").Value = Sheets("INVENTARIO GENERAL").Range("H" & linea)
        Range("C10").Value = Sheets("INVENTARIO GENERAL").Range("H" & linea)
        
        Range("F10").Value = Sheets("CRONOGRAMA MTTO").Range("U" & linea)
        Range("C7").Value = Sheets("CRONOGRAMA MTTO").Range("H" & linea)
        Range("F7").Value = Sheets("CRONOGRAMA MTTO").Range("G" & linea)
        
        If Me.LimpiezaGeneral = True Then
            Range("B20").Value = "X"
        End If
        If Me.Lubricacion = True Then
            Range("B21").Value = "X"
        End If
        If Me.RevisionElectrica = True Then
            Range("B22").Value = "X"
        End If
        If Me.RevisionElectronica = True Then
            Range("B23").Value = "X"
        End If
        If Me.RevisionSensores = True Then
            Range("B24").Value = "X"
        End If
        Sheets("CORRECTIVO").ExportAsFixedFormat Type:=xlTypePDF, Filename:=RutaArchivo, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False
    ElseIf Me.Instalacion = True Then
        Sheets("INSTALACIÓN").Select
        NombreArchivo = ActiveSheet.Name & " " & Sheets("INVENTARIO GENERAL").Range("B" & linea) & " " & Sheets("INVENTARIO GENERAL").Range("H" & linea)
        RutaArchivo = ActiveWorkbook.Path & "\" & NombreArchivo & ".pdf"
        
        Range("B20:B24").Value = ""
        Range("D20:D24").Value = ""
        Range("F20:F24").Value = ""
                
        Range("A18").Value = Me.Diagnostico.Value
        Range("F12").Value = "X"
        Range("D26").Value = Me.Voltaje.Value
        Range("D27").Value = Me.Amperaje.Value
        Range("F26").Value = Me.Presion.Value
        Range("F27").Value = Me.Temperatura.Value
        Range("D29").Value = Me.Horas.Value
        
               
        Range("D12").Value = Sheets("INVENTARIO GENERAL").Range("B" & linea)
        Range("D13").Value = Sheets("INVENTARIO GENERAL").Range("C" & linea)
        Range("D14").Value = Sheets("INVENTARIO GENERAL").Range("D" & linea)
        Range("D15").Value = Sheets("INVENTARIO GENERAL").Range("E" & linea)
        Range("D16").Value = Sheets("INVENTARIO GENERAL").Range("K" & linea)
        Range("F9").Value = Sheets("INVENTARIO GENERAL").Range("H" & linea)
        Range("C10").Value = Sheets("INVENTARIO GENERAL").Range("H" & linea)
        
        Range("F10").Value = Sheets("CRONOGRAMA MTTO").Range("U" & linea)
        Range("C7").Value = Sheets("CRONOGRAMA MTTO").Range("H" & linea)
        Range("F7").Value = Sheets("CRONOGRAMA MTTO").Range("G" & linea)
        
        If Me.LimpiezaGeneral = True Then
            Range("B20").Value = "X"
        End If
        If Me.Lubricacion = True Then
            Range("B21").Value = "X"
        End If
        If Me.RevisionElectrica = True Then
            Range("B22").Value = "X"
        End If
        If Me.RevisionElectronica = True Then
            Range("B23").Value = "X"
        End If
        If Me.RevisionSensores = True Then
            Range("B24").Value = "X"
        End If
        Sheets("INSTALACIÓN").ExportAsFixedFormat Type:=xlTypePDF, Filename:=RutaArchivo, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False
    Else
    MsgBox ("Seleccione tipo de Mantenimiento")
    End If
End Sub

Private Sub Lista_Click()
    Dim codigo As Integer
    codigo = Me.Lista.List(Lista.ListIndex, 0)
    Me.CodigoTxt.Value = codigo
End Sub


Private Sub TxtBusqueda_Change()
    NumeroDeDatos = Sheets("INVENTARIO GENERAL").Range("B" & Rows.Count).End(xlUp).Row
    Lista = Clear
    Lista.RowSource = Clear
    
    Y = 0
    
    For fila = 3 To NumeroDeDatos
        nombre = Sheets("INVENTARIO GENERAL").Cells(fila, 2).Value
        If UCase(nombre) Like "*" & UCase(Me.TxtBusqueda.Value) & "*" Then
            Me.Lista.AddItem
            Me.Lista.List(Y, 0) = Sheets("INVENTARIO GENERAL").Cells(fila, 1).Value
            Me.Lista.List(Y, 1) = Sheets("INVENTARIO GENERAL").Cells(fila, 2).Value
            Me.Lista.List(Y, 2) = Sheets("INVENTARIO GENERAL").Cells(fila, 3).Value
            Me.Lista.List(Y, 3) = Sheets("INVENTARIO GENERAL").Cells(fila, 4).Value
            Me.Lista.List(Y, 4) = Sheets("INVENTARIO GENERAL").Cells(fila, 5).Value
            Me.Lista.List(Y, 5) = Sheets("INVENTARIO GENERAL").Cells(fila, 6).Value
            Me.Lista.List(Y, 6) = Sheets("INVENTARIO GENERAL").Cells(fila, 7).Value
            Me.Lista.List(Y, 7) = Sheets("INVENTARIO GENERAL").Cells(fila, 8).Value
            Me.Lista.List(Y, 8) = Sheets("INVENTARIO GENERAL").Cells(fila, 9).Value
            Me.Lista.List(Y, 9) = Sheets("INVENTARIO GENERAL").Cells(fila, 10).Value
            Y = Y + 1
        End If
    Next
End Sub

Private Sub UserForm_Activate()
    Me.Lista.RowSource = "INVENTARIO"
    Me.Lista.ColumnCount = 12
End Sub

Private Sub Volver1_Click()
    Unload Me
    Inicio.Show
End Sub
