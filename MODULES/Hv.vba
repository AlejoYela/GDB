VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Hv 
   Caption         =   "Hoja de VIda"
   ClientHeight    =   11232
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   15024
   OleObjectBlob   =   "Hv.frx":0000
End
Attribute VB_Name = "Hv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BotonGuardar_Click()
    Dim fila As Object
    Dim linea As Integer
    Dim NombreArchivo, RutaArchivo As String
    Dim tech As Object
    
    valor_buscado = Me.CodigoTxtHv
    
    Set fila = Sheets("INVENTARIO GENERAL").Range("A:A").Find(valor_buscado, lookat:=xlWhole)
    
    linea = fila.Row
    
    Sheets("FORMATO HV").Select
        
    NombreArchivo = ActiveSheet.Name & " " & Sheets("INVENTARIO GENERAL").Range("B" & linea) & " " & Sheets("INVENTARIO GENERAL").Range("H" & linea)
    RutaArchivo = ActiveWorkbook.Path & "\HVS\" & NombreArchivo & ".pdf"
    
    If Me.txt_ruta <> "" Then
    
        Set tech = VBA.CreateObject("Scripting.FileSystemObject")
            origen = Me.txt_ruta.Value
            destino = "C:\Users\ALEJANDRO\Desktop\IPS INDÍGENA CUASPUD CARLOSAMA\INVENTARIO\FOTOS EQUIPOS\" & Me.CodigoTxtHv.Value & ".jpg"
            tech.CopyFile origen, destino
            
        Range("H12").Value = Me.CodigoTxtHv.Value & ".jpg"
        Else
            Range(H12).Value = "x.jpg"
    End If
    ru2 = "C:\Users\ALEJANDRO\Desktop\IPS INDÍGENA CUASPUD CARLOSAMA\INVENTARIO\FOTOS EQUIPOS\" & Range("H12").Value
    Image1.Picture = LoadPicture(ru2)
    
    Sheets("FORMATO HV").ExportAsFixedFormat Type:=xlTypePDF, Filename:=RutaArchivo, _
    Quality:=xlQualityStandard, IncludeDocProperties:=True, _
    IgnorePrintAreas:=False, OpenAfterPublish:=False
End Sub

Private Sub BotonSalir1_Click()
    Unload Me
    Inicio.Show
End Sub

Private Sub ListaHv_Click()
    Dim codigo As Integer
    codigo = Me.ListaHv.List(ListaHv.ListIndex, 0)
    Me.CodigoTxtHv.Value = codigo
End Sub

Private Sub Seleccionar_img_Click()
    Set explorar_archivo = Application.FileDialog(msoFileDialogFilePicker)
    explorar_archivo.Title = "Seleccionar foto del equipo"
    explorar_archivo.AllowMultiSelect = False
    explorar_archivo.Show
    
    ruta_imagen = explorar_archivo.SelectedItems(1)
    Me.txt_ruta.Value = ruta_imagen
    Image1.Picture = LoadPicture(ruta_imagen)
End Sub

Private Sub TxtBusquedaHv_Change()
    NumeroDeDatos = Sheets("INVENTARIO GENERAL").Range("B" & Rows.Count).End(xlUp).Row
    ListaHv = Clear
    ListaHv.RowSource = Clear
    
    Y = 0
    
    For fila = 3 To NumeroDeDatos
        nombre = Sheets("INVENTARIO GENERAL").Cells(fila, 2).Value
        If UCase(nombre) Like "*" & UCase(Me.TxtBusquedaHv.Value) & "*" Then
            Me.ListaHv.AddItem
            Me.ListaHv.List(Y, 0) = Sheets("INVENTARIO GENERAL").Cells(fila, 1).Value
            Me.ListaHv.List(Y, 1) = Sheets("INVENTARIO GENERAL").Cells(fila, 2).Value
            Me.ListaHv.List(Y, 2) = Sheets("INVENTARIO GENERAL").Cells(fila, 3).Value
            Me.ListaHv.List(Y, 3) = Sheets("INVENTARIO GENERAL").Cells(fila, 4).Value
            Me.ListaHv.List(Y, 4) = Sheets("INVENTARIO GENERAL").Cells(fila, 5).Value
            Me.ListaHv.List(Y, 5) = Sheets("INVENTARIO GENERAL").Cells(fila, 6).Value
            Me.ListaHv.List(Y, 6) = Sheets("INVENTARIO GENERAL").Cells(fila, 7).Value
            Me.ListaHv.List(Y, 7) = Sheets("INVENTARIO GENERAL").Cells(fila, 8).Value
            Me.ListaHv.List(Y, 8) = Sheets("INVENTARIO GENERAL").Cells(fila, 9).Value
            Me.ListaHv.List(Y, 9) = Sheets("INVENTARIO GENERAL").Cells(fila, 10).Value
            Y = Y + 1
        End If
    Next
End Sub

Private Sub UserForm_Activate()
    Me.ListaHv.RowSource = "INVENTARIO"
    Me.ListaHv.ColumnCount = 12
    
    Sheets("FORMATO HV").Select
    Me.FormaAdquisicion.List = Range("E6:E12").Value
    Me.ClasificacionBiomedica.List = Range("B15:B20").Value
End Sub
