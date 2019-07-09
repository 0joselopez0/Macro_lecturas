Attribute VB_Name = "Módulo1"

Sub Workbook_Open()

'CONTROL DE APERTURA
'****COMUNICA EL CIERRE DEL RESTO DE INSTANCIAS EXCEL EN EJECUCIÓN, DEBEN CERRARSE PARA UN CORRECTO FUNCIONAMIENTO DE LA MACRO
'El motivo reside en el propio funcionamiento de excel, para poder interactuar con los distintos libros y hojas
'se deben controlar las instancias activas para su asignación como variable ya que no permite realizar ciertas acciones en segundo plano
'_____________________________________________________________________


If Workbooks.Count >= 2 Then
    MsgBox "Para el correcto funcionamiento de la Macro, cierre el resto de libros Excel"
Else
    Dim actual As Workbook: Set actual = ActiveWorkbook 'Seteo del libro actual (macro.xlsm) como variable objeto
    Worksheets.Add
    
    

'APERTURA DE FICHEROS NECESARIOS Y LIMPIADO
'***COMUNICA LA NECESIDAD DE LOS ARCHIVOS NECESARIOS PARA GENERAR LA SALIDA LISTA PARA SU SUBIDA A NETA
'Obligatoriamente se deben seleccionar en el orden comunicado por las distintas ventanas, en caso contrario la macro no podrá realizar su trabajo debidamente
'No existen sentencias de control para verificar el fichero debido a su naturaleza CSV



    MsgBox ("Selecciona el archivo de ruta bajado de neta") 'Cuadro de diálogo
    my_FileNameNeta = Application.GetOpenFilename() 'Cuadro de diálogo
    Dim neta As Workbook: Set neta = Workbooks.Open(Filename:=my_FileNameNeta, Local:=True) 'Seteo del libro seleccionado (Ruta Neta) como variable objeto
    Dim copianeta As Worksheet: Set copianeta = neta.Worksheets(1) 'Copiado del libro (Ruta Neta) al libro actual para operar de manera más eficiente y sencilla
    copianeta.Cells.Copy actual.Worksheets(1).Range("A1")
    Workbooks(2).Close savechanges:=False
    Set actual = ActiveWorkbook
    

'Limpiado y reestructuración de los datos necesarios
'_____________________________________________________________________

    Range("E:E").Clear
    Range("J:AA").Clear
    Range("AC:AC").Clear
    Range("AD:AD").Clear
    Range("AF:AI").Clear
    Range("AO:AU").Clear
    Range("AV:AV").Clear
    Range("AB:AB").Copy
    Range("S:S").PasteSpecial
    Range("AB:AB").Clear
    Range("H:H").Copy
    Range("J:J").PasteSpecial
    Range("I:I").Copy
    Range("K:K").PasteSpecial
    Range("I:I").Clear
    Range("H:H").Clear
    Range("G:G").Copy
    Range("I:I").PasteSpecial
    Range("G:G").Clear
    Range("F:F").Copy
    Range("H:H").PasteSpecial
    Range("F:F").Clear
    Range("B:B").NumberFormat = "0000000000"
    

'Apertura y seteo de variables para el libro de lecturas
'_____________________________________________________________________


    Dim my_FileName As Variant
    MsgBox ("Abre el archivo de lecturas")
    my_FileName = Application.GetOpenFilename()
    Dim brf As Workbook: Set brf = Workbooks.Open(Filename:=my_FileName, Format:=2)
    brf.Sheets(1).Range("1:1").Delete
    

'Declaración de variables simples para el transporte y tratamiento del dato
'_____________________________________________________________________

    
    Dim numerofilas As Integer
    Dim etiqueta As String
    Dim fecha As String
    Dim hora As String
    Dim despiece() As String
    Dim exa() As String
    Dim sujeta() As String
    Dim alarma As String
    Dim caudal As Double
    Dim sinlectura As Workbook
    Set sinlectura = Workbooks.Add
    etiqueta = " "
    Dim dividido As Double
    Dim litros() As String
    Dim m3 As String
        
'Declaración de los libros (listados) de salida
'_____________________________________________________________________
    
    Dim descononcimiento As Variant
    Dim iperl As Workbook
    Set iperl = Workbooks.Add
    iperl.Sheets(1).Range("A1").Value = "Contadores con incidencia iPerl:"
    Dim bajando As Integer
    bajando = 0
    
'BUSQUEDA DE CONTADORES
'***La búsqueda de contadores se realiza mediante la comprobación del campo sigla en el fichero neta frente al campo sigla del fichero de lecturas,
'copiándose despues (de las celdas contigûas) los datos necesarios.
'_____________________________________________________________________
    
    actual.Activate
    numerofilas = actual.Sheets(1).Range("K1", Range("K1").End(xlDown)).Rows.Count
    numerofilastotal = actual.Sheets(1).Range("A1", Range("A1").End(xlDown)).Rows.Count
    actual.Sheets(1).Range("K1").Select
    Dim spliteandofecha() As String
    Dim fechaspliteadaneta As String
    Dim loteparanombre As String
    loteparanombre = actual.Sheets(1).Range("A1").Value
    
'Iterativo de búsqueda y adición
'_____________________________________________________________________

    For x = 1 To numerofilas

        etiqueta = ActiveCell.Value
   
        If Not brf.Sheets(1).Range("B:B").Find(etiqueta) Is Nothing Then
            
            spliteandofecha = Split(brf.Sheets(1).Range("B:B").Find(etiqueta).Offset(0, -1).Value, "/")
            fechaspliteadaneta = spliteandofecha(1) + "/" + spliteandofecha(0) + "/" + spliteandofecha(2)
            dividido = (brf.Sheets(1).Range("B:B").Find(etiqueta).Offset(0, 2).Value) / 1000
            litros = Split(dividido, ",")
            m3 = litros(0)
            ActiveCell.Offset(0, -4).Value = m3
            fecha = Format(fechaspliteadaneta, "yyyy-MM-dd hh:mm:ss,ss")
            hora = Format(brf.Sheets(1).Range("B:B").Find(etiqueta).Offset(0, -1).Value, "Hh:mm")
            exa = Split(brf.Sheets(1).Range("B:B").Find(etiqueta).Offset(0, 3).Value, "-")
               
            'Comprobación de la incidencia iPerl
               
            If exa(0) = "0x02 " Then
        
                iperl.Sheets(1).Range("A2").Offset(bajando, 0).Value = ActiveCell.Offset(0, -3).Value
                 iperl.Sheets(1).Range("B2").Offset(bajando, 0).Value = ActiveCell.Offset(0, -6).Value
                bajando = bajando + 1
                ActiveCell.Offset(0, 3).Value = "INC024"
            
            
        
            Else
                 ActiveCell.Offset(0, 3).Value = "INC001"
        
        
        
        
            End If
        
            ActiveCell.Offset(0, -6).NumberFormat = "@"
            ActiveCell.Offset(0, -6).Value = Replace(fecha, ",", ".")
            ActiveCell.Offset(0, -5).Value = hora
        
        
        End If
        etiqueta = " "
        ActiveCell.Offset(1, 0).Select
    Next

    Range("K:K").Clear
    
    
    
    
    
    actual.Sheets(1).Range("s1").Select



'Adición de la concesión
'_____________________________________________________________________

    Dim splitconcesion() As String
    splitconcesion = Split(actual.Sheets(2).Range("E2").Value, "-")

    For x = 1 To numerofilas

        ActiveCell.Value = splitconcesion(0)
        ActiveCell.Offset(1, 0).Select

    Next
    
'Descubrimiento y escritura de los contadores sin lectura
'_____________________________________________________________________

    Dim fecha2 As String
    Dim hora2 As String

    actual.Sheets(1).Range("G1").Select
    Dim recorrido As Integer
    recorrido = 0


    sinlectura.Sheets(1).Range("A1").Value = "Contadores sin lectura:"
 
    For x = 1 To numerofilas
   
        If ActiveCell.Value = "" Then
    
            ActiveCell.Offset(0, 7).Value = actual.Sheets(2).Range("E5").Value
            fecha2 = Format(actual.Sheets(2).Range("E3").Value, "yyyy-MM-dd hh:mm:s,ss")
            hora2 = Format(actual.Sheets(2).Range("E4").Value, "hh:mm")
            ActiveCell.Offset(0, -2).NumberFormat = "@"
            ActiveCell.Offset(0, -2).Value = Replace(fecha2, ",", ".")
            ActiveCell.Offset(0, -1).Value = hora
            sinlectura.Sheets(1).Range("A2").Offset(recorrido, 0).Value = ActiveCell.Offset(0, 1).Value
            sinlectura.Sheets(1).Range("B2").Offset(recorrido, 0).Value = ActiveCell.Offset(0, -3).Value
            recorrido = recorrido + 1
            ActiveCell.Offset(0, 7).Value = "INC012"
        
        End If
    
        ActiveCell.Offset(1, 0).Select

    Next

'Descubrimiento y escritura de los contadores potencialmente parados
'_____________________________________________________________________

    Dim potparados As Workbook
    Set potparados = Workbooks.Add

    potparados.Sheets(1).Range("A1").Value = "Contadores potencialmente parados:"
    actual.Activate
    actual.Sheets(1).Range("G1").Select
    recorrido = 0

    For x = 1 To numerofilas

        If ActiveCell.Offset(0, 7).Value = "INC001" Then
 
            If Not ActiveCell.Value = "" And ActiveCell.Value < (0.7 * ActiveCell.Offset(0, 33).Value) Then

                potparados.Sheets(1).Range("A2").Offset(recorrido, 0).Value = ActiveCell.Offset(0, 1).Value
                potparados.Sheets(1).Range("B2").Offset(recorrido, 0).Value = ActiveCell.Offset(0, -3).Value
                recorrido = recorrido + 1
                ActiveCell.Offset(0, 7).Value = "INC004"
        
            End If
        End If
        ActiveCell.Offset(1, 0).Select

    Next

'Descubrimiento y escritura de los contadores potencialmente en fuga
'_____________________________________________________________________

    Dim potfuga As Workbook
    Set potfuga = Workbooks.Add

    potfuga.Sheets(1).Range("A1").Value = "Contadores potencialmente fuga interior:"
    actual.Activate
    actual.Sheets(1).Range("G1").Select
    recorrido = 0

    For x = 1 To numerofilas
        If ActiveCell.Offset(0, 7).Value = "INC001" Then

            If (Not ActiveCell.Value = "") And (Not ActiveCell.Offset(0, 24).Value = "") And (Not ActiveCell.Offset(0, 29).Value = "") And IsNumeric(ActiveCell.Offset(0, 24).Value) And IsNumeric(ActiveCell.Offset(0, 29).Value) Then
        
                If ActiveCell.Value > CDbl(ActiveCell.Offset(0, 24)) * 1.3 And ActiveCell.Value > CDbl(ActiveCell.Offset(0, 29)) * 1.3 Then

                    ActiveCell.Offset(0, 7).Value = "INC015"
                    potfuga.Sheets(1).Range("A2").Offset(recorrido, 0).Value = ActiveCell.Offset(0, 1).Value
                     potfuga.Sheets(1).Range("B2").Offset(recorrido, 0).Value = ActiveCell.Offset(0, -3).Value
                    recorrido = recorrido + 1
                End If
            End If
        End If
        ActiveCell.Offset(1, 0).Select

    Next
    
'Eliminación de contadores inexistentes en ruta
'______________________________________________________________

      actual.Sheets(1).Range("N1").Select
    
    For x = 1 To numerofilastotal
        If ActiveCell.Value = "" Then
            Rows(ActiveCell.Row).EntireRow.Delete
        Else
            ActiveCell.Offset(1, 0).Select
        End If
    Next
  
'Limpiado de variables historicas
'_________________________________________________________
    
    actual.Sheets(1).Range("E1").Select
    actual.Sheets(1).Range("G1").Select
    Range("A1").Select
    actual.Sheets(1).Range("AE:AN").Clear

'Asignación de nombre de archivo y guardado
'___________________________________________________

    sinlectura.SaveAs Filename:="Sin_lectura_Lote_" + loteparanombre + ".csv", FileFormat:=xlCSV, Local:=True
    potfuga.SaveAs Filename:="Fuga_interna_Lote_" + loteparanombre + ".csv", FileFormat:=xlCSV, Local:=True
    potparados.SaveAs Filename:="Parados_Lote_" + loteparanombre + ".csv", FileFormat:=xlCSV, Local:=True
    iperl.SaveAs Filename:="Alarmas_iPerl_Lote_" + loteparanombre + ".csv", FileFormat:=xlCSV, Local:=True
    sinlectura.Saved = True
    sinlectura.Close
    potfuga.Saved = True
    potfuga.Close
    potparados.Saved = True
    potparados.Close
    iperl.Saved = True
    iperl.Close
    Workbooks(2).Close savechanges:=False
    Application.DisplayAlerts = False
    actual.Sheets(2).Delete
    Application.DisplayAlerts = True
    actual.SaveAs Filename:="Importacion_Lote_" + loteparanombre + ".csv", FileFormat:=xlCSV, Local:=True
    actual.Saved = True
    actual.Close


    MsgBox ("Archivo de subida y listados creados")
    ActiveWorkbook.Saved = True

End If

End Sub
