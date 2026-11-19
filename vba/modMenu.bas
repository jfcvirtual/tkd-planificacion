Attribute VB_Name = "modMenu"
'======================================================================
'  PLANIFICACIÓN DEL ENTRENAMIENTO DE TAEKWONDO  (versión 2)
'  Ciclo ANUAL (48 semanas) o SEMESTRAL (24 semanas)
'  Volúmenes e intensidades según las curvas de Matveyev,
'  microciclos semanales e índice de sobreentrenamiento.
'
'  Plantilla original: Justo Fuentes
'  Los CÁLCULOS los hacen las fórmulas del libro (hojas DATOS,
'  PERIODIZACION, PLAN_EFG, PLAN_EFE, PLAN_EFC, MICROCICLO_* y RESUMEN).
'  Esta macro solo: pide los datos, los escribe en DATOS, ajusta las
'  filas visibles según el ciclo e imprime / exporta a PDF.
'======================================================================
Option Explicit

' Contraseña de protección de las hojas ("" = sin contraseña)
Private Const CLAVE As String = ""
Private Const TITULO As String = "Planificación del entrenamiento - TKD"

Private Const COLOR_VERDE As Long = 3837496      ' RGB(56,142,58)
Private Const COLOR_AZUL As Long = 10049310      ' RGB(30,87,153)
Private Const COLOR_NARANJA As Long = 1213670    ' RGB(230,132,18)
Private Const COLOR_ROJO As Long = 3355596       ' RGB(204,51,51)
Private Const COLOR_GRIS As Long = 6316128       ' RGB(96,96,96)

'======================================================================
'  ARRANQUE (lo llama ThisWorkbook.Workbook_Open)
'======================================================================
Public Sub IniciarLibro()
    On Error GoTo Fallo
    Application.ScreenUpdating = False

    DesprotegerTodo
    HojaMenu.Rows(4).Hidden = True          ' aviso "habilite macros": ya no hace falta
    ColocarBotonesMenu
    ColocarBotonesEnHojas
    AjustarFilasPlan
    ProtegerTodo

    HojaMenu.Activate
    HojaMenu.Range("A1").Select
    Application.ScreenUpdating = True

    If MsgBox("Bienvenido(a) a la planificación del entrenamiento." & vbCrLf & vbCrLf & _
              "¿Desea crear ahora un NUEVO PLAN para un deportista?", _
              vbQuestion + vbYesNo, TITULO) = vbYes Then
        NuevoPlan
    End If
    Exit Sub
Fallo:
    Application.ScreenUpdating = True
    MsgBox "No se pudo preparar el menú: " & Err.Description, vbExclamation, TITULO
    ProtegerTodo
End Sub

Public Sub IrAlMenu()
    HojaMenu.Activate
    HojaMenu.Range("A1").Select
End Sub

'======================================================================
'  OPCIÓN 1: NUEVO PLAN (preguntas paso a paso)
'======================================================================
Public Sub NuevoPlan()
    Dim ciclo As String, opcion As Double, inicio As Date
    Dim nombre As String, sexo As String
    Dim edad As Double, peso As Double, vpr As Double, vpb As Double
    Dim v(1 To 7) As Double
    Dim ejercicio As Variant, formacion As Variant
    Dim ok As Boolean, i As Integer, resumen As String

    On Error GoTo Fallo

    ejercicio = Array("CROSS (carrera)", "ESCALERA", "PATADAS", "SACO", "CANILLERA", _
                      "COMBATE CONDICIONADO", "COMBATE LIBRE")
    formacion = Array("General - EFG", "General - EFG", "Especial - EFE", "Especial - EFE", _
                      "Especial - EFE", "Competitiva - EFC", "Competitiva - EFC")

    ' Valores iniciales = los que tiene el libro ahora
    With HojaDatos
        nombre = CStr(.Range("B3").Value)
        ciclo = UCase$(CStr(.Range("B4").Value))
        If IsDate(.Range("B5").Value) Then inicio = .Range("B5").Value Else inicio = Date
        edad = ValorNum(.Range("B6").Value, 20)
        sexo = UCase$(Trim$(CStr(.Range("B7").Value)))
        If sexo <> "M" And sexo <> "F" Then sexo = "M"
        peso = ValorNum(.Range("B8").Value, 60)
        vpr = ValorNum(.Range("B9").Value, 60)
        vpb = ValorNum(.Range("B10").Value, vpr)
        For i = 1 To 7
            v(i) = ValorNum(.Range("B" & (11 + i)).Value, 0)
        Next i
    End With

    Do
        ' ---------- Tipo de ciclo ----------
        opcion = PedirNumero("Paso 1 de 16" & vbCrLf & vbCrLf & _
                 "¿Qué tipo de ciclo desea planificar?" & vbCrLf & vbCrLf & _
                 "   Escriba  1  para CICLO ANUAL (48 semanas)" & vbCrLf & _
                 "   Escriba  2  para CICLO SEMESTRAL (24 semanas)", _
                 IIf(ciclo = "SEMESTRAL", 2, 1), 1, 2, ok)
        If Not ok Then GoTo Cancelado
        If Int(opcion) = 2 Then ciclo = "SEMESTRAL" Else ciclo = "ANUAL"

        inicio = PedirFecha("Paso 2 de 16" & vbCrLf & vbCrLf & _
                 "FECHA DE INICIO del ciclo (día/mes/año)." & vbCrLf & _
                 "Se recomienda que sea un LUNES.", inicio, ok)
        If Not ok Then GoTo Cancelado

        ' ---------- Datos personales ----------
        nombre = PedirTexto("Paso 3 de 16" & vbCrLf & vbCrLf & _
                 "Escriba el NOMBRE COMPLETO del deportista:", nombre, ok)
        If Not ok Then GoTo Cancelado

        edad = PedirNumero("Paso 4 de 16" & vbCrLf & vbCrLf & "EDAD del deportista (en años):", _
               edad, 5, 90, ok)
        If Not ok Then GoTo Cancelado

        sexo = PedirSexo("Paso 5 de 16" & vbCrLf & vbCrLf & "SEXO del deportista:" & vbCrLf & _
               "   Escriba  M  para masculino" & vbCrLf & "   Escriba  F  para femenino", sexo, ok)
        If Not ok Then GoTo Cancelado

        peso = PedirNumero("Paso 6 de 16" & vbCrLf & vbCrLf & _
               "PESO CORPORAL del deportista (en kilogramos)." & vbCrLf & _
               "Se usa para calcular el peso de las canilleras.", peso, 15, 200, ok)
        If Not ok Then GoTo Cancelado

        vpr = PedirNumero("Paso 7 de 16" & vbCrLf & vbCrLf & _
              "VPR - Velocidad del pulso en REPOSO" & vbCrLf & _
              "(pulsaciones por minuto, sentado y tranquilo):", vpr, 30, 120, ok)
        If Not ok Then GoTo Cancelado

        vpb = PedirNumero("Paso 8 de 16" & vbCrLf & vbCrLf & _
              "VPB - Pulso BASAL" & vbCrLf & _
              "(pulsaciones por minuto, al despertar y antes de levantarse):", vpb, 30, 120, ok)
        If Not ok Then GoTo Cancelado

        ' ---------- Cantidades al 100 % ----------
        For i = 1 To 7
            v(i) = PedirNumero("Paso " & (8 + i) & " de 16" & vbCrLf & vbCrLf & _
                   "Ejercicio: " & ejercicio(i - 1) & "   (Formación " & formacion(i - 1) & ")" & vbCrLf & vbCrLf & _
                   "Escriba la cantidad que equivale al 100 % del esfuerzo" & vbCrLf & _
                   "del deportista en este ejercicio (su máximo estimado):", v(i), 0.1, 100000, ok)
            If Not ok Then GoTo Cancelado
        Next i

        ' ---------- Confirmación ----------
        resumen = "Paso 16 de 16 - Revise los datos:" & vbCrLf & vbCrLf & _
            "Ciclo:  " & ciclo & IIf(ciclo = "ANUAL", " (48 semanas)", " (24 semanas)") & _
            "      Inicio:  " & Format(inicio, "dd/mm/yyyy") & vbCrLf & _
            "Deportista:  " & nombre & vbCrLf & _
            "Edad:  " & edad & " años      Sexo:  " & IIf(sexo = "M", "Masculino", "Femenino") & _
            "      Peso:  " & peso & " kg" & vbCrLf & _
            "VPR:  " & vpr & " ppm      VPB:  " & vpb & " ppm" & vbCrLf & vbCrLf & _
            "Cantidades al 100 %:" & vbCrLf & _
            "   Cross: " & v(1) & "    Escalera: " & v(2) & vbCrLf & _
            "   Patadas: " & v(3) & "    Saco: " & v(4) & "    Canillera: " & v(5) & vbCrLf & _
            "   Combate condicionado: " & v(6) & "    Combate libre: " & v(7) & vbCrLf & vbCrLf & _
            "¿Los datos son correctos?" & vbCrLf & _
            "(Sí = calcular el plan     No = corregir     Cancelar = salir sin cambios)"

        Select Case MsgBox(resumen, vbQuestion + vbYesNoCancel, TITULO)
            Case vbYes: Exit Do
            Case vbCancel: GoTo Cancelado
        End Select
    Loop

    ' ¿Borrar el registro de sobreentrenamiento del deportista anterior?
    If UltimaFilaIndice() >= 7 Then
        If MsgBox("La hoja del índice de sobreentrenamiento tiene registros." & vbCrLf & _
                  "¿Desea BORRARLOS para empezar el registro de " & nombre & "?", _
                  vbQuestion + vbYesNo, TITULO) = vbYes Then
            HojaIndice.Unprotect CLAVE
            HojaIndice.Range("A7:C372").ClearContents
        End If
    End If

    ' ---------- Escribir los datos en la hoja DATOS ----------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    DesprotegerTodo
    With HojaDatos
        .Range("B3").Value = nombre
        .Range("B4").Value = ciclo
        .Range("B5").Value = inicio
        .Range("B6").Value = edad
        .Range("B7").Value = sexo
        .Range("B8").Value = peso
        .Range("B9").Value = vpr
        .Range("B10").Value = vpb
        For i = 1 To 7
            .Range("B" & (11 + i)).Value = v(i)
        Next i
    End With
    Application.EnableEvents = True

    ' ---------- Calcular ----------
    Application.CalculateFull
    AjustarFilasPlan
    ProtegerTodo
    HojaMenu.Activate
    Application.ScreenUpdating = True

    If MsgBox("¡Plan calculado para " & nombre & "!" & vbCrLf & vbCrLf & _
              "Ciclo " & ciclo & ": del " & Format(inicio, "dd/mm/yyyy") & _
              " al " & Format(HojaDatos.Range("B23").Value, "dd/mm/yyyy") & vbCrLf & _
              "FC máxima estimada: " & Format(HojaDatos.Range("B21").Value, "0") & " ppm" & vbCrLf & _
              "Semana 1 (EFG): Cross = " & Format(HojaPlanEFG.Range("G6").Value, "0.0") & _
              "    VPE = " & Format(HojaPlanEFG.Range("J6").Value, "0") & " ppm" & vbCrLf & vbCrLf & _
              "El plan tiene: resumen, volúmenes e intensidades (EFG, EFE, EFC)," & vbCrLf & _
              "microciclos semanales y registro de sobreentrenamiento." & vbCrLf & vbCrLf & _
              "¿Desea GUARDAR ahora el plan completo en PDF?", _
              vbQuestion + vbYesNo, TITULO) = vbYes Then
        GuardarPDF
    End If
    Exit Sub

Cancelado:
    MsgBox "Proceso cancelado. No se modificó ningún dato.", vbInformation, TITULO
    Exit Sub
Fallo:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Ocurrió un error al guardar los datos: " & Err.Description, vbExclamation, TITULO
    ProtegerTodo
End Sub

'======================================================================
'  OPCIONES 2 a 4: VER HOJAS
'======================================================================
Public Sub VerResumen()
    MostrarHoja HojaResumen
End Sub

Public Sub VerPlan()
    MostrarHoja HojaPlanEFG
End Sub

Public Sub VerMicrociclos()
    MostrarHoja HojaMicroEPG
End Sub

Public Sub VerIndice()
    MostrarHoja HojaIndice
End Sub

Private Sub MostrarHoja(ByVal ws As Worksheet)
    ws.Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ws.Range("A1").Select
End Sub

'======================================================================
'  OPCIÓN 5: REGISTRAR EL ÍNDICE DE SOBREENTRENAMIENTO
'======================================================================
Public Sub RegistrarIndice()
    Dim fecha As Date, acostado As Double, depie As Double, ok As Boolean
    Dim fila As Long, dif As Double

    On Error GoTo Fallo
    fecha = PedirFecha("ÍNDICE DE SOBREENTRENAMIENTO - " & NombreActual() & vbCrLf & vbCrLf & _
            "Fecha de la medición (día/mes/año):", Date, ok)
    If Not ok Then Exit Sub
    acostado = PedirNumero("Pulso ACOSTADO (pulsaciones por minuto)," & vbCrLf & _
               "tomado al despertar:", 60, 25, 150, ok)
    If Not ok Then Exit Sub
    depie = PedirNumero("Pulso DE PIE (pulsaciones por minuto)," & vbCrLf & _
            "tomado inmediatamente al levantarse:", acostado + 10, 25, 220, ok)
    If Not ok Then Exit Sub

    fila = UltimaFilaIndice() + 1
    If fila < 7 Then fila = 7
    If fila > 372 Then
        MsgBox "La hoja de registro está llena (366 registros). Borre registros antiguos.", _
               vbExclamation, TITULO
        Exit Sub
    End If

    HojaIndice.Unprotect CLAVE
    HojaIndice.Cells(fila, 1).Value = fecha
    HojaIndice.Cells(fila, 2).Value = acostado
    HojaIndice.Cells(fila, 3).Value = depie
    ProtegerHoja HojaIndice

    dif = depie - acostado
    If dif > 15 Then
        MsgBox "¡ALERTA!  Diferencia de " & dif & " pulsaciones (mayor de 15)." & vbCrLf & vbCrLf & _
               "Posible riesgo de SOBREENTRENAMIENTO. Revise con el deportista:" & vbCrLf & _
               "   - horas de sueño y trasnochos" & vbCrLf & _
               "   - consumo de bebidas alcohólicas" & vbCrLf & _
               "   - cargas de entrenamiento de los últimos días" & vbCrLf & _
               "   - posible enfermedad o malestar" & vbCrLf & vbCrLf & _
               "Si la alerta se repite, conviene que lo revise un médico.", vbExclamation, TITULO
    Else
        MsgBox "Diferencia de " & dif & " pulsaciones: NORMAL." & vbCrLf & _
               "Registro guardado en la fila " & fila & ".", vbInformation, TITULO
    End If
    Exit Sub
Fallo:
    MsgBox "No se pudo registrar: " & Err.Description, vbExclamation, TITULO
    ProtegerHoja HojaIndice
End Sub

'======================================================================
'  OPCIONES 6 a 8: VISTA PREVIA, IMPRIMIR, PDF
'======================================================================
Public Sub VistaPrevia()
    On Error GoTo Fallo
    If Not HayDatos() Then Exit Sub
    PrepararImpresion
    ThisWorkbook.Sheets(HojasDelPlan()).PrintPreview
    IrAlMenu
    Exit Sub
Fallo:
    MsgBox "No se pudo mostrar la vista previa: " & Err.Description, vbExclamation, TITULO
    IrAlMenu
End Sub

Public Sub ImprimirPlan()
    Dim copias As Double, ok As Boolean
    On Error GoTo Fallo
    If Not HayDatos() Then Exit Sub
    If MsgBox("Se imprimirá el plan completo de:" & vbCrLf & "   " & NombreActual() & vbCrLf & vbCrLf & _
              "(Resumen, EFG, EFE, EFC, microciclos e índice de sobreentrenamiento)" & vbCrLf & vbCrLf & _
              "Impresora:  " & Application.ActivePrinter & vbCrLf & _
              "(Para elegir otra impresora use la opción 'Vista previa')." & vbCrLf & vbCrLf & _
              "¿Desea continuar?", vbQuestion + vbOKCancel, TITULO) <> vbOK Then Exit Sub
    copias = PedirNumero("¿Cuántas copias desea imprimir?", 1, 1, 20, ok)
    If Not ok Then Exit Sub
    PrepararImpresion
    ThisWorkbook.Sheets(HojasDelPlan()).PrintOut Copies:=CLng(copias)
    IrAlMenu
    MsgBox "El plan fue enviado a la impresora.", vbInformation, TITULO
    Exit Sub
Fallo:
    MsgBox "No se pudo imprimir: " & Err.Description & vbCrLf & _
           "Verifique que la impresora esté encendida y conectada.", vbExclamation, TITULO
    IrAlMenu
End Sub

Public Sub ImprimirHojaActiva()
    On Error GoTo Fallo
    If MsgBox("¿Imprimir la hoja " & ActiveSheet.Name & " de " & NombreActual() & "?", _
              vbQuestion + vbOKCancel, TITULO) <> vbOK Then Exit Sub
    PrepararImpresion
    ActiveSheet.PrintOut Copies:=1
    Exit Sub
Fallo:
    MsgBox "No se pudo imprimir: " & Err.Description, vbExclamation, TITULO
End Sub

Public Sub GuardarPDF()
    Dim carpeta As String, archivo As Variant
    On Error GoTo Fallo
    If Not HayDatos() Then Exit Sub

    carpeta = ThisWorkbook.Path
    If carpeta = "" Then carpeta = CurDir$
    archivo = Application.GetSaveAsFilename( _
        InitialFileName:=carpeta & Application.PathSeparator & "Plan_" & _
                         NombreArchivo(NombreActual()) & "_" & LCase$(CStr(HojaDatos.Range("B4").Value)) & ".pdf", _
        FileFilter:="Archivo PDF (*.pdf), *.pdf", _
        Title:="Guardar el plan de entrenamiento en PDF")
    If VarType(archivo) = vbBoolean Then Exit Sub        ' canceló

    PrepararImpresion
    ThisWorkbook.Sheets(HojasDelPlan()).Select
    ActiveSheet.ExportAsFixedFormat Type:=0, Filename:=CStr(archivo), _
        Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    HojaMenu.Select
    Exit Sub
Fallo:
    MsgBox "No se pudo crear el PDF: " & Err.Description, vbExclamation, TITULO
    HojaMenu.Select
End Sub

Public Sub GuardarLibro()
    On Error GoTo Fallo
    ThisWorkbook.Save
    MsgBox "El libro se guardó correctamente.", vbInformation, TITULO
    Exit Sub
Fallo:
    MsgBox "No se pudo guardar: " & Err.Description, vbExclamation, TITULO
End Sub

' Hojas que forman el plan impreso / PDF, en orden
Private Function HojasDelPlan() As Variant
    HojasDelPlan = Array(HojaResumen.Name, HojaPlanEFG.Name, HojaPlanEFE.Name, HojaPlanEFC.Name, _
                         HojaMicroEPG.Name, HojaMicroEPE.Name, HojaMicroEPC.Name, HojaIndice.Name)
End Function

' Área de impresión del índice: los registros + un formulario de 4 semanas en blanco
Private Sub PrepararImpresion()
    Dim ultima As Long
    On Error Resume Next
    ultima = UltimaFilaIndice() + 28
    If ultima < 34 Then ultima = 34
    If ultima > 372 Then ultima = 372
    HojaIndice.PageSetup.PrintArea = "$A$1:$E$" & ultima
End Sub

'======================================================================
'  FILAS VISIBLES SEGÚN EL CICLO (48 o 24 semanas)
'  También se llama desde la hoja DATOS si se cambia el ciclo a mano.
'======================================================================
Public Sub AjustarFilasPlan()
    Dim hojas As Variant, i As Integer, nsem As Long
    On Error Resume Next
    nsem = CLng(ValorNum(HojaDatos.Range("B20").Value, 48))
    If nsem < 1 Or nsem > 48 Then nsem = 48
    hojas = Array(HojaPlanEFG, HojaPlanEFE, HojaPlanEFC)
    For i = 0 To 2
        hojas(i).Unprotect CLAVE
        hojas(i).Rows("6:53").Hidden = False
        If nsem < 48 Then hojas(i).Rows((6 + nsem) & ":53").Hidden = True
        ProtegerHoja hojas(i)
    Next i
End Sub

'======================================================================
'  BOTONES
'======================================================================
Private Sub ColocarBotonesMenu()
    Dim ws As Worksheet, x As Double, y As Double, ancho As Double, i As Integer
    Dim textos As Variant, macros As Variant, colores As Variant

    Set ws = HojaMenu
    BorrarBotones ws
    textos = Array("1.  NUEVO PLAN: ingresar datos del deportista", _
                   "2.  Ver el resumen del plan", _
                   "3.  Ver volúmenes e intensidades (EFG, EFE, EFC)", _
                   "4.  Ver los microciclos semanales", _
                   "5.  Registrar índice de sobreentrenamiento", _
                   "6.  Vista previa del plan completo", _
                   "7.  Imprimir el plan completo", _
                   "8.  Guardar el plan completo en PDF", _
                   "9.  Guardar el libro")
    macros = Array("NuevoPlan", "VerResumen", "VerPlan", "VerMicrociclos", "RegistrarIndice", _
                   "VistaPrevia", "ImprimirPlan", "GuardarPDF", "GuardarLibro")
    colores = Array(COLOR_VERDE, COLOR_AZUL, COLOR_AZUL, COLOR_AZUL, COLOR_ROJO, _
                    COLOR_NARANJA, COLOR_NARANJA, COLOR_NARANJA, COLOR_GRIS)
    x = ws.Range("G6").Left
    y = ws.Range("G6").Top + 3
    ancho = ws.Range("G6").Width
    For i = 0 To 8
        AgregarBoton ws, "btnMenu" & (i + 1), CStr(textos(i)), CStr(macros(i)), _
                     x, y + i * 46, ancho, 38, CLng(colores(i)), 13
    Next i
End Sub

' "Volver al MENÚ" e "Imprimir esta hoja" a la derecha de cada hoja del plan
Private Sub ColocarBotonesEnHojas()
    Dim hojas As Variant, i As Integer, ws As Worksheet, x As Double, shp As Object
    Dim rng As Range

    hojas = Array(HojaResumen, HojaPlanEFG, HojaPlanEFE, HojaPlanEFC, _
                  HojaMicroEPG, HojaMicroEPE, HojaMicroEPC, HojaIndice)
    For i = 0 To UBound(hojas)
        Set ws = hojas(i)
        BorrarBotones ws
        Set rng = ws.Range("A1").MergeArea               ' el título ocupa todo el ancho
        x = rng.Left + rng.Width + 10
        Set shp = AgregarBoton(ws, "btnVolver", "<<  Volver al MENÚ", "IrAlMenu", _
                               x, 4, 160, 30, COLOR_VERDE, 11)
        NoImprimir shp
        Set shp = AgregarBoton(ws, "btnImprimir", "Imprimir esta hoja", "ImprimirHojaActiva", _
                               x, 40, 160, 30, COLOR_NARANJA, 11)
        NoImprimir shp
        If ws Is HojaPlanEFG Or ws Is HojaPlanEFE Or ws Is HojaMicroEPG Or ws Is HojaMicroEPE Then
            Set shp = AgregarBoton(ws, "btnSiguiente", "Siguiente hoja  >>", "HojaSiguiente", _
                                   x, 76, 160, 30, COLOR_AZUL, 11)
            NoImprimir shp
        End If
    Next i
End Sub

' Pasa de PLAN_EFG -> PLAN_EFE -> PLAN_EFC y de EPG -> EPE -> EPC
Public Sub HojaSiguiente()
    Select Case ActiveSheet.CodeName
        Case "HojaPlanEFG": MostrarHoja HojaPlanEFE
        Case "HojaPlanEFE": MostrarHoja HojaPlanEFC
        Case "HojaMicroEPG": MostrarHoja HojaMicroEPE
        Case "HojaMicroEPE": MostrarHoja HojaMicroEPC
    End Select
End Sub

Private Function AgregarBoton(ByVal ws As Worksheet, ByVal nombreForma As String, _
        ByVal texto As String, ByVal macro As String, _
        ByVal x As Double, ByVal y As Double, ByVal ancho As Double, ByVal alto As Double, _
        ByVal colorFondo As Long, ByVal tamLetra As Single) As Object
    Dim shp As Object
    Set shp = ws.Shapes.AddShape(5, x, y, ancho, alto)    ' 5 = rectángulo redondeado
    With shp
        .Name = nombreForma
        .OnAction = macro
        .Placement = 3
        .Fill.ForeColor.RGB = colorFondo
        .Line.Visible = 0
        With .TextFrame2
            .VerticalAnchor = 3
            .MarginLeft = 12
            .TextRange.Text = texto
            .TextRange.Font.Size = tamLetra
            .TextRange.Font.Bold = -1
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = 1
        End With
    End With
    Set AgregarBoton = shp
End Function

' Evita que el botón salga impreso (PrintObject está en el objeto de dibujo, no en Shape)
Private Sub NoImprimir(ByVal shp As Object)
    On Error Resume Next
    shp.OLEFormat.Object.PrintObject = False
End Sub

Private Sub BorrarBotones(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If LCase$(Left$(ws.Shapes(i).Name, 3)) = "btn" Then ws.Shapes(i).Delete
    Next i
End Sub

'======================================================================
'  PROTECCIÓN
'======================================================================
Private Function HojasProtegidas() As Variant
    HojasProtegidas = Array(HojaMenu, HojaResumen, HojaPlanEFG, HojaPlanEFE, HojaPlanEFC, _
                            HojaMicroEPG, HojaMicroEPE, HojaMicroEPC, HojaIndice, HojaDatos, _
                            HojaPeriodos, HojaEFG, HojaEFE, HojaEFC)
End Function

Private Sub ProtegerHoja(ByVal ws As Object)
    On Error Resume Next
    ws.Unprotect CLAVE
    ws.Protect Password:=CLAVE, DrawingObjects:=True, Contents:=True, _
               Scenarios:=True, UserInterfaceOnly:=True
End Sub

Private Sub ProtegerTodo()
    Dim ws As Variant
    For Each ws In HojasProtegidas()
        ProtegerHoja ws
    Next ws
End Sub

Private Sub DesprotegerTodo()
    Dim ws As Variant
    On Error Resume Next
    For Each ws In HojasProtegidas()
        ws.Unprotect CLAVE
    Next ws
End Sub

'======================================================================
'  UTILIDADES
'======================================================================
Private Function UltimaFilaIndice() As Long
    UltimaFilaIndice = HojaIndice.Cells(HojaIndice.Rows.Count, 1).End(-4162).Row   ' -4162 = xlUp
    If UltimaFilaIndice < 7 Then UltimaFilaIndice = 6
End Function

Private Function NombreActual() As String
    NombreActual = Trim$(CStr(HojaDatos.Range("B3").Value))
    If NombreActual = "" Then NombreActual = "(sin nombre)"
End Function

Private Function HayDatos() As Boolean
    If ValorNum(HojaDatos.Range("B6").Value, 0) <= 0 Or _
       ValorNum(HojaDatos.Range("B9").Value, 0) <= 0 Then
        MsgBox "Primero cree el plan del deportista (opción 1 del menú).", vbExclamation, TITULO
        HayDatos = False
    Else
        HayDatos = True
    End If
End Function

Private Function ValorNum(ByVal v As Variant, ByVal porDefecto As Double) As Double
    If IsNumeric(v) And Not IsEmpty(v) Then
        ValorNum = CDbl(v)
    Else
        ValorNum = porDefecto
    End If
End Function

Private Function NombreArchivo(ByVal texto As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        texto = Replace(texto, c, "_")
    Next c
    NombreArchivo = Replace(Trim$(texto), " ", "_")
End Function

'---------------- Cuadros de diálogo con validación ----------------

Private Function PedirNumero(ByVal mensaje As String, ByVal porDefecto As Double, _
        ByVal minimo As Double, ByVal maximo As Double, ByRef ok As Boolean) As Double
    Dim r As Variant
    Do
        r = Application.InputBox(Prompt:=mensaje, Title:=TITULO, Default:=porDefecto, Type:=1)
        If VarType(r) = vbBoolean Then          ' Cancelar
            ok = False
            Exit Function
        End If
        If r >= minimo And r <= maximo Then
            ok = True
            PedirNumero = r
            Exit Function
        End If
        MsgBox "El valor debe estar entre " & minimo & " y " & maximo & ".", vbExclamation, TITULO
    Loop
End Function

Private Function PedirTexto(ByVal mensaje As String, ByVal porDefecto As String, _
                            ByRef ok As Boolean) As String
    Dim r As Variant
    Do
        r = Application.InputBox(Prompt:=mensaje, Title:=TITULO, Default:=porDefecto, Type:=2)
        If VarType(r) = vbBoolean Then
            ok = False
            Exit Function
        End If
        If Trim$(CStr(r)) <> "" Then
            ok = True
            PedirTexto = Trim$(CStr(r))
            Exit Function
        End If
        MsgBox "Este dato no puede quedar vacío.", vbExclamation, TITULO
    Loop
End Function

Private Function PedirSexo(ByVal mensaje As String, ByVal porDefecto As String, _
                           ByRef ok As Boolean) As String
    Dim r As String
    Do
        r = UCase$(Left$(PedirTexto(mensaje, porDefecto, ok), 1))
        If Not ok Then Exit Function
        If r = "M" Or r = "F" Then
            PedirSexo = r
            Exit Function
        End If
        MsgBox "Escriba solamente la letra M (masculino) o F (femenino).", vbExclamation, TITULO
    Loop
End Function

' Pide una fecha dd/mm/aaaa. Si no es lunes, ofrece usar el lunes de esa semana.
Private Function PedirFecha(ByVal mensaje As String, ByVal porDefecto As Date, _
                            ByRef ok As Boolean) As Date
    Dim r As String, d As Date, lunes As Date
    Do
        r = PedirTexto(mensaje, Format(porDefecto, "dd/mm/yyyy"), ok)
        If Not ok Then Exit Function
        If IsDate(r) Then
            d = DateValue(r)
            If InStr(mensaje, "LUNES") > 0 And Weekday(d, vbMonday) <> 1 Then
                lunes = d - Weekday(d, vbMonday) + 1
                If MsgBox("El " & Format(d, "dd/mm/yyyy") & " no es lunes." & vbCrLf & _
                          "¿Desea iniciar el ciclo el lunes " & Format(lunes, "dd/mm/yyyy") & "?" & vbCrLf & _
                          "(No = usar la fecha escrita)", vbQuestion + vbYesNo, TITULO) = vbYes Then d = lunes
            End If
            PedirFecha = d
            ok = True
            Exit Function
        End If
        MsgBox "Escriba la fecha como día/mes/año, por ejemplo: 01/02/2027", vbExclamation, TITULO
    Loop
End Function
