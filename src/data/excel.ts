import { L, type Command } from '../types'

export const excelCommands: Command[] = [
  {
    id: 'xl-screenupdating',
    name: 'Application.ScreenUpdating',
    syntax: 'Application.ScreenUpdating = Boolean',
    scope: ['excel'],
    category: 'xl-app',
    subcategory: 'engine',
    description: L(
      'Attiva o spegne il ridisegno della griglia. Impostalo a False prima di loop lunghi e ripristinalo a True nel Finally logico.',
      'Turns worksheet redraw on or off. Set False before long loops and restore True in a logical Finally.',
      'Activa o apaga el redibujado de la cuadrícula. Ponlo a False antes de bucles largos y restáuralo a True.',
      'Active ou coupe le redessin de la grille. Passez à False avant les longues boucles et remettez True.',
      'Schaltet das Neuzeichnen des Rasters ein oder aus. Vor langen Schleifen auf False setzen und danach True wiederherstellen.',
    ),
    params: [
      {
        name: 'ScreenUpdating',
        description: L(
          'True ridisegna; False sopprime sfarfallio e accelera le macro.',
          'True redraws; False suppresses flicker and speeds macros.',
          'True redibuja; False suprime el parpadeo y acelera las macros.',
          'True redessine ; False supprime le scintillement et accélère les macros.',
          'True zeichnet neu; False unterdrückt Flackern und beschleunigt Makros.',
        ),
      },
    ],
    examples: [
      {
        title: L('Loop senza sfarfallio', 'Loop without flicker', 'Bucle sin parpadeo', 'Boucle sans scintillement', 'Schleife ohne Flackern'),
        code: `Sub DemoScreenUpdating()
    Dim i As Long
    Application.ScreenUpdating = False
    For i = 1 To 20
        Cells(i, 1).Value = i * 10
    Next i
    Application.ScreenUpdating = True
    Debug.Print "A20=" & Range("A20").Value
End Sub`,
        result: L(
          'A1:A20 = 10,20,…,200. Immediate Window: A20=200. La griglia si aggiorna solo alla fine.',
          'A1:A20 = 10,20,…,200. Immediate Window: A20=200. The grid refreshes only at the end.',
          'A1:A20 = 10,20,…,200. Immediate Window: A20=200. La cuadrícula se actualiza solo al final.',
          'A1:A20 = 10,20,…,200. Immediate Window : A20=200. La grille se rafraîchit seulement à la fin.',
          'A1:A20 = 10,20,…,200. Direktbereich: A20=200. Das Raster wird erst am Ende aktualisiert.',
        ),
      },
    ],
    notes: L(
      'Usa sempre On Error GoTo per ripristinare True se la macro può fallire, altrimenti Excel resta “congelato”.',
      'Always restore True in an error handler if the macro can fail, or Excel stays visually frozen.',
      'Restaura siempre True en el controlador de errores o Excel queda visualmente congelado.',
      'Restaurez toujours True dans le gestionnaire d’erreurs sinon Excel reste figé.',
      'True immer im Fehlerhandler wiederherstellen, sonst bleibt Excel optisch eingefroren.',
    ),
    related: ['xl-enableevents', 'xl-calculation', 'xl-statusbar'],
  },
  {
    id: 'xl-calculation',
    name: 'Application.Calculation',
    syntax: 'Application.Calculation = xlCalculationAutomatic | xlCalculationManual | xlCalculationSemiautomatic',
    scope: ['excel'],
    category: 'xl-app',
    subcategory: 'engine',
    description: L(
      'Controlla quando Excel ricalcola le formule. Manuale evita ricalcoli intermedi durante scritture massive.',
      'Controls when Excel recalculates formulas. Manual avoids mid-loop recalculation during bulk writes.',
      'Controla cuándo Excel recalcula las fórmulas. Manual evita recálculos intermedios.',
      'Contrôle quand Excel recalcule les formules. Manuel évite les recalculs intermédiaires.',
      'Steuert, wann Excel Formeln neu berechnet. Manuell vermeidet Zwischenberechnungen.',
    ),
    params: [
      {
        name: 'Calculation',
        description: L(
          'xlCalculationAutomatic (-4105), xlCalculationManual (-4135), xlCalculationSemiautomatic (2).',
          'xlCalculationAutomatic (-4105), xlCalculationManual (-4135), xlCalculationSemiautomatic (2).',
          'xlCalculationAutomatic (-4105), xlCalculationManual (-4135), xlCalculationSemiautomatic (2).',
          'xlCalculationAutomatic (-4105), xlCalculationManual (-4135), xlCalculationSemiautomatic (2).',
          'xlCalculationAutomatic (-4105), xlCalculationManual (-4135), xlCalculationSemiautomatic (2).',
        ),
      },
    ],
    examples: [
      {
        title: L('Scrivi formule in modalità manuale', 'Write formulas in manual mode', 'Escribir fórmulas en modo manual', 'Écrire des formules en mode manuel', 'Formeln im manuellen Modus schreiben'),
        code: `Sub DemoCalculation()
    Application.Calculation = xlCalculationManual
    Range("A1").Value = 4
    Range("A2").Value = 6
    Range("A3").Formula = "=A1+A2"
    Debug.Print "Before Calculate: " & Range("A3").Text
    Application.Calculate
    Debug.Print "After Calculate: " & Range("A3").Value
    Application.Calculation = xlCalculationAutomatic
End Sub`,
        result: L(
          'A1=4, A2=6, A3=10. Immediate Window: Before Calculate: 0 (o vuoto) poi After Calculate: 10.',
          'A1=4, A2=6, A3=10. Immediate Window: Before Calculate: 0 (or blank) then After Calculate: 10.',
          'A1=4, A2=6, A3=10. Immediate Window: Before Calculate: 0 (o vacío) luego After Calculate: 10.',
          'A1=4, A2=6, A3=10. Immediate Window : Before Calculate : 0 (ou vide) puis After Calculate : 10.',
          'A1=4, A2=6, A3=10. Direktbereich: Before Calculate: 0 (oder leer), dann After Calculate: 10.',
        ),
      },
    ],
    notes: L(
      'Ripristina sempre Automatic. Application.Calculate ricalcola tutto; Range.Calculate solo l’intervallo.',
      'Always restore Automatic. Application.Calculate recalculates everything; Range.Calculate only that range.',
      'Restaura siempre Automatic. Application.Calculate recalcula todo; Range.Calculate solo el rango.',
      'Restaurez toujours Automatic. Application.Calculate recalcule tout ; Range.Calculate seulement la plage.',
      'Automatic immer wiederherstellen. Application.Calculate berechnet alles; Range.Calculate nur den Bereich.',
    ),
    related: ['xl-range-calculate', 'xl-screenupdating', 'xl-range-formula'],
  },
  {
    id: 'xl-enableevents',
    name: 'Application.EnableEvents',
    syntax: 'Application.EnableEvents = Boolean',
    scope: ['excel'],
    category: 'xl-app',
    subcategory: 'engine',
    description: L(
      'Abilita o disabilita gli eventi di cartella e foglio. Imposta False prima di scrivere celle da Worksheet_Change per evitare ricorsione.',
      'Enables or disables workbook and sheet events. Set False before writing cells from Worksheet_Change to avoid recursion.',
      'Activa o desactiva los eventos de libro y hoja. Pon False antes de escribir celdas desde Worksheet_Change.',
      'Active ou désactive les événements classeur/feuille. Mettez False avant d’écrire des cellules depuis Worksheet_Change.',
      'Aktiviert oder deaktiviert Mappen- und Blatt-Ereignisse. Vor Zellschreibvorgängen aus Worksheet_Change auf False setzen.',
    ),
    params: [
      {
        name: 'EnableEvents',
        description: L(
          'True (predefinito) solleva eventi; False li sopprime per il thread VBA corrente.',
          'True (default) raises events; False suppresses them for the current VBA thread.',
          'True (predeterminado) dispara eventos; False los suprime en el hilo VBA actual.',
          'True (défaut) déclenche les événements ; False les coupe pour le thread VBA courant.',
          'True (Standard) löst Ereignisse aus; False unterdrückt sie im aktuellen VBA-Thread.',
        ),
      },
    ],
    examples: [
      {
        title: L('Scrivere senza rilanciare Change', 'Write without re-firing Change', 'Escribir sin relanzar Change', 'Écrire sans relancer Change', 'Schreiben ohne Change erneut'),
        code: `Sub DemoEnableEvents()
    Application.EnableEvents = False
    Range("A1").Value = "safe write"
    Application.EnableEvents = True
    Debug.Print "EnableEvents=" & Application.EnableEvents
End Sub`,
        result: L(
          'A1 = safe write. Worksheet_Change non parte. Immediate Window: EnableEvents=True.',
          'A1 = safe write. Worksheet_Change does not run. Immediate Window: EnableEvents=True.',
          'A1 = safe write. Worksheet_Change no se ejecuta. Immediate Window: EnableEvents=True.',
          'A1 = safe write. Worksheet_Change ne s’exécute pas. Immediate Window : EnableEvents=True.',
          'A1 = safe write. Worksheet_Change läuft nicht. Direktbereich: EnableEvents=True.',
        ),
      },
    ],
    notes: L(
      'Se dimentichi di rimettere True, Open/Change/Activate smettono di funzionare fino al riavvio di Excel.',
      'If you forget to set True again, Open/Change/Activate stop firing until Excel is restarted.',
      'Si olvidas True, Open/Change/Activate dejan de dispararse hasta reiniciar Excel.',
      'Si vous oubliez True, Open/Change/Activate cessent jusqu’au redémarrage d’Excel.',
      'Ohne True feuern Open/Change/Activate erst nach Excel-Neustart wieder.',
    ),
    related: ['xl-worksheet-change', 'xl-displayalerts', 'xl-screenupdating'],
  },
  {
    id: 'xl-displayalerts',
    name: 'Application.DisplayAlerts',
    syntax: 'Application.DisplayAlerts = Boolean',
    scope: ['excel'],
    category: 'xl-app',
    subcategory: 'engine',
    description: L(
      'Mostra o nasconde i dialoghi nativi (conferma elimina foglio, sovrascrivi file). False sceglie automaticamente Sì/predefinito.',
      'Shows or hides native dialogs (delete sheet, overwrite file). False auto-picks Yes/default.',
      'Muestra u oculta los diálogos nativos (eliminar hoja, sobrescribir). False elige Sí/predeterminado.',
      'Affiche ou masque les dialogues natifs (supprimer feuille, écraser). False choisit Oui/défaut.',
      'Zeigt oder verbirgt native Dialoge (Blatt löschen, überschreiben). False wählt Ja/Standard.',
    ),
    params: [
      {
        name: 'DisplayAlerts',
        description: L(
          'True mostra avvisi; False li sopprime per la macro in corso.',
          'True shows alerts; False suppresses them for the running macro.',
          'True muestra avisos; False los suprime durante la macro.',
          'True affiche les alertes ; False les coupe pendant la macro.',
          'True zeigt Warnungen; False unterdrückt sie während des Makros.',
        ),
      },
    ],
    examples: [
      {
        title: L('Elimina foglio senza conferma', 'Delete a sheet without prompt', 'Eliminar hoja sin confirmación', 'Supprimer une feuille sans invite', 'Blatt ohne Nachfrage löschen'),
        code: `Sub DemoDisplayAlerts()
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "TempDel"
    Application.DisplayAlerts = False
    Worksheets("TempDel").Delete
    Application.DisplayAlerts = True
    Debug.Print "TempDel gone: " & (Not SheetExists("TempDel"))
End Sub

Function SheetExists(n As String) As Boolean
    Dim s As Worksheet
    On Error Resume Next
    Set s = Worksheets(n)
    SheetExists = Not s Is Nothing
End Function`,
        result: L(
          'Il foglio TempDel viene creato e rimosso senza MsgBox. Immediate Window: TempDel gone: True.',
          'Sheet TempDel is created then removed with no MsgBox. Immediate Window: TempDel gone: True.',
          'La hoja TempDel se crea y elimina sin MsgBox. Immediate Window: TempDel gone: True.',
          'La feuille TempDel est créée puis supprimée sans MsgBox. Immediate Window : TempDel gone: True.',
          'Blatt TempDel wird ohne MsgBox erstellt und gelöscht. Direktbereich: TempDel gone: True.',
        ),
      },
    ],
    related: ['xl-worksheet-delete', 'xl-workbook-close', 'xl-quit'],
  },
  {
    id: 'xl-statusbar',
    name: 'Application.StatusBar',
    syntax: 'Application.StatusBar = String | False',
    scope: ['excel'],
    category: 'xl-app',
    subcategory: 'engine',
    description: L(
      'Scrive un messaggio nella barra di stato in basso. Assegna False per restituire il controllo a Excel.',
      'Writes a message to the bottom status bar. Assign False to give control back to Excel.',
      'Escribe un mensaje en la barra de estado. Asigna False para devolver el control a Excel.',
      'Écrit un message dans la barre d’état. Assignez False pour rendre la main à Excel.',
      'Schreibt eine Meldung in die Statusleiste. False gibt die Kontrolle an Excel zurück.',
    ),
    params: [
      {
        name: 'StatusBar',
        description: L(
          'Testo da mostrare, oppure False per ripristinare “Pronto”.',
          'Text to show, or False to restore Ready.',
          'Texto a mostrar, o False para restaurar Listo.',
          'Texte à afficher, ou False pour restaurer Prêt.',
          'Anzuzeigender Text, oder False für Bereit.',
        ),
      },
    ],
    examples: [
      {
        title: L('Progresso in barra di stato', 'Progress on the status bar', 'Progreso en la barra de estado', 'Progression dans la barre d’état', 'Fortschritt in der Statusleiste'),
        code: `Sub DemoStatusBar()
    Dim i As Long
    For i = 1 To 5
        Application.StatusBar = "Row " & i & " of 5"
        Cells(i, 1).Value = i
    Next i
    Application.StatusBar = False
    Debug.Print "A5=" & Range("A5").Value
End Sub`,
        result: L(
          'A1:A5 = 1…5. La barra mostra “Row n of 5” poi torna a Pronto. Immediate Window: A5=5.',
          'A1:A5 = 1…5. The bar shows “Row n of 5” then returns to Ready. Immediate Window: A5=5.',
          'A1:A5 = 1…5. La barra muestra “Row n of 5” y vuelve a Listo. Immediate Window: A5=5.',
          'A1:A5 = 1…5. La barre affiche « Row n of 5 » puis Prêt. Immediate Window : A5=5.',
          'A1:A5 = 1…5. Die Leiste zeigt „Row n of 5“, dann Bereit. Direktbereich: A5=5.',
        ),
      },
    ],
    related: ['xl-screenupdating', 'xl-calculation'],
  },
  {
    id: 'xl-cutcopymode',
    name: 'Application.CutCopyMode',
    syntax: 'Application.CutCopyMode = False | xlCopy | xlCut',
    scope: ['excel'],
    category: 'xl-app',
    subcategory: 'engine',
    description: L(
      'Stato della selezione “formichine”. False cancella Appunti Excel e il bordo animato dopo Copy/Cut.',
      'State of the marching-ants selection. False clears Excel’s clipboard and the animated border after Copy/Cut.',
      'Estado de la selección de hormigas. False limpia el portapapeles de Excel y el borde animado.',
      'État de la sélection en pointillés. False vide le Presse-papiers Excel et le cadre animé.',
      'Status der Laufrahmen-Auswahl. False leert die Excel-Zwischenablage und den animierten Rahmen.',
    ),
    params: [
      {
        name: 'CutCopyMode',
        description: L(
          'False = cancella; xlCopy (1) / xlCut (2) sono valori letti dopo Copy/Taglia.',
          'False = clear; xlCopy (1) / xlCut (2) are values read after Copy/Cut.',
          'False = borrar; xlCopy (1) / xlCut (2) se leen tras Copiar/Cortar.',
          'False = effacer ; xlCopy (1) / xlCut (2) se lisent après Copier/Couper.',
          'False = löschen; xlCopy (1) / xlCut (2) nach Kopieren/Ausschneiden.',
        ),
      },
    ],
    examples: [
      {
        title: L('Copia e togli le formichine', 'Copy then clear marching ants', 'Copiar y quitar las hormigas', 'Copier puis ôter les pointillés', 'Kopieren und Laufrahmen löschen'),
        code: `Sub DemoCutCopyMode()
    Range("A1").Value = "hello"
    Range("A1").Copy Destination:=Range("B1")
    Debug.Print "Mode after Copy: " & Application.CutCopyMode
    Application.CutCopyMode = False
    Debug.Print "Mode after clear: " & Application.CutCopyMode
End Sub`,
        result: L(
          'B1 = hello. Immediate Window: Mode after Copy: 1 poi Mode after clear: 0. Il bordo tratteggiato scompare.',
          'B1 = hello. Immediate Window: Mode after Copy: 1 then Mode after clear: 0. The dashed border vanishes.',
          'B1 = hello. Immediate Window: Mode after Copy: 1 luego Mode after clear: 0. El borde discontinuo desaparece.',
          'B1 = hello. Immediate Window : Mode after Copy : 1 puis Mode after clear : 0. Le cadre disparaît.',
          'B1 = hello. Direktbereich: Mode after Copy: 1, dann Mode after clear: 0. Der Rahmen verschwindet.',
        ),
      },
    ],
    related: ['xl-range-copy-cut', 'xl-range-pastespecial'],
  },
  {
    id: 'xl-quit',
    name: 'Application.Quit',
    syntax: 'Application.Quit',
    scope: ['excel'],
    category: 'xl-app',
    subcategory: 'engine',
    description: L(
      'Chiude l’applicazione Excel. Con DisplayAlerts = False non chiede di salvare le cartelle sporche.',
      'Closes the Excel application. With DisplayAlerts = False it will not prompt to save dirty workbooks.',
      'Cierra la aplicación Excel. Con DisplayAlerts = False no pide guardar libros sucios.',
      'Ferme l’application Excel. Avec DisplayAlerts = False, pas d’invite pour les classeurs modifiés.',
      'Beendet die Excel-Anwendung. Mit DisplayAlerts = False keine Speicherabfrage für geänderte Mappen.',
    ),
    examples: [
      {
        title: L('Salva questa cartella e chiudi Excel', 'Save this workbook and quit Excel', 'Guardar este libro y salir de Excel', 'Enregistrer ce classeur et quitter Excel', 'Diese Mappe speichern und Excel beenden'),
        code: `Sub DemoQuit()
    If ThisWorkbook.Path = "" Then
        Debug.Print "Save the workbook first; Quit skipped."
        Exit Sub
    End If
    ThisWorkbook.Save
    Application.DisplayAlerts = False
    Application.Quit
End Sub`,
        result: L(
          'Se la cartella ha un percorso: viene salvata ed Excel termina. Altrimenti Immediate Window: Save the workbook first; Quit skipped.',
          'If the workbook has a path: it is saved and Excel exits. Otherwise Immediate Window: Save the workbook first; Quit skipped.',
          'Si el libro tiene ruta: se guarda y Excel termina. Si no, Immediate Window: Save the workbook first; Quit skipped.',
          'Si le classeur a un chemin : il est enregistré et Excel se ferme. Sinon Immediate Window : Save the workbook first; Quit skipped.',
          'Hat die Mappe einen Pfad: sie wird gespeichert und Excel endet. Sonst Direktbereich: Save the workbook first; Quit skipped.',
        ),
      },
    ],
    notes: L(
      'Non usare Quit in add-in caricati da altri processi se non vuoi chiudere tutta la sessione utente.',
      'Do not call Quit from add-ins loaded by other processes unless you intend to close the user’s session.',
      'No llames Quit desde complementos cargados por otros procesos si no quieres cerrar la sesión.',
      'N’appelez pas Quit depuis un complément chargé par un autre processus.',
      'Quit nicht aus Add-Ins aufrufen, die andere Prozesse geladen haben.',
    ),
    related: ['xl-workbook-save', 'xl-workbook-close', 'xl-displayalerts'],
  },
  {
    id: 'xl-run',
    name: 'Application.Run',
    syntax: 'Application.Run Macro, [Arg1], [Arg2], ...',
    scope: ['excel'],
    category: 'xl-app',
    subcategory: 'engine',
    description: L(
      'Esegue una Sub/Function per nome (anche in un’altra cartella). Restituisce il valore se è una Function.',
      'Runs a Sub/Function by name (including in another workbook). Returns the value when it is a Function.',
      'Ejecuta una Sub/Function por nombre (también en otro libro). Devuelve el valor si es Function.',
      'Exécute une Sub/Function par nom (y compris dans un autre classeur). Renvoie la valeur si c’est une Function.',
      'Führt eine Sub/Function namentlich aus (auch in einer anderen Mappe). Liefert den Wert bei einer Function.',
    ),
    params: [
      {
        name: 'Macro',
        description: L(
          'Nome procedura, o "\'Libro.xlsm\'!Modulo.Proc".',
          'Procedure name, or "\'Book.xlsm\'!Module.Proc".',
          'Nombre de procedimiento, o "\'Libro.xlsm\'!Modulo.Proc".',
          'Nom de procédure, ou "\'Classeur.xlsm\'!Module.Proc".',
          'Prozedurname oder "\'Mappe.xlsm\'!Modul.Proc".',
        ),
      },
      {
        name: 'Arg1…',
        optional: true,
        description: L(
          'Argomenti positional passati alla procedura.',
          'Positional arguments passed to the procedure.',
          'Argumentos posicionales pasados al procedimiento.',
          'Arguments positionnels transmis à la procédure.',
          'Positionsargumente an die Prozedur.',
        ),
      },
    ],
    examples: [
      {
        title: L('Chiama una Function con due argomenti', 'Call a Function with two arguments', 'Llamar una Function con dos argumentos', 'Appeler une Function avec deux arguments', 'Function mit zwei Argumenten aufrufen'),
        code: `Function AddTwo(a As Double, b As Double) As Double
    AddTwo = a + b
End Function

Sub DemoRun()
    Dim n As Double
    n = Application.Run("AddTwo", 3, 4)
    Range("A1").Value = n
    Debug.Print "A1=" & Range("A1").Value
End Sub`,
        result: L(
          'A1 = 7. Immediate Window: A1=7.',
          'A1 = 7. Immediate Window: A1=7.',
          'A1 = 7. Immediate Window: A1=7.',
          'A1 = 7. Immediate Window : A1=7.',
          'A1 = 7. Direktbereich: A1=7.',
        ),
      },
    ],
    related: ['xl-ontime', 'xl-this-active-workbook'],
  },
  {
    id: 'xl-ontime',
    name: 'Application.OnTime',
    syntax: 'Application.OnTime EarliestTime, Procedure, [LatestTime], [Schedule]',
    scope: ['excel'],
    category: 'xl-app',
    subcategory: 'engine',
    description: L(
      'Pianifica una procedura VBA a un orario. Conserva l’istante esatto in una variabile di modulo per poterlo annullare.',
      'Schedules a VBA procedure at a time. Store the exact instant in a module variable so you can cancel it.',
      'Programa un procedimiento VBA a una hora. Guarda el instante exacto en una variable de módulo para cancelarlo.',
      'Planifie une procédure VBA à une heure. Stockez l’instant exact dans une variable de module pour l’annuler.',
      'Plant eine VBA-Prozedur zu einem Zeitpunkt. Den exakten Zeitpunkt in einer Modulvariablen merken, um zu stornieren.',
    ),
    params: [
      {
        name: 'EarliestTime',
        description: L('Quando eseguire (Date/Time).', 'When to run (Date/Time).', 'Cuándo ejecutar (Date/Time).', 'Quand exécuter (Date/Time).', 'Ausführungszeitpunkt (Date/Time).'),
      },
      {
        name: 'Procedure',
        description: L('Nome della Sub pubblica da avviare.', 'Name of the public Sub to start.', 'Nombre de la Sub pública a iniciar.', 'Nom de la Sub publique à lancer.', 'Name der zu startenden öffentlichen Sub.'),
      },
      {
        name: 'Schedule',
        optional: true,
        description: L(
          'True (predefinito) prenota; False annulla la stessa coppia orario+procedura.',
          'True (default) books; False cancels the same time+procedure pair.',
          'True (predeterminado) reserva; False cancela el mismo par hora+procedimiento.',
          'True (défaut) réserve ; False annule la même paire heure+procédure.',
          'True (Standard) plant; False storniert dasselbe Zeit+Prozedur-Paar.',
        ),
      },
    ],
    examples: [
      {
        title: L('Esegui Ping tra 5 secondi', 'Run Ping in 5 seconds', 'Ejecutar Ping en 5 segundos', 'Exécuter Ping dans 5 secondes', 'Ping in 5 Sekunden ausführen'),
        code: `Public gPingWhen As Date

Sub SchedulePing()
    gPingWhen = Now + TimeValue("00:00:05")
    Application.OnTime gPingWhen, "Ping"
    Range("A1").Value = "scheduled"
    Debug.Print "OnTime set for " & Format(gPingWhen, "hh:nn:ss")
End Sub

Sub Ping()
    Range("A1").Value = "fired " & Format(Now, "hh:nn:ss")
End Sub

Sub CancelPing()
    On Error Resume Next
    Application.OnTime gPingWhen, "Ping", , False
End Sub`,
        result: L(
          'Subito A1 = scheduled. Dopo 5 s A1 = fired hh:nn:ss. CancelPing evita l’esecuzione se chiamato prima.',
          'Immediately A1 = scheduled. After 5 s A1 = fired hh:nn:ss. CancelPing prevents the run if called first.',
          'De inmediato A1 = scheduled. Tras 5 s A1 = fired hh:nn:ss. CancelPing evita la ejecución.',
          'Aussitôt A1 = scheduled. Après 5 s A1 = fired hh:nn:ss. CancelPing empêche l’exécution.',
          'Sofort A1 = scheduled. Nach 5 s A1 = fired hh:nn:ss. CancelPing verhindert den Lauf.',
        ),
      },
    ],
    notes: L(
      'Per annullare, EarliestTime deve essere identico al valore usato in prenotazione, non un nuovo Now.',
      'To cancel, EarliestTime must be the same value used when booking, not a fresh Now.',
      'Para cancelar, EarliestTime debe ser el mismo valor de la reserva, no un Now nuevo.',
      'Pour annuler, EarliestTime doit être la même valeur que lors de la réservation.',
      'Zum Stornieren muss EarliestTime derselbe Wert wie bei der Buchung sein, nicht ein neues Now.',
    ),
    related: ['xl-wait', 'xl-run'],
  },
  {
    id: 'xl-wait',
    name: 'Application.Wait',
    syntax: 'Application.Wait Now + TimeValue("hh:nn:ss")',
    scope: ['excel'],
    category: 'xl-app',
    subcategory: 'engine',
    description: L(
      'Blocca Excel fino a un orario. A differenza di OnTime non lascia usare l’interfaccia; va bene per pause brevi.',
      'Blocks Excel until a time. Unlike OnTime it does not leave the UI usable; fine for short pauses.',
      'Bloquea Excel hasta una hora. A diferencia de OnTime no deja usable la IU; vale para pausas cortas.',
      'Bloque Excel jusqu’à une heure. Contrairement à OnTime, l’UI n’est pas utilisable ; pour pauses courtes.',
      'Blockiert Excel bis zu einem Zeitpunkt. Anders als OnTime bleibt die UI unbenutzbar; für kurze Pausen.',
    ),
    params: [
      {
        name: 'Time',
        description: L(
          'Istante di ripresa, di solito Now + TimeValue("00:00:02").',
          'Resume instant, usually Now + TimeValue("00:00:02").',
          'Instante de reanudación, normalmente Now + TimeValue("00:00:02").',
          'Instant de reprise, en général Now + TimeValue("00:00:02").',
          'Fortsetzungszeitpunkt, meist Now + TimeValue("00:00:02").',
        ),
      },
    ],
    examples: [
      {
        title: L('Pausa di 2 secondi', '2-second pause', 'Pausa de 2 segundos', 'Pause de 2 secondes', '2-Sekunden-Pause'),
        code: `Sub DemoWait()
    Range("A1").Value = "start " & Format(Now, "hh:nn:ss")
    Application.Wait Now + TimeValue("00:00:02")
    Range("A2").Value = "end " & Format(Now, "hh:nn:ss")
    Debug.Print Range("A1").Value & " -> " & Range("A2").Value
End Sub`,
        result: L(
          'A1 e A2 differiscono di circa 2 secondi. Excel resta occupato durante l’attesa.',
          'A1 and A2 differ by about 2 seconds. Excel stays busy during the wait.',
          'A1 y A2 diferencian unos 2 segundos. Excel permanece ocupado.',
          'A1 et A2 diffèrent d’environ 2 secondes. Excel reste occupé.',
          'A1 und A2 unterscheiden sich um etwa 2 Sekunden. Excel bleibt beschäftigt.',
        ),
      },
    ],
    related: ['xl-ontime', 'xl-screenupdating'],
  },
  {
    id: 'xl-dialogs',
    name: 'Application.Dialogs',
    syntax: 'Application.Dialogs(xlDialogId).Show [Arg1], ...',
    scope: ['excel'],
    category: 'xl-app',
    subcategory: 'engine',
    description: L(
      'Apre un dialogo built-in di Excel (Apri, Salva con nome, Carattere, Stampa). Show restituisce True se l’utente conferma.',
      'Opens a built-in Excel dialog (Open, Save As, Font, Print). Show returns True if the user confirms.',
      'Abre un diálogo integrado de Excel (Abrir, Guardar como, Fuente, Imprimir). Show es True si confirma.',
      'Ouvre un dialogue Excel intégré (Ouvrir, Enregistrer sous, Police, Imprimer). Show vaut True si confirmé.',
      'Öffnet einen eingebauten Excel-Dialog (Öffnen, Speichern unter, Schrift, Drucken). Show ist True bei Bestätigung.',
    ),
    params: [
      {
        name: 'xlDialogId',
        description: L(
          'Costante: xlDialogOpen (1), xlDialogSaveAs (5), xlDialogPrint (8), xlDialogFont (26).',
          'Constant: xlDialogOpen (1), xlDialogSaveAs (5), xlDialogPrint (8), xlDialogFont (26).',
          'Constante: xlDialogOpen (1), xlDialogSaveAs (5), xlDialogPrint (8), xlDialogFont (26).',
          'Constante : xlDialogOpen (1), xlDialogSaveAs (5), xlDialogPrint (8), xlDialogFont (26).',
          'Konstante: xlDialogOpen (1), xlDialogSaveAs (5), xlDialogPrint (8), xlDialogFont (26).',
        ),
      },
    ],
    examples: [
      {
        title: L('Dialogo Carattere sulla cella attiva', 'Font dialog on the active cell', 'Diálogo Fuente en la celda activa', 'Dialogue Police sur la cellule active', 'Schrift-Dialog auf der aktiven Zelle'),
        code: `Sub DemoDialogs()
    Range("A1").Value = "Title"
    Range("A1").Select
    If Application.Dialogs(xlDialogFont).Show Then
        Debug.Print "Font applied: " & Range("A1").Font.Name
    Else
        Debug.Print "User cancelled"
    End If
End Sub`,
        result: L(
          'Si apre il dialogo Carattere. OK applica il font ad A1; Annulla stampa User cancelled.',
          'The Font dialog opens. OK applies the font to A1; Cancel prints User cancelled.',
          'Se abre el diálogo Fuente. Aceptar aplica la fuente a A1; Cancelar imprime User cancelled.',
          'Le dialogue Police s’ouvre. OK applique la police à A1 ; Annuler affiche User cancelled.',
          'Der Schrift-Dialog öffnet sich. OK wendet die Schrift auf A1 an; Abbrechen schreibt User cancelled.',
        ),
      },
    ],
    related: ['xl-workbooks-open', 'xl-workbook-save', 'xl-range-font'],
  },
  {
    id: 'xl-worksheetfunction',
    name: 'Application.WorksheetFunction',
    syntax: 'Application.WorksheetFunction.FunctionName(...)',
    scope: ['excel'],
    category: 'xl-app',
    subcategory: 'wsfn',
    description: L(
      'Espone le funzioni del foglio a VBA. WorksheetFunction.X solleva errore 1004 se X fallisce; Application.X restituisce un valore di errore CVErr.',
      'Exposes worksheet functions to VBA. WorksheetFunction.X raises 1004 if X fails; Application.X returns a CVErr error value.',
      'Expone las funciones de hoja a VBA. WorksheetFunction.X lanza 1004 si falla; Application.X devuelve CVErr.',
      'Expose les fonctions feuille à VBA. WorksheetFunction.X lève 1004 si échec ; Application.X renvoie CVErr.',
      'Stellt Tabellenfunktionen in VBA bereit. WorksheetFunction.X löst 1004 aus; Application.X liefert CVErr.',
    ),
    examples: [
      {
        title: L('Stessa SUM, due stili di errore', 'Same SUM, two error styles', 'La misma SUM, dos estilos de error', 'La même SOMME, deux styles d’erreur', 'Dieselbe SUMME, zwei Fehlerstile'),
        code: `Sub DemoWorksheetFunction()
    Range("A1:A3").Value = Application.Transpose(Array(10, 20, 30))
    Debug.Print Application.WorksheetFunction.Sum(Range("A1:A3"))
    Debug.Print IsError(Application.Match("nope", Range("A1:A3"), 0))
End Sub`,
        result: L(
          'A1:A3 = 10,20,30. Immediate Window: 60 poi True (Match non trovato via Application.Match).',
          'A1:A3 = 10,20,30. Immediate Window: 60 then True (no match via Application.Match).',
          'A1:A3 = 10,20,30. Immediate Window: 60 luego True (sin coincidencia con Application.Match).',
          'A1:A3 = 10,20,30. Immediate Window : 60 puis True (pas de match via Application.Match).',
          'A1:A3 = 10,20,30. Direktbereich: 60, dann True (kein Treffer über Application.Match).',
        ),
      },
    ],
    notes: L(
      'Preferisci Application.X + IsError quando il fallimento è previsto (CERCA.VERT, CONFRONTA).',
      'Prefer Application.X + IsError when failure is expected (VLOOKUP, MATCH).',
      'Prefiere Application.X + IsError cuando el fallo es esperado (VLOOKUP, MATCH).',
      'Préférez Application.X + IsError quand l’échec est attendu (RECHERCHEV, EQUIV).',
      'Application.X + IsError bevorzugen, wenn Fehlschlag erwartet ist (SVERWEIS, VERGLEICH).',
    ),
    related: ['xl-wsfn-sum', 'xl-wsfn-vlookup', 'xl-wsfn-iferror', 'xl-wsfn-match'],
  },
  {
    id: 'xl-wsfn-sum',
    name: 'WorksheetFunction.Sum',
    syntax: 'Application.WorksheetFunction.Sum(Arg1, [Arg2], ...)',
    scope: ['excel'],
    category: 'xl-app',
    subcategory: 'wsfn',
    description: L(
      'Somma numeri, intervalli o unioni come la funzione SOMMA del foglio. Ignora testo e vuoti.',
      'Adds numbers, ranges or unions like worksheet SUM. Ignores text and blanks.',
      'Suma números, rangos o uniones como SUMA. Ignora texto y vacíos.',
      'Additionne nombres, plages ou unions comme SOMME. Ignore texte et vides.',
      'Addiert Zahlen, Bereiche oder Vereinigungen wie SUMME. Ignoriert Text und Leerzellen.',
    ),
    params: [
      {
        name: 'Arg1…',
        description: L('Valori o Range da sommare (fino a 30 argomenti in VBA classico).', 'Values or Range objects to add (up to 30 arguments in classic VBA).', 'Valores o Range a sumar (hasta 30 argumentos en VBA clásico).', 'Valeurs ou Range à additionner (jusqu’à 30 arguments en VBA classique).', 'Werte oder Range-Objekte (bis 30 Argumente im klassischen VBA).'),
      },
    ],
    examples: [
      {
        title: L('Somma A1:A3', 'Sum A1:A3', 'Sumar A1:A3', 'Sommer A1:A3', 'A1:A3 summieren'),
        code: `Sub DemoSum()
    Range("A1").Value = 10
    Range("A2").Value = 15.5
    Range("A3").Value = "skip"
    Range("B1").Value = Application.WorksheetFunction.Sum(Range("A1:A3"))
    Debug.Print "B1=" & Range("B1").Value
End Sub`,
        result: L(
          'B1 = 25.5. Il testo in A3 è ignorato. Immediate Window: B1=25.5.',
          'B1 = 25.5. Text in A3 is ignored. Immediate Window: B1=25.5.',
          'B1 = 25.5. El texto en A3 se ignora. Immediate Window: B1=25.5.',
          'B1 = 25,5. Le texte en A3 est ignoré. Immediate Window : B1=25.5.',
          'B1 = 25.5. Text in A3 wird ignoriert. Direktbereich: B1=25.5.',
        ),
      },
    ],
    related: ['xl-worksheetfunction', 'xl-wsfn-counta', 'xl-union-intersect'],
  },
  {
    id: 'xl-wsfn-vlookup',
    name: 'WorksheetFunction.VLookup',
    syntax: 'Application.WorksheetFunction.VLookup(LookupValue, TableArray, ColIndex, [RangeLookup])',
    scope: ['excel'],
    category: 'xl-app',
    subcategory: 'wsfn',
    description: L(
      'Cerca LookupValue nella prima colonna di TableArray e restituisce la colonna ColIndex. RangeLookup False = corrispondenza esatta.',
      'Looks up LookupValue in the first column of TableArray and returns column ColIndex. RangeLookup False = exact match.',
      'Busca LookupValue en la primera columna de TableArray y devuelve ColIndex. RangeLookup False = exacto.',
      'Cherche LookupValue dans la 1re colonne de TableArray et renvoie ColIndex. RangeLookup False = exact.',
      'Sucht LookupValue in der ersten Spalte von TableArray und liefert Spalte ColIndex. RangeLookup False = exakt.',
    ),
    params: [
      { name: 'LookupValue', description: L('Valore da trovare.', 'Value to find.', 'Valor a buscar.', 'Valeur à trouver.', 'Zu suchender Wert.') },
      { name: 'TableArray', description: L('Intervallo tabella (chiave a sinistra).', 'Table range (key on the left).', 'Rango de tabla (clave a la izquierda).', 'Plage table (clé à gauche).', 'Tabellenbereich (Schlüssel links).') },
      { name: 'ColIndex', description: L('Indice colonna risultato (1-based).', 'Result column index (1-based).', 'Índice de columna resultado (base 1).', 'Index de colonne résultat (base 1).', 'Ergebnis-Spaltenindex (1-basiert).') },
      { name: 'RangeLookup', optional: true, description: L('False esatto; True approssimato (tabella ordinata).', 'False exact; True approximate (sorted table).', 'False exacto; True aproximado (tabla ordenada).', 'False exact ; True approximatif (table triée).', 'False exakt; True ungefähr (sortierte Tabelle).') },
    ],
    examples: [
      {
        title: L('Prezzo esatto di Alice', 'Exact price for Alice', 'Precio exacto de Alice', 'Prix exact d’Alice', 'Exakter Preis für Alice'),
        code: `Sub DemoVLookup()
    Range("A1").Value = "Name": Range("B1").Value = "Price"
    Range("A2").Value = "Alice": Range("B2").Value = 12.5
    Range("A3").Value = "Bob": Range("B3").Value = 9
    Range("D1").Value = Application.WorksheetFunction.VLookup("Alice", Range("A1:B3"), 2, False)
    Debug.Print "D1=" & Range("D1").Value
End Sub`,
        result: L(
          'Tabella A1:B3 con Alice/12.5 e Bob/9. D1 = 12.5. Immediate Window: D1=12.5.',
          'Table A1:B3 with Alice/12.5 and Bob/9. D1 = 12.5. Immediate Window: D1=12.5.',
          'Tabla A1:B3 con Alice/12.5 y Bob/9. D1 = 12.5. Immediate Window: D1=12.5.',
          'Table A1:B3 avec Alice/12,5 et Bob/9. D1 = 12.5. Immediate Window : D1=12.5.',
          'Tabelle A1:B3 mit Alice/12.5 und Bob/9. D1 = 12.5. Direktbereich: D1=12.5.',
        ),
      },
    ],
    notes: L(
      'Se la chiave manca, WorksheetFunction.VLookup solleva 1004. Usa Application.VLookup + IsError oppure IfError.',
      'If the key is missing, WorksheetFunction.VLookup raises 1004. Use Application.VLookup + IsError or IfError.',
      'Si falta la clave, WorksheetFunction.VLookup lanza 1004. Usa Application.VLookup + IsError o IfError.',
      'Si la clé manque, WorksheetFunction.VLookup lève 1004. Utilisez Application.VLookup + IsError ou IfError.',
      'Fehlt der Schlüssel, löst WorksheetFunction.VLookup 1004 aus. Application.VLookup + IsError oder IfError nutzen.',
    ),
    related: ['xl-wsfn-match', 'xl-wsfn-iferror', 'xl-range-find'],
  },
  {
    id: 'xl-wsfn-match',
    name: 'WorksheetFunction.Match',
    syntax: 'Application.WorksheetFunction.Match(LookupValue, LookupArray, [MatchType])',
    scope: ['excel'],
    category: 'xl-app',
    subcategory: 'wsfn',
    description: L(
      'Restituisce la posizione 1-based di LookupValue in LookupArray. MatchType 0 = esatto.',
      'Returns the 1-based position of LookupValue in LookupArray. MatchType 0 = exact.',
      'Devuelve la posición base 1 de LookupValue en LookupArray. MatchType 0 = exacto.',
      'Renvoie la position (base 1) de LookupValue dans LookupArray. MatchType 0 = exact.',
      'Liefert die 1-basierte Position von LookupValue in LookupArray. MatchType 0 = exakt.',
    ),
    params: [
      { name: 'LookupValue', description: L('Valore da localizzare.', 'Value to locate.', 'Valor a localizar.', 'Valeur à localiser.', 'Zu findender Wert.') },
      { name: 'LookupArray', description: L('Riga o colonna di ricerca.', 'Lookup row or column.', 'Fila o columna de búsqueda.', 'Ligne ou colonne de recherche.', 'Suchzeile oder -spalte.') },
      { name: 'MatchType', optional: true, description: L('0 esatto; 1 / -1 approssimato (ordinato).', '0 exact; 1 / -1 approximate (sorted).', '0 exacto; 1 / -1 aproximado (ordenado).', '0 exact ; 1 / -1 approximatif (trié).', '0 exakt; 1 / -1 ungefähr (sortiert).') },
    ],
    examples: [
      {
        title: L('Riga di Bob nella colonna A', 'Row of Bob in column A', 'Fila de Bob en la columna A', 'Ligne de Bob dans la colonne A', 'Zeile von Bob in Spalte A'),
        code: `Sub DemoMatch()
    Range("A1").Value = "Alice"
    Range("A2").Value = "Bob"
    Range("A3").Value = "Cara"
    Range("B1").Value = Application.WorksheetFunction.Match("Bob", Range("A1:A3"), 0)
    Debug.Print "B1=" & Range("B1").Value
End Sub`,
        result: L(
          'B1 = 2 (seconda cella di A1:A3). Immediate Window: B1=2.',
          'B1 = 2 (second cell of A1:A3). Immediate Window: B1=2.',
          'B1 = 2 (segunda celda de A1:A3). Immediate Window: B1=2.',
          'B1 = 2 (deuxième cellule de A1:A3). Immediate Window : B1=2.',
          'B1 = 2 (zweite Zelle von A1:A3). Direktbereich: B1=2.',
        ),
      },
    ],
    related: ['xl-wsfn-vlookup', 'xl-range-find', 'xl-worksheetfunction'],
  },
  {
    id: 'xl-wsfn-counta',
    name: 'WorksheetFunction.CountA',
    syntax: 'Application.WorksheetFunction.CountA(Arg1, [Arg2], ...)',
    scope: ['excel'],
    category: 'xl-app',
    subcategory: 'wsfn',
    description: L(
      'Conta le celle non vuote (numeri, testo, errori). Non conta le celle davvero vuote.',
      'Counts non-empty cells (numbers, text, errors). Does not count truly blank cells.',
      'Cuenta celdas no vacías (números, texto, errores). No cuenta las realmente vacías.',
      'Compte les cellules non vides (nombres, texte, erreurs). Ne compte pas les vides réels.',
      'Zählt nicht leere Zellen (Zahlen, Text, Fehler). Zählt wirklich leere Zellen nicht.',
    ),
    examples: [
      {
        title: L('Quante celle compilate in A1:A4', 'How many filled cells in A1:A4', 'Cuántas celdas llenas en A1:A4', 'Combien de cellules remplies dans A1:A4', 'Wie viele gefüllte Zellen in A1:A4'),
        code: `Sub DemoCountA()
    Range("A1").Value = "x"
    Range("A2").Value = 7
    Range("A3").ClearContents
    Range("A4").Value = ""
    Range("B1").Value = Application.WorksheetFunction.CountA(Range("A1:A4"))
    Debug.Print "B1=" & Range("B1").Value
End Sub`,
        result: L(
          'B1 = 2 (A1 e A2). A3 e A4 vuote non contano. Immediate Window: B1=2.',
          'B1 = 2 (A1 and A2). Blank A3 and A4 do not count. Immediate Window: B1=2.',
          'B1 = 2 (A1 y A2). A3 y A4 vacías no cuentan. Immediate Window: B1=2.',
          'B1 = 2 (A1 et A2). A3 et A4 vides ne comptent pas. Immediate Window : B1=2.',
          'B1 = 2 (A1 und A2). Leere A3 und A4 zählen nicht. Direktbereich: B1=2.',
        ),
      },
    ],
    related: ['xl-wsfn-sum', 'xl-range-specialcells', 'xl-range-clear'],
  },
  {
    id: 'xl-wsfn-iferror',
    name: 'WorksheetFunction.IfError',
    syntax: 'Application.WorksheetFunction.IfError(Value, ValueIfError)',
    scope: ['excel'],
    category: 'xl-app',
    subcategory: 'wsfn',
    description: L(
      'Restituisce Value se non è un errore, altrimenti ValueIfError. Il primo argomento deve già essere un valore (usa Application.VLookup, non WorksheetFunction.VLookup).',
      'Returns Value if it is not an error, otherwise ValueIfError. The first argument must already be a value (use Application.VLookup, not WorksheetFunction.VLookup).',
      'Devuelve Value si no es error; si no, ValueIfError. El primer argumento debe ser un valor (Application.VLookup).',
      'Renvoie Value s’il n’est pas une erreur, sinon ValueIfError. Le 1er argument doit déjà être une valeur (Application.VLookup).',
      'Liefert Value, wenn kein Fehler, sonst ValueIfError. Das erste Argument muss schon ein Wert sein (Application.VLookup).',
    ),
    params: [
      { name: 'Value', description: L('Espressione o risultato funzione (anche CVErr).', 'Expression or function result (including CVErr).', 'Expresión o resultado (incluido CVErr).', 'Expression ou résultat (y compris CVErr).', 'Ausdruck oder Funktionsergebnis (auch CVErr).') },
      { name: 'ValueIfError', description: L('Sostituto se Value è un errore.', 'Substitute when Value is an error.', 'Sustituto si Value es error.', 'Substitut si Value est une erreur.', 'Ersatz, wenn Value ein Fehler ist.') },
    ],
    examples: [
      {
        title: L('VLOOKUP protetto da IfError', 'VLOOKUP guarded by IfError', 'VLOOKUP protegido con IfError', 'RECHERCHEV protégé par IfError', 'SVERWEIS mit IfError abgesichert'),
        code: `Sub DemoIfError()
    Range("A1").Value = "Alice": Range("B1").Value = 12.5
    Range("A2").Value = "Bob": Range("B2").Value = 9
    Range("D1").Value = Application.WorksheetFunction.IfError( _
        Application.VLookup("Cara", Range("A1:B2"), 2, False), "n/a")
    Range("D2").Value = Application.WorksheetFunction.IfError( _
        Application.VLookup("Alice", Range("A1:B2"), 2, False), "n/a")
    Debug.Print Range("D1").Value & "," & Range("D2").Value
End Sub`,
        result: L(
          'D1 = n/a (Cara assente). D2 = 12.5. Immediate Window: n/a,12.5.',
          'D1 = n/a (Cara missing). D2 = 12.5. Immediate Window: n/a,12.5.',
          'D1 = n/a (falta Cara). D2 = 12.5. Immediate Window: n/a,12.5.',
          'D1 = n/a (Cara absente). D2 = 12.5. Immediate Window : n/a,12.5.',
          'D1 = n/a (Cara fehlt). D2 = 12.5. Direktbereich: n/a,12.5.',
        ),
      },
    ],
    related: ['xl-wsfn-vlookup', 'xl-worksheetfunction'],
  },
  {
    id: 'xl-workbooks-open',
    name: 'Workbooks.Open',
    syntax: 'Workbooks.Open Filename, [UpdateLinks], [ReadOnly], ...',
    scope: ['excel'],
    category: 'xl-books',
    subcategory: 'workbooks',
    description: L(
      'Apre una cartella da disco e la rende ActiveWorkbook. Restituisce l’oggetto Workbook.',
      'Opens a workbook from disk and makes it ActiveWorkbook. Returns the Workbook object.',
      'Abre un libro del disco y lo convierte en ActiveWorkbook. Devuelve el objeto Workbook.',
      'Ouvre un classeur depuis le disque et en fait ActiveWorkbook. Renvoie l’objet Workbook.',
      'Öffnet eine Mappe von Datenträger und macht sie zu ActiveWorkbook. Liefert das Workbook-Objekt.',
    ),
    params: [
      { name: 'Filename', description: L('Percorso completo del file.', 'Full path of the file.', 'Ruta completa del archivo.', 'Chemin complet du fichier.', 'Vollständiger Dateipfad.') },
      { name: 'ReadOnly', optional: true, description: L('True apre in sola lettura.', 'True opens read-only.', 'True abre de solo lectura.', 'True ouvre en lecture seule.', 'True öffnet schreibgeschützt.') },
    ],
    examples: [
      {
        title: L('Apri se esiste, altrimenti segnala', 'Open if present, otherwise report', 'Abrir si existe, si no informar', 'Ouvrir s’il existe, sinon signaler', 'Öffnen falls vorhanden, sonst melden'),
        code: `Sub DemoWorkbooksOpen()
    Dim p As String, wb As Workbook
    p = "C:/Data/Sales.xlsx"
    If Dir(p) = "" Then
        Range("A1").Value = "missing: " & p
        Debug.Print Range("A1").Value
        Exit Sub
    End If
    Set wb = Workbooks.Open(p)
    Debug.Print "Opened " & wb.Name
End Sub`,
        result: L(
          'Se C:/Data/Sales.xlsx esiste: la cartella si apre; Immediate Window: Opened Sales.xlsx. Altrimenti A1 = missing: C:/Data/Sales.xlsx.',
          'If C:/Data/Sales.xlsx exists: the workbook opens; Immediate Window: Opened Sales.xlsx. Else A1 = missing: C:/Data/Sales.xlsx.',
          'Si existe C:/Data/Sales.xlsx: se abre; Immediate Window: Opened Sales.xlsx. Si no, A1 = missing: C:/Data/Sales.xlsx.',
          'Si C:/Data/Sales.xlsx existe : le classeur s’ouvre ; Immediate Window : Opened Sales.xlsx. Sinon A1 = missing: C:/Data/Sales.xlsx.',
          'Wenn C:/Data/Sales.xlsx existiert: Mappe öffnet sich; Direktbereich: Opened Sales.xlsx. Sonst A1 = missing: C:/Data/Sales.xlsx.',
        ),
      },
    ],
    related: ['xl-workbooks-add', 'xl-workbook-close', 'xl-workbook-path'],
  },
  {
    id: 'xl-workbooks-add',
    name: 'Workbooks.Add',
    syntax: 'Workbooks.Add([Template])',
    scope: ['excel'],
    category: 'xl-books',
    subcategory: 'workbooks',
    description: L(
      'Crea una nuova cartella (come File → Nuovo). Template xlWBATWorksheet crea un solo foglio.',
      'Creates a new workbook (File → New). Template xlWBATWorksheet creates a single sheet.',
      'Crea un libro nuevo (Archivo → Nuevo). Template xlWBATWorksheet crea una sola hoja.',
      'Crée un nouveau classeur (Fichier → Nouveau). Template xlWBATWorksheet crée une seule feuille.',
      'Erstellt eine neue Mappe (Datei → Neu). Template xlWBATWorksheet erzeugt ein einzelnes Blatt.',
    ),
    params: [
      {
        name: 'Template',
        optional: true,
        description: L(
          'Percorso modello, o xlWBATWorksheet (-4167) per un foglio vuoto.',
          'Template path, or xlWBATWorksheet (-4167) for one blank sheet.',
          'Ruta de plantilla, o xlWBATWorksheet (-4167) para una hoja en blanco.',
          'Chemin de modèle, ou xlWBATWorksheet (-4167) pour une feuille vide.',
          'Vorlagenpfad oder xlWBATWorksheet (-4167) für ein leeres Blatt.',
        ),
      },
    ],
    examples: [
      {
        title: L('Nuova cartella a un foglio', 'New one-sheet workbook', 'Nuevo libro de una hoja', 'Nouveau classeur à une feuille', 'Neue Einblatt-Mappe'),
        code: `Sub DemoWorkbooksAdd()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Range("A1").Value = "new book"
    Debug.Print wb.Name & " sheets=" & wb.Worksheets.Count
    wb.Close SaveChanges:=False
End Sub`,
        result: L(
          'Nasce BookN con 1 foglio, A1 = new book, poi si chiude senza salvare. Immediate Window: BookN sheets=1.',
          'BookN appears with 1 sheet, A1 = new book, then closes unsaved. Immediate Window: BookN sheets=1.',
          'Aparece BookN con 1 hoja, A1 = new book, luego se cierra sin guardar. Immediate Window: BookN sheets=1.',
          'BookN apparaît avec 1 feuille, A1 = new book, puis se ferme sans enregistrer. Immediate Window : BookN sheets=1.',
          'BookN erscheint mit 1 Blatt, A1 = new book, dann ungespeichert geschlossen. Direktbereich: BookN sheets=1.',
        ),
      },
    ],
    related: ['xl-workbooks-open', 'xl-workbooks-count', 'xl-workbook-close'],
  },
  {
    id: 'xl-workbook-close',
    name: 'Workbook.Close',
    syntax: 'Workbook.Close [SaveChanges], [Filename], [RouteWorkbook]',
    scope: ['excel'],
    category: 'xl-books',
    subcategory: 'workbooks',
    description: L(
      'Chiude una cartella. SaveChanges True/False evita il prompt; omettilo per chiedere all’utente.',
      'Closes a workbook. SaveChanges True/False skips the prompt; omit it to ask the user.',
      'Cierra un libro. SaveChanges True/False evita el aviso; omítelo para preguntar.',
      'Ferme un classeur. SaveChanges True/False évite l’invite ; omettez-le pour demander.',
      'Schließt eine Mappe. SaveChanges True/False überspringt die Abfrage; weglassen fragt den Benutzer.',
    ),
    params: [
      { name: 'SaveChanges', optional: true, description: L('True salva; False scarta; omesso = chiedi.', 'True saves; False discards; omitted = ask.', 'True guarda; False descarta; omitido = preguntar.', 'True enregistre ; False ignore ; omis = demander.', 'True speichert; False verwirft; weggelassen = fragen.') },
    ],
    examples: [
      {
        title: L('Chiudi una cartella temporanea', 'Close a temporary workbook', 'Cerrar un libro temporal', 'Fermer un classeur temporaire', 'Temporäre Mappe schließen'),
        code: `Sub DemoWorkbookClose()
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Range("A1").Value = "temp"
    Debug.Print "Before=" & Workbooks.Count
    wb.Close SaveChanges:=False
    Debug.Print "After=" & Workbooks.Count
End Sub`,
        result: L(
          'Workbooks.Count scende di 1. Nessun dialogo di salvataggio. Immediate Window: Before=n After=n-1.',
          'Workbooks.Count drops by 1. No save dialog. Immediate Window: Before=n After=n-1.',
          'Workbooks.Count baja 1. Sin diálogo de guardar. Immediate Window: Before=n After=n-1.',
          'Workbooks.Count diminue de 1. Pas de dialogue d’enregistrement. Immediate Window : Before=n After=n-1.',
          'Workbooks.Count sinkt um 1. Kein Speichern-Dialog. Direktbereich: Before=n After=n-1.',
        ),
      },
    ],
    related: ['xl-workbook-save', 'xl-displayalerts', 'xl-workbooks-add'],
  },
  {
    id: 'xl-workbook-save',
    name: 'Workbook.Save / SaveAs',
    syntax: 'Workbook.Save  |  Workbook.SaveAs Filename, [FileFormat], [Password], ...',
    scope: ['excel'],
    category: 'xl-books',
    subcategory: 'workbooks',
    description: L(
      'Save scrive sul percorso corrente. SaveAs crea un nuovo file (xlOpenXMLWorkbook = 51 per .xlsx; xlOpenXMLWorkbookMacroEnabled = 52 per .xlsm).',
      'Save writes to the current path. SaveAs creates a new file (xlOpenXMLWorkbook = 51 for .xlsx; xlOpenXMLWorkbookMacroEnabled = 52 for .xlsm).',
      'Save escribe en la ruta actual. SaveAs crea un archivo nuevo (51 = .xlsx; 52 = .xlsm).',
      'Save écrit sur le chemin courant. SaveAs crée un nouveau fichier (51 = .xlsx ; 52 = .xlsm).',
      'Save schreibt auf den aktuellen Pfad. SaveAs erzeugt eine neue Datei (51 = .xlsx; 52 = .xlsm).',
    ),
    params: [
      { name: 'Filename', optional: true, description: L('Percorso di destinazione per SaveAs.', 'Destination path for SaveAs.', 'Ruta de destino para SaveAs.', 'Chemin de destination pour SaveAs.', 'Zielpfad für SaveAs.') },
      { name: 'FileFormat', optional: true, description: L('51 .xlsx, 52 .xlsm, 56 .xls.', '51 .xlsx, 52 .xlsm, 56 .xls.', '51 .xlsx, 52 .xlsm, 56 .xls.', '51 .xlsx, 52 .xlsm, 56 .xls.', '51 .xlsx, 52 .xlsm, 56 .xls.') },
    ],
    examples: [
      {
        title: L('SaveAs di una cartella nuova', 'SaveAs a brand-new workbook', 'SaveAs de un libro nuevo', 'SaveAs d’un nouveau classeur', 'SaveAs einer neuen Mappe'),
        code: `Sub DemoSave()
    Dim wb As Workbook, p As String
    p = Environ("TEMP") & "\\xl-demo.xlsx"
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Range("A1").Value = "saved"
    wb.SaveAs Filename:=p, FileFormat:=xlOpenXMLWorkbook
    Debug.Print "Saved " & wb.FullName
    wb.Close SaveChanges:=False
End Sub`,
        result: L(
          'Crea %TEMP%\\xl-demo.xlsx con A1 = saved. Immediate Window: Saved C:\\…\\xl-demo.xlsx.',
          'Creates %TEMP%\\xl-demo.xlsx with A1 = saved. Immediate Window: Saved C:\\…\\xl-demo.xlsx.',
          'Crea %TEMP%\\xl-demo.xlsx con A1 = saved. Immediate Window: Saved C:\\…\\xl-demo.xlsx.',
          'Crée %TEMP%\\xl-demo.xlsx avec A1 = saved. Immediate Window : Saved C:\\…\\xl-demo.xlsx.',
          'Erzeugt %TEMP%\\xl-demo.xlsx mit A1 = saved. Direktbereich: Saved C:\\…\\xl-demo.xlsx.',
        ),
      },
    ],
    notes: L(
      'ThisWorkbook.Save sul file che contiene il codice. Un .xlsx salvato da un .xlsm perde i moduli se usi formato 51.',
      'ThisWorkbook.Save writes the file that contains the code. Saving a .xlsm as format 51 strips modules.',
      'ThisWorkbook.Save escribe el archivo que contiene el código. Guardar .xlsm como 51 elimina módulos.',
      'ThisWorkbook.Save écrit le fichier qui contient le code. En 51, un .xlsm perd ses modules.',
      'ThisWorkbook.Save schreibt die Datei mit dem Code. Format 51 entfernt Module aus einem .xlsm.',
    ),
    related: ['xl-workbook-close', 'xl-workbook-path', 'xl-this-active-workbook'],
  },
  {
    id: 'xl-this-active-workbook',
    name: 'ThisWorkbook / ActiveWorkbook',
    syntax: 'ThisWorkbook  |  ActiveWorkbook',
    scope: ['excel'],
    category: 'xl-books',
    subcategory: 'workbooks',
    description: L(
      'ThisWorkbook è sempre la cartella che contiene questo codice. ActiveWorkbook è la cartella in primo piano (può essere un’altra).',
      'ThisWorkbook is always the workbook that contains this code. ActiveWorkbook is the foreground workbook (it can be another one).',
      'ThisWorkbook es siempre el libro que contiene este código. ActiveWorkbook es el libro en primer plano.',
      'ThisWorkbook est toujours le classeur qui contient ce code. ActiveWorkbook est le classeur au premier plan.',
      'ThisWorkbook ist immer die Mappe mit diesem Code. ActiveWorkbook ist die Vordergrund-Mappe.',
    ),
    examples: [
      {
        title: L('Confronta i due riferimenti', 'Compare the two references', 'Comparar las dos referencias', 'Comparer les deux références', 'Die zwei Bezüge vergleichen'),
        code: `Sub DemoThisVsActive()
    Range("A1").Value = ThisWorkbook.Name
    Range("A2").Value = ActiveWorkbook.Name
    Range("A3").Value = (ThisWorkbook Is ActiveWorkbook)
    Debug.Print "same=" & Range("A3").Value
End Sub`,
        result: L(
          'A1 = nome del file con la macro. A2 = cartella attiva. A3 = True se coincidono. Immediate Window: same=True/False.',
          'A1 = name of the file with the macro. A2 = active workbook. A3 = True if they are the same. Immediate Window: same=True/False.',
          'A1 = nombre del archivo con la macro. A2 = libro activo. A3 = True si coinciden. Immediate Window: same=True/False.',
          'A1 = nom du fichier de la macro. A2 = classeur actif. A3 = True s’ils sont identiques. Immediate Window : same=True/False.',
          'A1 = Name der Datei mit dem Makro. A2 = aktive Mappe. A3 = True wenn gleich. Direktbereich: same=True/False.',
        ),
      },
    ],
    notes: L(
      'Nelle add-in ThisWorkbook è l’add-in, ActiveWorkbook è il file dell’utente: non confonderli.',
      'In add-ins ThisWorkbook is the add-in and ActiveWorkbook is the user’s file — do not mix them up.',
      'En complementos ThisWorkbook es el complemento y ActiveWorkbook el archivo del usuario.',
      'Dans un complément ThisWorkbook est le complément, ActiveWorkbook le fichier utilisateur.',
      'In Add-Ins ist ThisWorkbook das Add-In, ActiveWorkbook die Benutzerdatei.',
    ),
    related: ['xl-workbook-path', 'xl-workbooks-count', 'xl-workbooks-open'],
  },
  {
    id: 'xl-workbook-path',
    name: 'Workbook.Path / Name / FullName',
    syntax: 'Workbook.Path  |  Workbook.Name  |  Workbook.FullName',
    scope: ['excel'],
    category: 'xl-books',
    subcategory: 'workbooks',
    description: L(
      'Name è il file (Book1.xlsx). Path è la cartella senza slash finale. FullName = Path & "\\" & Name. Path è "" se non ancora salvata.',
      'Name is the file (Book1.xlsx). Path is the folder with no trailing slash. FullName = Path & "\\" & Name. Path is "" until saved.',
      'Name es el archivo (Book1.xlsx). Path es la carpeta sin barra final. FullName = Path & "\\" & Name. Path es "" si no se ha guardado.',
      'Name est le fichier (Book1.xlsx). Path est le dossier sans slash final. FullName = Path & "\\" & Name. Path vaut "" tant que non enregistré.',
      'Name ist die Datei (Book1.xlsx). Path ist der Ordner ohne abschließenden Slash. FullName = Path & "\\" & Name. Path ist "" bis zum Speichern.',
    ),
    examples: [
      {
        title: L('Identità della cartella del codice', 'Identity of the code workbook', 'Identidad del libro del código', 'Identité du classeur du code', 'Identität der Code-Mappe'),
        code: `Sub DemoWorkbookPath()
    Range("A1").Value = ThisWorkbook.Name
    Range("A2").Value = ThisWorkbook.Path
    Range("A3").Value = ThisWorkbook.FullName
    Debug.Print "Name=" & ThisWorkbook.Name
    Debug.Print "unsaved=" & (ThisWorkbook.Path = "")
End Sub`,
        result: L(
          'A1 = Book1.xlsm (o simile). A2 = cartella o vuoto. A3 = percorso intero. Immediate Window ripete Name e unsaved.',
          'A1 = Book1.xlsm (or similar). A2 = folder or blank. A3 = full path. Immediate Window repeats Name and unsaved.',
          'A1 = Book1.xlsm (o similar). A2 = carpeta o vacío. A3 = ruta completa. Immediate Window repite Name y unsaved.',
          'A1 = Book1.xlsm (ou similaire). A2 = dossier ou vide. A3 = chemin complet. Immediate Window répète Name et unsaved.',
          'A1 = Book1.xlsm (oder ähnlich). A2 = Ordner oder leer. A3 = voller Pfad. Direktbereich wiederholt Name und unsaved.',
        ),
      },
    ],
    related: ['xl-workbook-save', 'xl-this-active-workbook', 'xl-workbooks-open'],
  },
  {
    id: 'xl-workbooks-count',
    name: 'Workbooks.Count',
    syntax: 'Workbooks.Count',
    scope: ['excel'],
    category: 'xl-books',
    subcategory: 'workbooks',
    description: L(
      'Numero di cartelle aperte nella sessione, inclusi add-in nascosti se caricati come Workbook.',
      'Number of open workbooks in the session, including hidden add-ins loaded as Workbook objects.',
      'Número de libros abiertos en la sesión, incluidos complementos ocultos cargados como Workbook.',
      'Nombre de classeurs ouverts, y compris les compléments cachés chargés comme Workbook.',
      'Anzahl offener Mappen in der Sitzung, inklusive als Workbook geladener Add-Ins.',
    ),
    examples: [
      {
        title: L('Conta prima e dopo Add', 'Count before and after Add', 'Contar antes y después de Add', 'Compter avant et après Add', 'Zählen vor und nach Add'),
        code: `Sub DemoWorkbooksCount()
    Dim n As Long, wb As Workbook
    n = Workbooks.Count
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Range("A1").Value = n
    Range("A2").Value = Workbooks.Count
    Debug.Print "was " & n & " now " & Workbooks.Count
    wb.Close SaveChanges:=False
End Sub`,
        result: L(
          'A1 = conteggio iniziale, A2 = iniziale+1. Immediate Window: was n now n+1. La cartella extra si chiude.',
          'A1 = starting count, A2 = start+1. Immediate Window: was n now n+1. The extra workbook closes.',
          'A1 = recuento inicial, A2 = inicial+1. Immediate Window: was n now n+1. El libro extra se cierra.',
          'A1 = compte initial, A2 = initial+1. Immediate Window : was n now n+1. Le classeur extra se ferme.',
          'A1 = Startzahl, A2 = Start+1. Direktbereich: was n now n+1. Die Extra-Mappe schließt sich.',
        ),
      },
    ],
    related: ['xl-workbooks-add', 'xl-this-active-workbook', 'xl-workbook-close'],
  },
  {
    id: 'xl-worksheets-add',
    name: 'Worksheets.Add',
    syntax: 'Worksheets.Add([Before], [After], [Count], [Type])',
    scope: ['excel'],
    category: 'xl-books',
    subcategory: 'sheets',
    description: L(
      'Inserisce uno o più fogli. Senza Before/After il nuovo foglio va prima del foglio attivo. Restituisce il Worksheet (o il primo se Count > 1).',
      'Inserts one or more worksheets. With no Before/After the new sheet goes before the active sheet. Returns the Worksheet (or the first if Count > 1).',
      'Inserta una o más hojas. Sin Before/After la nueva hoja va antes de la activa. Devuelve el Worksheet.',
      'Insère une ou plusieurs feuilles. Sans Before/After, la nouvelle feuille va avant la feuille active. Renvoie le Worksheet.',
      'Fügt ein oder mehrere Blätter ein. Ohne Before/After landet das neue Blatt vor dem aktiven. Liefert das Worksheet.',
    ),
    params: [
      { name: 'Before', optional: true, description: L('Foglio davanti al quale inserire.', 'Sheet to insert before.', 'Hoja delante de la cual insertar.', 'Feuille devant laquelle insérer.', 'Blatt, vor dem eingefügt wird.') },
      { name: 'After', optional: true, description: L('Foglio dopo il quale inserire.', 'Sheet to insert after.', 'Hoja detrás de la cual insertar.', 'Feuille après laquelle insérer.', 'Blatt, hinter dem eingefügt wird.') },
      { name: 'Count', optional: true, description: L('Quanti fogli creare (default 1).', 'How many sheets to create (default 1).', 'Cuántas hojas crear (predeterminado 1).', 'Combien de feuilles (défaut 1).', 'Wie viele Blätter (Standard 1).') },
    ],
    examples: [
      {
        title: L('Aggiungi un foglio in coda', 'Add a sheet at the end', 'Añadir una hoja al final', 'Ajouter une feuille à la fin', 'Blatt am Ende anfügen'),
        code: `Sub DemoWorksheetsAdd()
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = "Tail"
    ws.Range("A1").Value = "last"
    Debug.Print ws.Name & " index=" & ws.Index
End Sub`,
        result: L(
          'Nuovo foglio Tail in ultima posizione, A1 = last. Immediate Window: Tail index=N.',
          'New sheet Tail in last position, A1 = last. Immediate Window: Tail index=N.',
          'Nueva hoja Tail en última posición, A1 = last. Immediate Window: Tail index=N.',
          'Nouvelle feuille Tail en dernière position, A1 = last. Immediate Window : Tail index=N.',
          'Neues Blatt Tail an letzter Position, A1 = last. Direktbereich: Tail index=N.',
        ),
      },
    ],
    related: ['xl-worksheet-delete', 'xl-worksheet-name', 'xl-sheets-vs-worksheets'],
  },
  {
    id: 'xl-worksheet-delete',
    name: 'Worksheet.Delete',
    syntax: 'Worksheet.Delete',
    scope: ['excel'],
    category: 'xl-books',
    subcategory: 'sheets',
    description: L(
      'Elimina il foglio. Excel chiede conferma salvo DisplayAlerts = False. Non puoi eliminare l’ultimo foglio visibile della cartella.',
      'Deletes the sheet. Excel prompts unless DisplayAlerts = False. You cannot delete the last visible sheet in the workbook.',
      'Elimina la hoja. Excel pide confirmación salvo DisplayAlerts = False. No puedes borrar la última hoja visible.',
      'Supprime la feuille. Excel demande confirmation sauf DisplayAlerts = False. Impossible de supprimer la dernière feuille visible.',
      'Löscht das Blatt. Excel fragt nach, außer DisplayAlerts = False. Das letzte sichtbare Blatt kann nicht gelöscht werden.',
    ),
    examples: [
      {
        title: L('Crea e elimina TempDel', 'Create and delete TempDel', 'Crear y eliminar TempDel', 'Créer et supprimer TempDel', 'TempDel erstellen und löschen'),
        code: `Sub DemoWorksheetDelete()
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = "TempDel"
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    Debug.Print "TempDel exists=" & SheetExists2("TempDel")
End Sub

Function SheetExists2(n As String) As Boolean
    Dim s As Worksheet
    On Error Resume Next
    Set s = Worksheets(n)
    SheetExists2 = Not s Is Nothing
End Function`,
        result: L(
          'TempDel sparisce senza dialogo. Immediate Window: TempDel exists=False.',
          'TempDel disappears with no dialog. Immediate Window: TempDel exists=False.',
          'TempDel desaparece sin diálogo. Immediate Window: TempDel exists=False.',
          'TempDel disparaît sans dialogue. Immediate Window : TempDel exists=False.',
          'TempDel verschwindet ohne Dialog. Direktbereich: TempDel exists=False.',
        ),
      },
    ],
    related: ['xl-displayalerts', 'xl-worksheets-add', 'xl-worksheet-visible'],
  },
  {
    id: 'xl-worksheet-copy-move',
    name: 'Worksheet.Copy / Move',
    syntax: 'Worksheet.Copy [Before], [After]  |  Worksheet.Move [Before], [After]',
    scope: ['excel'],
    category: 'xl-books',
    subcategory: 'sheets',
    description: L(
      'Copy duplica il foglio (senza argomenti apre una nuova cartella). Move lo riposiziona nella stessa o in un’altra cartella.',
      'Copy duplicates the sheet (no arguments opens a new workbook). Move repositions it in the same or another workbook.',
      'Copy duplica la hoja (sin argumentos abre un libro nuevo). Move la recoloca en el mismo u otro libro.',
      'Copy duplique la feuille (sans arguments ouvre un nouveau classeur). Move la repositionne.',
      'Copy dupliziert das Blatt (ohne Argumente öffnet eine neue Mappe). Move verschiebt es.',
    ),
    params: [
      { name: 'Before', optional: true, description: L('Foglio di riferimento davanti.', 'Reference sheet to place before.', 'Hoja de referencia delante.', 'Feuille de référence devant.', 'Referenzblatt davor.') },
      { name: 'After', optional: true, description: L('Foglio di riferimento dietro.', 'Reference sheet to place after.', 'Hoja de referencia detrás.', 'Feuille de référence derrière.', 'Referenzblatt danach.') },
    ],
    examples: [
      {
        title: L('Copia il foglio accanto e poi spostalo', 'Copy the sheet next door then move it', 'Copiar la hoja al lado y luego moverla', 'Copier la feuille à côté puis la déplacer', 'Blatt daneben kopieren und verschieben'),
        code: `Sub DemoSheetCopyMove()
    Dim src As Worksheet
    Set src = ActiveSheet
    src.Range("A1").Value = "orig"
    src.Copy After:=src
    ActiveSheet.Name = "CopyOf" & Left(src.Name, 20)
    ActiveSheet.Range("A2").Value = "copy"
    ActiveSheet.Move Before:=Worksheets(1)
    Debug.Print ActiveSheet.Name & " now index=" & ActiveSheet.Index
End Sub`,
        result: L(
          'Nuovo foglio CopyOf… con A1=orig, A2=copy, spostato in posizione 1. Immediate Window: nome index=1.',
          'New sheet CopyOf… with A1=orig, A2=copy, moved to position 1. Immediate Window: name index=1.',
          'Nueva hoja CopyOf… con A1=orig, A2=copy, movida a la posición 1. Immediate Window: nombre index=1.',
          'Nouvelle feuille CopyOf… avec A1=orig, A2=copy, déplacée en position 1. Immediate Window : nom index=1.',
          'Neues Blatt CopyOf… mit A1=orig, A2=copy, an Position 1. Direktbereich: Name index=1.',
        ),
      },
    ],
    related: ['xl-worksheets-add', 'xl-worksheet-name', 'xl-workbooks-add'],
  },
  {
    id: 'xl-worksheet-name',
    name: 'Worksheet.Name',
    syntax: 'Worksheet.Name = String',
    scope: ['excel'],
    category: 'xl-books',
    subcategory: 'sheets',
    description: L(
      'Legge o imposta la linguetta del foglio (max 31 caratteri). Vietati : \\ / ? * [ ]. Il nome deve essere unico nella cartella.',
      'Reads or sets the sheet tab (max 31 characters). Forbidden: : \\ / ? * [ ]. The name must be unique in the workbook.',
      'Lee o asigna la pestaña (máx. 31). Prohibidos : \\ / ? * [ ]. El nombre debe ser único.',
      'Lit ou définit l’onglet (max 31). Interdits : \\ / ? * [ ]. Le nom doit être unique.',
      'Liest oder setzt den Blattreiter (max. 31). Verboten: : \\ / ? * [ ]. Der Name muss eindeutig sein.',
    ),
    examples: [
      {
        title: L('Rinomina il foglio attivo', 'Rename the active sheet', 'Renombrar la hoja activa', 'Renommer la feuille active', 'Aktives Blatt umbenennen'),
        code: `Sub DemoWorksheetName()
    Dim oldName As String
    oldName = ActiveSheet.Name
    ActiveSheet.Name = "Sales"
    Range("A1").Value = ActiveSheet.Name
    Debug.Print "was " & oldName & " now " & ActiveSheet.Name
End Sub`,
        result: L(
          'Linguetta = Sales, A1 = Sales. Immediate Window: was <vecchio> now Sales. Errore 1004 se Sales esiste già.',
          'Tab = Sales, A1 = Sales. Immediate Window: was <old> now Sales. Error 1004 if Sales already exists.',
          'Pestaña = Sales, A1 = Sales. Immediate Window: was <viejo> now Sales. Error 1004 si Sales ya existe.',
          'Onglet = Sales, A1 = Sales. Immediate Window : was <ancien> now Sales. Erreur 1004 si Sales existe.',
          'Reiter = Sales, A1 = Sales. Direktbereich: was <alt> now Sales. Fehler 1004, wenn Sales schon existiert.',
        ),
      },
    ],
    related: ['xl-worksheets-add', 'xl-sheets-vs-worksheets', 'xl-worksheet-activate'],
  },
  {
    id: 'xl-worksheet-visible',
    name: 'Worksheet.Visible',
    syntax: 'Worksheet.Visible = xlSheetVisible | xlSheetHidden | xlSheetVeryHidden',
    scope: ['excel'],
    category: 'xl-books',
    subcategory: 'sheets',
    description: L(
      'xlSheetVisible (-1) mostra la linguetta. xlSheetHidden (0) nasconde (svelabile da UI). xlSheetVeryHidden (2) nasconde anche da Formato → Nascondi: solo VBA lo riporta.',
      'xlSheetVisible (-1) shows the tab. xlSheetHidden (0) hides it (UI can unhide). xlSheetVeryHidden (2) hides it from Format → Hide as well: only VBA restores it.',
      'xlSheetVisible (-1) muestra la pestaña. xlSheetHidden (0) la oculta (la IU puede mostrar). xlSheetVeryHidden (2) solo VBA la restaura.',
      'xlSheetVisible (-1) montre l’onglet. xlSheetHidden (0) le cache (l’UI peut le réafficher). xlSheetVeryHidden (2) : seul VBA le rétablit.',
      'xlSheetVisible (-1) zeigt den Reiter. xlSheetHidden (0) blendet aus (UI kann einblenden). xlSheetVeryHidden (2): nur VBA holt es zurück.',
    ),
    examples: [
      {
        title: L('Nascondi e svela via VBA', 'Hide and unhide via VBA', 'Ocultar y mostrar vía VBA', 'Masquer et réafficher via VBA', 'Aus- und einblenden per VBA'),
        code: `Sub DemoWorksheetVisible()
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = "Secret"
    ws.Visible = xlSheetVeryHidden
    Debug.Print "hidden=" & ws.Visible
    ws.Visible = xlSheetVisible
    ws.Range("A1").Value = "back"
    Debug.Print "visible=" & ws.Visible
End Sub`,
        result: L(
          'Secret sparisce dalle linguette (Visible=2) poi torna, A1 = back. Immediate Window: hidden=2 visible=-1.',
          'Secret vanishes from tabs (Visible=2) then returns, A1 = back. Immediate Window: hidden=2 visible=-1.',
          'Secret desaparece de las pestañas (Visible=2) y vuelve, A1 = back. Immediate Window: hidden=2 visible=-1.',
          'Secret disparaît des onglets (Visible=2) puis revient, A1 = back. Immediate Window : hidden=2 visible=-1.',
          'Secret verschwindet von den Reitern (Visible=2) und kommt zurück, A1 = back. Direktbereich: hidden=2 visible=-1.',
        ),
      },
    ],
    related: ['xl-worksheet-activate', 'xl-sheets-vs-worksheets', 'xl-worksheet-delete'],
  },
  {
    id: 'xl-worksheet-protect',
    name: 'Worksheet.Protect / Unprotect',
    syntax: 'Worksheet.Protect [Password], [DrawingObjects], [Contents], [Scenarios], [UserInterfaceOnly], ...  |  Worksheet.Unprotect [Password]',
    scope: ['excel'],
    category: 'xl-books',
    subcategory: 'sheets',
    description: L(
      'Blocca modifiche utente alle celle locked. UserInterfaceOnly:=True lascia scrivere le macro senza togliere la protezione (da reimpostare a ogni Open).',
      'Blocks user edits to locked cells. UserInterfaceOnly:=True lets macros write without unprotecting (must be set again on each Open).',
      'Bloquea ediciones de usuario en celdas locked. UserInterfaceOnly:=True deja escribir a las macros (hay que reponerlo en cada Open).',
      'Bloque les éditions utilisateur des cellules locked. UserInterfaceOnly:=True laisse les macros écrire (à rétablir à chaque Open).',
      'Sperrt Benutzeränderungen an locked Zellen. UserInterfaceOnly:=True lässt Makros schreiben (pro Open neu setzen).',
    ),
    params: [
      { name: 'Password', optional: true, description: L('Password (facoltativa ma consigliata).', 'Password (optional but recommended).', 'Contraseña (opcional pero recomendada).', 'Mot de passe (facultatif mais conseillé).', 'Kennwort (optional, empfohlen).') },
      { name: 'UserInterfaceOnly', optional: true, description: L('True: VBA può scrivere, l’utente no.', 'True: VBA may write, the user may not.', 'True: VBA puede escribir, el usuario no.', 'True : VBA peut écrire, pas l’utilisateur.', 'True: VBA darf schreiben, der Benutzer nicht.') },
    ],
    examples: [
      {
        title: L('Proteggi, scrivi da VBA, sproteggi', 'Protect, write from VBA, unprotect', 'Proteger, escribir desde VBA, desproteger', 'Protéger, écrire depuis VBA, déprotéger', 'Schützen, per VBA schreiben, Schutz aufheben'),
        code: `Sub DemoProtect()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Range("A1").Value = "locked demo"
    ws.Protect Password:="demo", UserInterfaceOnly:=True
    ws.Range("A2").Value = "macro ok"
    Debug.Print "ProtectContents=" & ws.ProtectContents
    ws.Unprotect Password:="demo"
    Debug.Print "after Unprotect=" & ws.ProtectContents
End Sub`,
        result: L(
          'A1 = locked demo, A2 = macro ok. Immediate Window: ProtectContents=True poi after Unprotect=False. L’utente non può editare mentre è protetto.',
          'A1 = locked demo, A2 = macro ok. Immediate Window: ProtectContents=True then after Unprotect=False. The user cannot edit while protected.',
          'A1 = locked demo, A2 = macro ok. Immediate Window: ProtectContents=True luego after Unprotect=False. El usuario no puede editar.',
          'A1 = locked demo, A2 = macro ok. Immediate Window : ProtectContents=True puis after Unprotect=False. L’utilisateur ne peut pas éditer.',
          'A1 = locked demo, A2 = macro ok. Direktbereich: ProtectContents=True, dann after Unprotect=False. Benutzer kann nicht editieren.',
        ),
      },
    ],
    related: ['xl-worksheet-name', 'xl-workbook-open-event', 'xl-range-value'],
  },
  {
    id: 'xl-worksheet-activate',
    name: 'Worksheet.Activate',
    syntax: 'Worksheet.Activate',
    scope: ['excel'],
    category: 'xl-books',
    subcategory: 'sheets',
    description: L(
      'Porta il foglio in primo piano (ActiveSheet). Un foglio VeryHidden va prima reso Visible. Preferisci riferimenti espliciti Worksheets("X").Range invece di Activate.',
      'Brings the sheet to the front (ActiveSheet). A VeryHidden sheet must be Visible first. Prefer explicit Worksheets("X").Range over Activate.',
      'Trae la hoja al frente (ActiveSheet). Una hoja VeryHidden debe ser Visible antes. Prefiere Worksheets("X").Range.',
      'Met la feuille au premier plan (ActiveSheet). Une feuille VeryHidden doit d’abord être Visible. Préférez Worksheets("X").Range.',
      'Holt das Blatt nach vorn (ActiveSheet). Ein VeryHidden-Blatt muss zuerst Visible sein. Lieber Worksheets("X").Range.',
    ),
    examples: [
      {
        title: L('Attiva il primo foglio e scrivi', 'Activate the first sheet and write', 'Activar la primera hoja y escribir', 'Activer la première feuille et écrire', 'Erstes Blatt aktivieren und schreiben'),
        code: `Sub DemoWorksheetActivate()
    Worksheets(1).Activate
    ActiveSheet.Range("A1").Value = "active"
    Debug.Print ActiveSheet.Name & " A1=" & Range("A1").Value
End Sub`,
        result: L(
          'Il primo foglio diventa attivo, A1 = active. Immediate Window: <nome> A1=active. La linguetta risulta selezionata.',
          'The first sheet becomes active, A1 = active. Immediate Window: <name> A1=active. The tab appears selected.',
          'La primera hoja queda activa, A1 = active. Immediate Window: <nombre> A1=active. La pestaña se ve seleccionada.',
          'La première feuille devient active, A1 = active. Immediate Window : <nom> A1=active. L’onglet apparaît sélectionné.',
          'Das erste Blatt wird aktiv, A1 = active. Direktbereich: <Name> A1=active. Der Reiter ist ausgewählt.',
        ),
      },
    ],
    related: ['xl-worksheet-visible', 'xl-this-active-workbook', 'xl-worksheet-activate-event'],
  },
  {
    id: 'xl-sheets-vs-worksheets',
    name: 'Sheets vs Worksheets',
    syntax: 'Sheets(index | name)  |  Worksheets(index | name)',
    scope: ['excel'],
    category: 'xl-books',
    subcategory: 'sheets',
    description: L(
      'Worksheets contiene solo fogli di calcolo. Sheets include anche grafici (Chart), macro Excel 4 e dialoghi. Sheets.Count è >= Worksheets.Count.',
      'Worksheets holds only worksheets. Sheets also includes chart sheets, Excel 4 macro sheets and dialogs. Sheets.Count is >= Worksheets.Count.',
      'Worksheets solo hojas de cálculo. Sheets incluye gráficos, macros Excel 4 y diálogos. Sheets.Count >= Worksheets.Count.',
      'Worksheets ne contient que les feuilles de calcul. Sheets inclut graphiques, macros Excel 4 et dialogues. Sheets.Count >= Worksheets.Count.',
      'Worksheets enthält nur Tabellenblätter. Sheets umfasst auch Diagrammblätter, Excel-4-Makros und Dialoge. Sheets.Count >= Worksheets.Count.',
    ),
    examples: [
      {
        title: L('Confronta i due conteggi', 'Compare the two counts', 'Comparar los dos recuentos', 'Comparer les deux comptes', 'Die zwei Zählungen vergleichen'),
        code: `Sub DemoSheetsVsWorksheets()
    Range("A1").Value = Worksheets.Count
    Range("A2").Value = Sheets.Count
    Range("A3").Value = TypeName(Sheets(1))
    Debug.Print "WS=" & Worksheets.Count & " Sheets=" & Sheets.Count
End Sub`,
        result: L(
          'A1 = n fogli di calcolo, A2 = n+eventuali chart. A3 = Worksheet o Chart. Immediate Window: WS=n Sheets=m.',
          'A1 = n worksheets, A2 = n plus any chart sheets. A3 = Worksheet or Chart. Immediate Window: WS=n Sheets=m.',
          'A1 = n hojas de cálculo, A2 = n más gráficos. A3 = Worksheet o Chart. Immediate Window: WS=n Sheets=m.',
          'A1 = n feuilles de calcul, A2 = n plus graphiques. A3 = Worksheet ou Chart. Immediate Window : WS=n Sheets=m.',
          'A1 = n Tabellenblätter, A2 = n plus Diagrammblätter. A3 = Worksheet oder Chart. Direktbereich: WS=n Sheets=m.',
        ),
      },
    ],
    notes: L(
      'Worksheets("Grafico1") fallisce se Grafico1 è un Chart sheet: usa Sheets("Grafico1").',
      'Worksheets("Chart1") fails if Chart1 is a chart sheet: use Sheets("Chart1").',
      'Worksheets("Gráfico1") falla si es una hoja de gráfico: usa Sheets("Gráfico1").',
      'Worksheets("Graphique1") échoue si c’est une feuille graphique : utilisez Sheets.',
      'Worksheets("Diagramm1") schlägt fehl, wenn es ein Diagrammblatt ist: Sheets verwenden.',
    ),
    related: ['xl-worksheets-add', 'xl-worksheet-name', 'xl-chartobjects'],
  },
  {
    id: 'xl-range-value',
    name: 'Range.Value / Value2',
    syntax: 'Range.Value [ = Variant ]  |  Range.Value2 [ = Variant ]',
    scope: ['excel'],
    category: 'xl-range',
    subcategory: 'readwrite',
    description: L(
      'Value legge/scrive il valore della cella (Date e Currency restano tipizzati). Value2 restituisce lo stesso numero sottostante senza sottotipo Date/Currency: le date diventano Double (serial).',
      'Value reads/writes the cell value (Date and Currency stay typed). Value2 returns the same underlying number without Date/Currency subtypes: dates become Double (serial).',
      'Value lee/escribe el valor (Date y Currency tipados). Value2 devuelve el número sin subtipo Date/Currency: las fechas son Double.',
      'Value lit/écrit la valeur (Date et Currency typés). Value2 renvoie le nombre sans sous-type Date/Currency : les dates deviennent Double.',
      'Value liest/schreibt den Wert (Date und Currency typisiert). Value2 liefert die Zahl ohne Date/Currency-Subtyp: Daten werden Double.',
    ),
    examples: [
      {
        title: L('Scrivi numeri e confronta Value/Value2 su una data', 'Write numbers and compare Value/Value2 on a date', 'Escribir números y comparar Value/Value2 en una fecha', 'Écrire des nombres et comparer Value/Value2 sur une date', 'Zahlen schreiben und Value/Value2 bei einem Datum vergleichen'),
        code: `Sub DemoRangeValue()
    Range("A1").Value = 42
    Range("A2").Value = DateSerial(2026, 9, 5)
    Range("B1").Value = TypeName(Range("A2").Value)
    Range("B2").Value = TypeName(Range("A2").Value2)
    Debug.Print "A1=" & Range("A1").Value
    Debug.Print "Value=" & Range("B1").Value & " Value2=" & Range("B2").Value
End Sub`,
        result: L(
          'A1 = 42, A2 = 05/09/2026. B1 = Date, B2 = Double. Immediate Window: A1=42 e i due TypeName.',
          'A1 = 42, A2 = 2026-09-05. B1 = Date, B2 = Double. Immediate Window: A1=42 and both TypeName values.',
          'A1 = 42, A2 = 05/09/2026. B1 = Date, B2 = Double. Immediate Window: A1=42 y ambos TypeName.',
          'A1 = 42, A2 = 05/09/2026. B1 = Date, B2 = Double. Immediate Window : A1=42 et les deux TypeName.',
          'A1 = 42, A2 = 05.09.2026. B1 = Date, B2 = Double. Direktbereich: A1=42 und beide TypeName.',
        ),
      },
    ],
    notes: L(
      'Assegnare un array a Range.Value scrive un blocco in un colpo solo, molto più veloce del loop cella-per-cella.',
      'Assigning an array to Range.Value writes a block in one shot, much faster than a cell-by-cell loop.',
      'Asignar un array a Range.Value escribe un bloque de una vez, más rápido que el bucle celda a celda.',
      'Affecter un tableau à Range.Value écrit un bloc d’un coup, bien plus rapide qu’une boucle.',
      'Ein Array an Range.Value zuweisen schreibt einen Block auf einmal, viel schneller als Zellschleifen.',
    ),
    related: ['xl-range-formula', 'xl-cells-range-activecell', 'xl-range-text-format'],
  },
  {
    id: 'xl-range-formula',
    name: 'Range.Formula / FormulaR1C1',
    syntax: 'Range.Formula = String  |  Range.FormulaR1C1 = String',
    scope: ['excel'],
    category: 'xl-range',
    subcategory: 'readwrite',
    description: L(
      'Formula usa riferimenti A1 ("=A1+B1"). FormulaR1C1 usa R1C1 ("=RC[-1]+R[-1]C"), ideale quando costruisci formule in loop. Entrambe accettano anche costanti.',
      'Formula uses A1 references ("=A1+B1"). FormulaR1C1 uses R1C1 ("=RC[-1]+R[-1]C"), ideal when building formulas in loops. Both also accept constants.',
      'Formula usa referencias A1 ("=A1+B1"). FormulaR1C1 usa R1C1 ("=RC[-1]+R[-1]C"), ideal en bucles. Ambas aceptan constantes.',
      'Formula utilise A1 ("=A1+B1"). FormulaR1C1 utilise R1C1 ("=RC[-1]+R[-1]C"), idéal en boucle. Les deux acceptent des constantes.',
      'Formula nutzt A1 ("=A1+B1"). FormulaR1C1 nutzt R1C1 ("=RC[-1]+R[-1]C"), ideal in Schleifen. Beide akzeptieren Konstanten.',
    ),
    examples: [
      {
        title: L('Stessa somma in A1 e in R1C1', 'Same sum in A1 and in R1C1', 'La misma suma en A1 y en R1C1', 'La même somme en A1 et en R1C1', 'Dieselbe Summe in A1 und R1C1'),
        code: `Sub DemoFormula()
    Range("A1").Value = 3
    Range("B1").Value = 4
    Range("C1").Formula = "=A1+B1"
    Range("C2").FormulaR1C1 = "=R[-1]C[-2]+R[-1]C[-1]"
    Debug.Print Range("C1").Formula & " -> " & Range("C1").Value
    Debug.Print Range("C2").FormulaR1C1 & " -> " & Range("C2").Value
End Sub`,
        result: L(
          'C1 = 7 con formula =A1+B1. C2 = 7 con =R[-1]C[-2]+R[-1]C[-1]. Immediate Window stampa entrambe.',
          'C1 = 7 with formula =A1+B1. C2 = 7 with =R[-1]C[-2]+R[-1]C[-1]. Immediate Window prints both.',
          'C1 = 7 con fórmula =A1+B1. C2 = 7 con =R[-1]C[-2]+R[-1]C[-1]. Immediate Window imprime ambas.',
          'C1 = 7 avec formule =A1+B1. C2 = 7 avec =R[-1]C[-2]+R[-1]C[-1]. Immediate Window affiche les deux.',
          'C1 = 7 mit Formel =A1+B1. C2 = 7 mit =R[-1]C[-2]+R[-1]C[-1]. Direktbereich gibt beide aus.',
        ),
      },
    ],
    notes: L(
      'FormulaLocal usa i nomi funzione della lingua UI (=SOMMA). In VBA preferisci Formula / FormulaR1C1 in inglese.',
      'FormulaLocal uses UI-language function names (=SOMMA). In VBA prefer Formula / FormulaR1C1 in English.',
      'FormulaLocal usa nombres de la IU (=SUMA). En VBA prefiere Formula / FormulaR1C1 en inglés.',
      'FormulaLocal utilise les noms UI (=SOMME). En VBA préférez Formula / FormulaR1C1 en anglais.',
      'FormulaLocal nutzt UI-Namen (=SUMME). In VBA Formula / FormulaR1C1 auf Englisch bevorzugen.',
    ),
    related: ['xl-range-value', 'xl-range-calculate', 'xl-range-text-format'],
  },
  {
    id: 'xl-range-text-format',
    name: 'Range.Text / NumberFormat',
    syntax: 'Range.Text  |  Range.NumberFormat = String',
    scope: ['excel'],
    category: 'xl-range',
    subcategory: 'readwrite',
    description: L(
      'Text è la stringa visualizzata (sola lettura, dipende da formato e larghezza colonna). NumberFormat è il codice formato ("0.00", "yyyy-mm-dd", "@").',
      'Text is the displayed string (read-only, depends on format and column width). NumberFormat is the format code ("0.00", "yyyy-mm-dd", "@").',
      'Text es la cadena mostrada (solo lectura, depende del formato y el ancho). NumberFormat es el código ("0.00", "yyyy-mm-dd", "@").',
      'Text est la chaîne affichée (lecture seule, selon format et largeur). NumberFormat est le code ("0.00", "yyyy-mm-dd", "@").',
      'Text ist die angezeigte Zeichenfolge (nur lesen, abhängig von Format und Breite). NumberFormat ist der Code ("0.00", "yyyy-mm-dd", "@").',
    ),
    examples: [
      {
        title: L('Formatta un importo e leggi Text', 'Format an amount and read Text', 'Formatear un importe y leer Text', 'Mettre en forme un montant et lire Text', 'Betrag formatieren und Text lesen'),
        code: `Sub DemoTextFormat()
    Range("A1").Value = 1234.5
    Range("A1").NumberFormat = "#,##0.00"
    Range("B1").Value = Range("A1").Text
    Range("B2").Value = Range("A1").NumberFormat
    Debug.Print "Text=[" & Range("A1").Text & "] fmt=" & Range("A1").NumberFormat
End Sub`,
        result: L(
          'A1 mostra 1,234.50 (o 1.234,50). B1 = quella stringa, B2 = #,##0.00. Immediate Window ripete Text e formato.',
          'A1 shows 1,234.50 (or 1.234,50). B1 = that string, B2 = #,##0.00. Immediate Window repeats Text and format.',
          'A1 muestra 1,234.50 (o 1.234,50). B1 = esa cadena, B2 = #,##0.00. Immediate Window repite Text y formato.',
          'A1 affiche 1 234,50 (ou 1,234.50). B1 = cette chaîne, B2 = #,##0.00. Immediate Window répète Text et format.',
          'A1 zeigt 1.234,50 (oder 1,234.50). B1 = diese Zeichenfolge, B2 = #,##0.00. Direktbereich wiederholt Text und Format.',
        ),
      },
    ],
    notes: L(
      'Se la colonna è troppo stretta Text può essere "#####". Allarga o usa AutoFit prima di leggere Text.',
      'If the column is too narrow Text can be "#####". Widen or AutoFit before reading Text.',
      'Si la columna es estrecha Text puede ser "#####". Ensancha o AutoFit antes de leer Text.',
      'Si la colonne est trop étroite Text peut être "#####". Élargissez ou AutoFit avant de lire Text.',
      'Ist die Spalte zu schmal, kann Text "#####" sein. Vor dem Lesen verbreitern oder AutoFit.',
    ),
    related: ['xl-range-value', 'xl-range-autofit', 'xl-range-font'],
  },
  {
    id: 'xl-cells-range-activecell',
    name: 'Cells / Range("A1") / ActiveCell',
    syntax: 'Cells([RowIndex], [ColumnIndex])  |  Range(Address)  |  ActiveCell',
    scope: ['excel'],
    category: 'xl-range',
    subcategory: 'readwrite',
    description: L(
      'Tre modi per puntare celle. Cells(r, c) è 1-based e comodo nei loop. Range("A1") / Range("A1:B2") usa indirizzi A1. ActiveCell è la cella attiva del foglio attivo.',
      'Three ways to point at cells. Cells(r, c) is 1-based and handy in loops. Range("A1") / Range("A1:B2") uses A1 addresses. ActiveCell is the active cell on the active sheet.',
      'Tres formas de apuntar celdas. Cells(r, c) es base 1 e ideal en bucles. Range("A1") usa direcciones A1. ActiveCell es la celda activa.',
      'Trois façons de viser une cellule. Cells(r, c) est en base 1, pratique en boucle. Range("A1") utilise A1. ActiveCell est la cellule active.',
      'Drei Wege zu Zellen. Cells(r, c) ist 1-basiert und gut in Schleifen. Range("A1") nutzt A1-Adressen. ActiveCell ist die aktive Zelle.',
    ),
    params: [
      { name: 'RowIndex', optional: true, description: L('Riga 1-based per Cells.', '1-based row for Cells.', 'Fila base 1 para Cells.', 'Ligne base 1 pour Cells.', '1-basierte Zeile für Cells.') },
      { name: 'ColumnIndex', optional: true, description: L('Colonna 1-based o lettera ("C").', '1-based column or letter ("C").', 'Columna base 1 o letra ("C").', 'Colonne base 1 ou lettre ("C").', '1-basierte Spalte oder Buchstabe ("C").') },
    ],
    examples: [
      {
        title: L('Scrivi con tutti e tre gli stili', 'Write with all three styles', 'Escribir con los tres estilos', 'Écrire avec les trois styles', 'Mit allen drei Stilen schreiben'),
        code: `Sub DemoCellsRange()
    Cells(1, 1).Value = "via Cells"
    Range("B1").Value = "via Range"
    Range("C1").Select
    ActiveCell.Value = "via ActiveCell"
    Debug.Print Range("A1").Value & " | " & Range("B1").Value & " | " & Range("C1").Value
End Sub`,
        result: L(
          'A1 = via Cells, B1 = via Range, C1 = via ActiveCell (C1 selezionata). Immediate Window stampa le tre stringhe.',
          'A1 = via Cells, B1 = via Range, C1 = via ActiveCell (C1 selected). Immediate Window prints the three strings.',
          'A1 = via Cells, B1 = via Range, C1 = via ActiveCell (C1 seleccionada). Immediate Window imprime las tres.',
          'A1 = via Cells, B1 = via Range, C1 = via ActiveCell (C1 sélectionnée). Immediate Window affiche les trois.',
          'A1 = via Cells, B1 = via Range, C1 = via ActiveCell (C1 ausgewählt). Direktbereich gibt die drei aus.',
        ),
      },
    ],
    notes: L(
      'Cells senza argomenti è tutto il foglio. Qualifica sempre il foglio: Worksheets("Sales").Cells(1, 1).',
      'Cells with no arguments is the entire sheet. Always qualify the sheet: Worksheets("Sales").Cells(1, 1).',
      'Cells sin argumentos es toda la hoja. Califica siempre: Worksheets("Sales").Cells(1, 1).',
      'Cells sans arguments = toute la feuille. Qualifiez toujours : Worksheets("Sales").Cells(1, 1).',
      'Cells ohne Argumente ist das ganze Blatt. Immer qualifizieren: Worksheets("Sales").Cells(1, 1).',
    ),
    related: ['xl-range-value', 'xl-range-offset-resize', 'xl-range-address-row-col'],
  },
  {
    id: 'xl-range-offset-resize',
    name: 'Range.Offset / Resize',
    syntax: 'Range.Offset([RowOffset], [ColumnOffset])  |  Range.Resize([RowSize], [ColumnSize])',
    scope: ['excel'],
    category: 'xl-range',
    subcategory: 'navigate',
    description: L(
      'Offset sposta l’angolo rispetto al range (può essere negativo). Resize cambia altezza/larghezza restando ancorati all’angolo in alto a sinistra. Incollali: Offset(1,0).Resize(10,1).',
      'Offset moves the corner relative to the range (can be negative). Resize changes height/width while staying anchored at the top-left. Chain them: Offset(1,0).Resize(10,1).',
      'Offset mueve la esquina respecto al rango (puede ser negativo). Resize cambia alto/ancho anclado arriba-izquierda. Encadena: Offset(1,0).Resize(10,1).',
      'Offset déplace le coin (peut être négatif). Resize change hauteur/largeur en restant ancré en haut à gauche. Enchaînez : Offset(1,0).Resize(10,1).',
      'Offset verschiebt die Ecke relativ zum Bereich (kann negativ sein). Resize ändert Höhe/Breite, oben links verankert. Verketten: Offset(1,0).Resize(10,1).',
    ),
    params: [
      { name: 'RowOffset', optional: true, description: L('Righe da spostare (0 = stessa riga).', 'Rows to move (0 = same row).', 'Filas a mover (0 = misma fila).', 'Lignes à déplacer (0 = même ligne).', 'Zu verschiebende Zeilen (0 = dieselbe).') },
      { name: 'ColumnOffset', optional: true, description: L('Colonne da spostare.', 'Columns to move.', 'Columnas a mover.', 'Colonnes à déplacer.', 'Zu verschiebende Spalten.') },
      { name: 'RowSize', optional: true, description: L('Nuovo numero di righe (>= 1).', 'New row count (>= 1).', 'Nuevo número de filas (>= 1).', 'Nouveau nombre de lignes (>= 1).', 'Neue Zeilenzahl (>= 1).') },
      { name: 'ColumnSize', optional: true, description: L('Nuovo numero di colonne (>= 1).', 'New column count (>= 1).', 'Nuevo número de columnas (>= 1).', 'Nouveau nombre de colonnes (>= 1).', 'Neue Spaltenzahl (>= 1).') },
    ],
    examples: [
      {
        title: L('Dal titolo alla colonna dati sotto', 'From the header to the data column below', 'Del título a la columna de datos debajo', 'Du titre à la colonne de données dessous', 'Vom Titel zur Datenspalte darunter'),
        code: `Sub DemoOffsetResize()
    Range("B2").Value = "Qty"
    Range("B2").Offset(1, 0).Resize(3, 1).Value = Application.Transpose(Array(5, 8, 2))
    Debug.Print Range("B2").Offset(1, 0).Resize(3, 1).Address & "=" & _
        Range("B3").Value & "," & Range("B4").Value & "," & Range("B5").Value
End Sub`,
        result: L(
          'B2 = Qty, B3:B5 = 5,8,2. Immediate Window: $B$3:$B$5=5,8,2.',
          'B2 = Qty, B3:B5 = 5,8,2. Immediate Window: $B$3:$B$5=5,8,2.',
          'B2 = Qty, B3:B5 = 5,8,2. Immediate Window: $B$3:$B$5=5,8,2.',
          'B2 = Qty, B3:B5 = 5,8,2. Immediate Window : $B$3:$B$5=5,8,2.',
          'B2 = Qty, B3:B5 = 5,8,2. Direktbereich: $B$3:$B$5=5,8,2.',
        ),
      },
    ],
    related: ['xl-range-address-row-col', 'xl-cells-range-activecell', 'xl-range-end'],
  },
  {
    id: 'xl-range-address-row-col',
    name: 'Range.Address / Row / Column',
    syntax: 'Range.Address([RowAbsolute], [ColumnAbsolute], [ReferenceStyle])  |  Range.Row  |  Range.Column',
    scope: ['excel'],
    category: 'xl-range',
    subcategory: 'navigate',
    description: L(
      'Address restituisce l’indirizzo ($C$5 o C5). Row e Column sono gli indici 1-based dell’angolo in alto a sinistra (C5 → 5 e 3).',
      'Address returns the address ($C$5 or C5). Row and Column are the 1-based indexes of the top-left corner (C5 → 5 and 3).',
      'Address devuelve la dirección ($C$5 o C5). Row y Column son los índices base 1 de la esquina superior izquierda (C5 → 5 y 3).',
      'Address renvoie l’adresse ($C$5 ou C5). Row et Column sont les index base 1 du coin haut-gauche (C5 → 5 et 3).',
      'Address liefert die Adresse ($C$5 oder C5). Row und Column sind die 1-basierten Indizes der linken oberen Ecke (C5 → 5 und 3).',
    ),
    examples: [
      {
        title: L('Indirizzo assoluto e relativo', 'Absolute and relative address', 'Dirección absoluta y relativa', 'Adresse absolue et relative', 'Absolute und relative Adresse'),
        code: `Sub DemoAddress()
    Range("C5").Value = "here"
    Range("A1").Value = Range("C5").Address
    Range("A2").Value = Range("C5").Address(False, False)
    Range("A3").Value = Range("C5").Row
    Range("A4").Value = Range("C5").Column
    Debug.Print Range("A1").Value & " " & Range("A2").Value & " r" & Range("A3").Value & " c" & Range("A4").Value
End Sub`,
        result: L(
          'A1 = $C$5, A2 = C5, A3 = 5, A4 = 3. Immediate Window: $C$5 C5 r5 c3.',
          'A1 = $C$5, A2 = C5, A3 = 5, A4 = 3. Immediate Window: $C$5 C5 r5 c3.',
          'A1 = $C$5, A2 = C5, A3 = 5, A4 = 3. Immediate Window: $C$5 C5 r5 c3.',
          'A1 = $C$5, A2 = C5, A3 = 5, A4 = 3. Immediate Window : $C$5 C5 r5 c3.',
          'A1 = $C$5, A2 = C5, A3 = 5, A4 = 3. Direktbereich: $C$5 C5 r5 c3.',
        ),
      },
    ],
    related: ['xl-range-offset-resize', 'xl-range-rows-columns', 'xl-names-add'],
  },
  {
    id: 'xl-range-rows-columns',
    name: 'Range.Rows / Columns',
    syntax: 'Range.Rows[(Index)]  |  Range.Columns[(Index)]',
    scope: ['excel'],
    category: 'xl-range',
    subcategory: 'navigate',
    description: L(
      'Collezioni delle righe e colonne interne a un Range. .Count dà le dimensioni; Rows(2) è la seconda riga del blocco, non la riga 2 del foglio.',
      'Collections of the rows and columns inside a Range. .Count is the size; Rows(2) is the second row of the block, not sheet row 2.',
      'Colecciones de filas y columnas internas. .Count es el tamaño; Rows(2) es la segunda fila del bloque, no la fila 2 de la hoja.',
      'Collections des lignes et colonnes internes. .Count est la taille ; Rows(2) est la 2e ligne du bloc, pas la ligne 2 de la feuille.',
      'Sammlungen der Zeilen und Spalten in einem Range. .Count ist die Größe; Rows(2) ist die zweite Zeile des Blocks, nicht Blattzeile 2.',
    ),
    examples: [
      {
        title: L('Dimensioni e seconda riga del blocco', 'Size and second row of the block', 'Tamaño y segunda fila del bloque', 'Taille et deuxième ligne du bloc', 'Größe und zweite Zeile des Blocks'),
        code: `Sub DemoRowsColumns()
    Range("A1:C4").Value = 1
    Range("E1").Value = Range("A1:C4").Rows.Count
    Range("E2").Value = Range("A1:C4").Columns.Count
    Range("A1:C4").Rows(2).Interior.Color = RGB(255, 255, 0)
    Debug.Print "rows=" & Range("E1").Value & " cols=" & Range("E2").Value
End Sub`,
        result: L(
          'A1:C4 = 1. E1 = 4, E2 = 3. A2:C2 giallo. Immediate Window: rows=4 cols=3.',
          'A1:C4 = 1. E1 = 4, E2 = 3. A2:C2 yellow. Immediate Window: rows=4 cols=3.',
          'A1:C4 = 1. E1 = 4, E2 = 3. A2:C2 amarillo. Immediate Window: rows=4 cols=3.',
          'A1:C4 = 1. E1 = 4, E2 = 3. A2:C2 jaune. Immediate Window : rows=4 cols=3.',
          'A1:C4 = 1. E1 = 4, E2 = 3. A2:C2 gelb. Direktbereich: rows=4 cols=3.',
        ),
      },
    ],
    related: ['xl-entire-row-column', 'xl-range-address-row-col', 'xl-range-interior'],
  },
  {
    id: 'xl-range-end',
    name: 'Range.End',
    syntax: 'Range.End(xlToLeft | xlToRight | xlUp | xlDown)',
    scope: ['excel'],
    category: 'xl-range',
    subcategory: 'navigate',
    description: L(
      'Come Ctrl+freccia: salta al bordo del blocco contiguo. xlDown (-4121), xlUp (-4162), xlToRight (-4161), xlToLeft (-4159). Su una colonna vuota va all’ultima riga del foglio.',
      'Like Ctrl+arrow: jumps to the edge of the contiguous block. xlDown (-4121), xlUp (-4162), xlToRight (-4161), xlToLeft (-4159). On an empty column it goes to the last sheet row.',
      'Como Ctrl+flecha: salta al borde del bloque contiguo. xlDown (-4121), xlUp (-4162), xlToRight (-4161), xlToLeft (-4159). En columna vacía va a la última fila.',
      'Comme Ctrl+flèche : saute au bord du bloc contigu. xlDown (-4121), xlUp (-4162), xlToRight (-4161), xlToLeft (-4159). Colonne vide → dernière ligne.',
      'Wie Strg+Pfeil: springt an den Rand des zusammenhängenden Blocks. xlDown (-4121), xlUp (-4162), xlToRight (-4161), xlToLeft (-4159). Leere Spalte → letzte Zeile.',
    ),
    examples: [
      {
        title: L('Ultima cella piena sotto A1', 'Last filled cell below A1', 'Última celda llena bajo A1', 'Dernière cellule remplie sous A1', 'Letzte gefüllte Zelle unter A1'),
        code: `Sub DemoEnd()
    Range("A1").Value = "a"
    Range("A2").Value = "b"
    Range("A3").Value = "c"
    Range("B1").Value = Range("A1").End(xlDown).Address
    Range("B2").Value = Range("A1").End(xlDown).Value
    Debug.Print Range("B1").Value & "=" & Range("B2").Value
End Sub`,
        result: L(
          'B1 = $A$3, B2 = c. Immediate Window: $A$3=c.',
          'B1 = $A$3, B2 = c. Immediate Window: $A$3=c.',
          'B1 = $A$3, B2 = c. Immediate Window: $A$3=c.',
          'B1 = $A$3, B2 = c. Immediate Window : $A$3=c.',
          'B1 = $A$3, B2 = c. Direktbereich: $A$3=c.',
        ),
      },
    ],
    notes: L(
      'Pattern ultimo dato: se A1 è vuoto non usarlo. Meglio Cells(Rows.Count, 1).End(xlUp).',
      'Last-data pattern: do not start from empty A1. Prefer Cells(Rows.Count, 1).End(xlUp).',
      'Patrón último dato: no partas de A1 vacío. Mejor Cells(Rows.Count, 1).End(xlUp).',
      'Motif dernière donnée : ne partez pas d’A1 vide. Préférez Cells(Rows.Count, 1).End(xlUp).',
      'Letztes-Datum-Muster: nicht bei leerem A1 starten. Besser Cells(Rows.Count, 1).End(xlUp).',
    ),
    related: ['xl-range-currentregion', 'xl-range-specialcells', 'xl-range-offset-resize'],
  },
  {
    id: 'xl-range-currentregion',
    name: 'Range.CurrentRegion / UsedRange',
    syntax: 'Range.CurrentRegion  |  Worksheet.UsedRange',
    scope: ['excel'],
    category: 'xl-range',
    subcategory: 'navigate',
    description: L(
      'CurrentRegion è il blocco delimitato da righe e colonne vuote (come Ctrl+A sul dato). UsedRange è l’estensione usata del foglio e può includere formati orfani.',
      'CurrentRegion is the block bounded by empty rows and columns (like Ctrl+A on the data). UsedRange is the sheet’s used extent and can include orphan formats.',
      'CurrentRegion es el bloque delimitado por filas/columnas vacías (Ctrl+A). UsedRange es la extensión usada y puede incluir formatos huérfanos.',
      'CurrentRegion est le bloc borné par des lignes/colonnes vides (Ctrl+A). UsedRange est l’étendue utilisée et peut inclure des formats orphelins.',
      'CurrentRegion ist der von leeren Zeilen/Spalten begrenzte Block (Strg+A). UsedRange ist die genutzte Ausdehnung und kann verwaiste Formate enthalten.',
    ),
    examples: [
      {
        title: L('Regione di una tabella 3×2', 'Region of a 3×2 table', 'Región de una tabla 3×2', 'Région d’une table 3×2', 'Region einer 3×2-Tabelle'),
        code: `Sub DemoCurrentRegion()
    Range("A1").Value = "Name": Range("B1").Value = "Qty"
    Range("A2").Value = "Ann": Range("B2").Value = 3
    Range("A3").Value = "Ben": Range("B3").Value = 5
    Range("D1").Value = Range("A1").CurrentRegion.Address
    Range("D2").Value = ActiveSheet.UsedRange.Address
    Debug.Print "region=" & Range("D1").Value & " used=" & Range("D2").Value
End Sub`,
        result: L(
          'D1 = $A$1:$B$3. D2 spesso $A$1:$D$2 (include D perché ci hai scritto). Immediate Window ripete entrambi.',
          'D1 = $A$1:$B$3. D2 is often $A$1:$D$2 (includes D because you wrote there). Immediate Window repeats both.',
          'D1 = $A$1:$B$3. D2 suele ser $A$1:$D$2 (incluye D). Immediate Window repite ambos.',
          'D1 = $A$1:$B$3. D2 est souvent $A$1:$D$2 (inclut D). Immediate Window répète les deux.',
          'D1 = $A$1:$B$3. D2 ist oft $A$1:$D$2 (enthält D). Direktbereich wiederholt beide.',
        ),
      },
    ],
    related: ['xl-range-end', 'xl-range-specialcells', 'xl-range-sort'],
  },
  {
    id: 'xl-range-specialcells',
    name: 'Range.SpecialCells',
    syntax: 'Range.SpecialCells(Type, [Value])',
    scope: ['excel'],
    category: 'xl-range',
    subcategory: 'navigate',
    description: L(
      'Filtra un range per tipo: costanti, formule, vuoti, visibili, ultima cella. Solleva 1004 se non trova nulla: intercetta l’errore.',
      'Filters a range by type: constants, formulas, blanks, visible, last cell. Raises 1004 if nothing matches: trap the error.',
      'Filtra un rango por tipo: constantes, fórmulas, vacíos, visibles, última celda. Lanza 1004 si no hay nada.',
      'Filtre une plage par type : constantes, formules, vides, visibles, dernière cellule. Lève 1004 si rien.',
      'Filtert einen Bereich nach Typ: Konstanten, Formeln, Leere, Sichtbare, letzte Zelle. Löst 1004 aus, wenn nichts passt.',
    ),
    params: [
      {
        name: 'Type',
        description: L(
          'xlCellTypeConstants (2), xlCellTypeFormulas (-4123), xlCellTypeBlanks (4), xlCellTypeLastCell (11), xlCellTypeVisible (12).',
          'xlCellTypeConstants (2), xlCellTypeFormulas (-4123), xlCellTypeBlanks (4), xlCellTypeLastCell (11), xlCellTypeVisible (12).',
          'xlCellTypeConstants (2), xlCellTypeFormulas (-4123), xlCellTypeBlanks (4), xlCellTypeLastCell (11), xlCellTypeVisible (12).',
          'xlCellTypeConstants (2), xlCellTypeFormulas (-4123), xlCellTypeBlanks (4), xlCellTypeLastCell (11), xlCellTypeVisible (12).',
          'xlCellTypeConstants (2), xlCellTypeFormulas (-4123), xlCellTypeBlanks (4), xlCellTypeLastCell (11), xlCellTypeVisible (12).',
        ),
      },
    ],
    examples: [
      {
        title: L('Pulisci solo le celle vuote del blocco', 'Clear only blank cells in the block', 'Borrar solo las celdas vacías del bloque', 'Effacer seulement les cellules vides du bloc', 'Nur leere Zellen im Block löschen'),
        code: `Sub DemoSpecialCells()
    Range("A1").Value = 1
    Range("A2").ClearContents
    Range("A3").Value = 3
    Range("A4").Formula = "=A1+A3"
    On Error Resume Next
    Range("A1:A4").SpecialCells(xlCellTypeBlanks).Interior.Color = RGB(255, 200, 200)
    On Error GoTo 0
    Debug.Print "formulas=" & Range("A1:A4").SpecialCells(xlCellTypeFormulas).Address
End Sub`,
        result: L(
          'A2 sfondo rosa (vuota). Immediate Window: formulas=$A$4. A1=1, A3=3, A4=4.',
          'A2 pink background (blank). Immediate Window: formulas=$A$4. A1=1, A3=3, A4=4.',
          'A2 fondo rosa (vacía). Immediate Window: formulas=$A$4. A1=1, A3=3, A4=4.',
          'A2 fond rose (vide). Immediate Window : formulas=$A$4. A1=1, A3=3, A4=4.',
          'A2 rosa Hintergrund (leer). Direktbereich: formulas=$A$4. A1=1, A3=3, A4=4.',
        ),
      },
    ],
    related: ['xl-range-end', 'xl-wsfn-counta', 'xl-range-autofilter'],
  },
  {
    id: 'xl-entire-row-column',
    name: 'EntireRow / EntireColumn',
    syntax: 'Range.EntireRow  |  Range.EntireColumn',
    scope: ['excel'],
    category: 'xl-range',
    subcategory: 'navigate',
    description: L(
      'Espande il range all’intera riga o colonna del foglio. Utile per Hidden, Insert, Delete e AutoFit.',
      'Expands the range to the entire sheet row or column. Useful for Hidden, Insert, Delete and AutoFit.',
      'Expande el rango a toda la fila o columna. Útil para Hidden, Insert, Delete y AutoFit.',
      'Étend la plage à toute la ligne ou colonne. Utile pour Hidden, Insert, Delete et AutoFit.',
      'Erweitert den Bereich auf die ganze Blattzeile oder -spalte. Nützlich für Hidden, Insert, Delete und AutoFit.',
    ),
    examples: [
      {
        title: L('Nascondi colonna C e colora riga 2', 'Hide column C and color row 2', 'Ocultar columna C y colorear fila 2', 'Masquer la colonne C et colorer la ligne 2', 'Spalte C ausblenden und Zeile 2 färben'),
        code: `Sub DemoEntire()
    Range("B2").Value = "x"
    Range("B2").EntireRow.Interior.Color = RGB(220, 230, 255)
    Range("C1").EntireColumn.Hidden = True
    Debug.Print "row2 hidden=" & Rows(2).Hidden & " colC hidden=" & Columns("C").Hidden
End Sub`,
        result: L(
          'Riga 2 azzurra, colonna C nascosta. Immediate Window: row2 hidden=False colC hidden=True.',
          'Row 2 light blue, column C hidden. Immediate Window: row2 hidden=False colC hidden=True.',
          'Fila 2 azulada, columna C oculta. Immediate Window: row2 hidden=False colC hidden=True.',
          'Ligne 2 bleu clair, colonne C masquée. Immediate Window : row2 hidden=False colC hidden=True.',
          'Zeile 2 hellblau, Spalte C ausgeblendet. Direktbereich: row2 hidden=False colC hidden=True.',
        ),
      },
    ],
    related: ['xl-range-rows-columns', 'xl-range-insert-delete', 'xl-range-display'],
  },
  {
    id: 'xl-range-copy-cut',
    name: 'Range.Copy / Cut',
    syntax: 'Range.Copy [Destination]  |  Range.Cut [Destination]',
    scope: ['excel'],
    category: 'xl-range',
    subcategory: 'edit',
    description: L(
      'Copy duplica valori e formati. Cut sposta (svuota l’origine). Con Destination l’operazione è immediata; senza, restano le formichine per Paste/PasteSpecial.',
      'Copy duplicates values and formats. Cut moves (clears the source). With Destination the action is immediate; without it, marching ants wait for Paste/PasteSpecial.',
      'Copy duplica valores y formatos. Cut mueve (vacía el origen). Con Destination es inmediato; sin ella quedan hormigas para Paste.',
      'Copy duplique valeurs et formats. Cut déplace (vide la source). Avec Destination c’est immédiat ; sinon les pointillés attendent Paste.',
      'Copy dupliziert Werte und Formate. Cut verschiebt (leert die Quelle). Mit Destination sofort; sonst wartet der Laufrahmen auf Paste.',
    ),
    params: [
      { name: 'Destination', optional: true, description: L('Range di destinazione.', 'Destination range.', 'Rango de destino.', 'Plage de destination.', 'Zielbereich.') },
    ],
    examples: [
      {
        title: L('Copia A1 su C1 e taglia B1 su D1', 'Copy A1 to C1 and cut B1 to D1', 'Copiar A1 a C1 y cortar B1 a D1', 'Copier A1 vers C1 et couper B1 vers D1', 'A1 nach C1 kopieren und B1 nach D1 ausschneiden'),
        code: `Sub DemoCopyCut()
    Range("A1").Value = "keep"
    Range("B1").Value = "move"
    Range("A1").Copy Destination:=Range("C1")
    Range("B1").Cut Destination:=Range("D1")
    Application.CutCopyMode = False
    Debug.Print "A1=" & Range("A1").Value & " B1=" & Range("B1").Value & _
        " C1=" & Range("C1").Value & " D1=" & Range("D1").Value
End Sub`,
        result: L(
          'A1 = keep, B1 vuota, C1 = keep, D1 = move. Immediate Window: A1=keep B1= C1=keep D1=move.',
          'A1 = keep, B1 blank, C1 = keep, D1 = move. Immediate Window: A1=keep B1= C1=keep D1=move.',
          'A1 = keep, B1 vacía, C1 = keep, D1 = move. Immediate Window: A1=keep B1= C1=keep D1=move.',
          'A1 = keep, B1 vide, C1 = keep, D1 = move. Immediate Window : A1=keep B1= C1=keep D1=move.',
          'A1 = keep, B1 leer, C1 = keep, D1 = move. Direktbereich: A1=keep B1= C1=keep D1=move.',
        ),
      },
    ],
    related: ['xl-range-pastespecial', 'xl-cutcopymode'],
  },
  {
    id: 'xl-range-pastespecial',
    name: 'Range.PasteSpecial',
    syntax: 'Range.PasteSpecial [Paste], [Operation], [SkipBlanks], [Transpose]',
    scope: ['excel'],
    category: 'xl-range',
    subcategory: 'edit',
    description: L(
      'Incolla dopo Copy scegliendo cosa: valori, formati, formule, operazioni (Add) o trasposizione. Richiede CutCopyMode attivo.',
      'Pastes after Copy, choosing what: values, formats, formulas, operations (Add) or transpose. Requires an active CutCopyMode.',
      'Pega tras Copy eligiendo qué: valores, formatos, fórmulas, operaciones (Add) o trasponer. Requiere CutCopyMode activo.',
      'Colle après Copy en choisissant : valeurs, formats, formules, opérations (Add) ou transposition. CutCopyMode requis.',
      'Fügt nach Copy wählbar ein: Werte, Formate, Formeln, Operationen (Add) oder Transponieren. CutCopyMode nötig.',
    ),
    params: [
      { name: 'Paste', optional: true, description: L('xlPasteValues (-4163), xlPasteFormats (-4122), xlPasteFormulas (-4123), xlPasteAll (-4104).', 'xlPasteValues (-4163), xlPasteFormats (-4122), xlPasteFormulas (-4123), xlPasteAll (-4104).', 'xlPasteValues (-4163), xlPasteFormats (-4122), xlPasteFormulas (-4123), xlPasteAll (-4104).', 'xlPasteValues (-4163), xlPasteFormats (-4122), xlPasteFormulas (-4123), xlPasteAll (-4104).', 'xlPasteValues (-4163), xlPasteFormats (-4122), xlPasteFormulas (-4123), xlPasteAll (-4104).') },
      { name: 'Operation', optional: true, description: L('xlPasteSpecialOperationAdd (2) somma sui destini.', 'xlPasteSpecialOperationAdd (2) adds onto destinations.', 'xlPasteSpecialOperationAdd (2) suma sobre el destino.', 'xlPasteSpecialOperationAdd (2) additionne sur la destination.', 'xlPasteSpecialOperationAdd (2) addiert auf das Ziel.') },
      { name: 'Transpose', optional: true, description: L('True scambia righe e colonne.', 'True swaps rows and columns.', 'True intercambia filas y columnas.', 'True inverse lignes et colonnes.', 'True tauscht Zeilen und Spalten.') },
    ],
    examples: [
      {
        title: L('Incolla solo valori e trasponi', 'Paste values only and transpose', 'Pegar solo valores y trasponer', 'Coller les valeurs seules et transposer', 'Nur Werte einfügen und transponieren'),
        code: `Sub DemoPasteSpecial()
    Range("A1").Formula = "=1+1"
    Range("A1").Copy
    Range("C1").PasteSpecial Paste:=xlPasteValues
    Range("A1").Copy
    Range("E1").PasteSpecial Paste:=xlPasteValues, Transpose:=True
    Application.CutCopyMode = False
    Debug.Print "A1 formula=" & Range("A1").Formula & " C1 value=" & Range("C1").Value
End Sub`,
        result: L(
          'A1 resta =1+1 (valore 2). C1 = 2 senza formula. E1 = 2. Immediate Window: A1 formula==1+1 C1 value=2.',
          'A1 stays =1+1 (value 2). C1 = 2 with no formula. E1 = 2. Immediate Window: A1 formula==1+1 C1 value=2.',
          'A1 sigue =1+1 (valor 2). C1 = 2 sin fórmula. E1 = 2. Immediate Window: A1 formula==1+1 C1 value=2.',
          'A1 reste =1+1 (valeur 2). C1 = 2 sans formule. E1 = 2. Immediate Window : A1 formula==1+1 C1 value=2.',
          'A1 bleibt =1+1 (Wert 2). C1 = 2 ohne Formel. E1 = 2. Direktbereich: A1 formula==1+1 C1 value=2.',
        ),
      },
    ],
    related: ['xl-range-copy-cut', 'xl-cutcopymode', 'xl-range-value'],
  },
  {
    id: 'xl-range-insert-delete',
    name: 'Range.Insert / Delete',
    syntax: 'Range.Insert [Shift], [CopyOrigin]  |  Range.Delete [Shift]',
    scope: ['excel'],
    category: 'xl-range',
    subcategory: 'edit',
    description: L(
      'Insert crea celle/righe/colonne spostando il resto. Delete le rimuove. Shift: xlShiftDown (-4121), xlShiftToRight (-4161), xlUp (-4162), xlToLeft (-4159).',
      'Insert creates cells/rows/columns and shifts the rest. Delete removes them. Shift: xlShiftDown (-4121), xlShiftToRight (-4161), xlUp (-4162), xlToLeft (-4159).',
      'Insert crea celdas/filas/columnas y desplaza el resto. Delete las quita. Shift: xlShiftDown, xlShiftToRight, xlUp, xlToLeft.',
      'Insert crée cellules/lignes/colonnes et décale le reste. Delete les enlève. Shift : xlShiftDown, xlShiftToRight, xlUp, xlToLeft.',
      'Insert erzeugt Zellen/Zeilen/Spalten und verschiebt den Rest. Delete entfernt sie. Shift: xlShiftDown, xlShiftToRight, xlUp, xlToLeft.',
    ),
    examples: [
      {
        title: L('Inserisci una riga e cancella una cella', 'Insert a row and delete a cell', 'Insertar una fila y borrar una celda', 'Insérer une ligne et supprimer une cellule', 'Zeile einfügen und Zelle löschen'),
        code: `Sub DemoInsertDelete()
    Range("A1").Value = "one"
    Range("A2").Value = "two"
    Range("A2").EntireRow.Insert
    Range("A2").Value = "inserted"
    Range("A3").Delete Shift:=xlUp
    Debug.Print Range("A1").Value & "," & Range("A2").Value & "," & Range("A3").Value
End Sub`,
        result: L(
          'A1 = one, A2 = inserted, A3 vuota (two è stato eliminato). Immediate Window: one,inserted,',
          'A1 = one, A2 = inserted, A3 blank (two was deleted). Immediate Window: one,inserted,',
          'A1 = one, A2 = inserted, A3 vacía (two se eliminó). Immediate Window: one,inserted,',
          'A1 = one, A2 = inserted, A3 vide (two a été supprimé). Immediate Window : one,inserted,',
          'A1 = one, A2 = inserted, A3 leer (two wurde gelöscht). Direktbereich: one,inserted,',
        ),
      },
    ],
    related: ['xl-entire-row-column', 'xl-range-clear', 'xl-displayalerts'],
  },
  {
    id: 'xl-range-clear',
    name: 'Range.Clear / ClearContents / ClearFormats',
    syntax: 'Range.Clear  |  Range.ClearContents  |  Range.ClearFormats',
    scope: ['excel'],
    category: 'xl-range',
    subcategory: 'edit',
    description: L(
      'Clear svuota valori, formule, formati, commenti e convalida. ClearContents toglie solo valori/formule. ClearFormats toglie solo il formato.',
      'Clear wipes values, formulas, formats, comments and validation. ClearContents removes only values/formulas. ClearFormats removes only formatting.',
      'Clear borra valores, fórmulas, formatos, comentarios y validación. ClearContents solo valores/fórmulas. ClearFormats solo formato.',
      'Clear efface valeurs, formules, formats, commentaires et validation. ClearContents seulement valeurs/formules. ClearFormats seulement le format.',
      'Clear löscht Werte, Formeln, Formate, Kommentare und Validierung. ClearContents nur Werte/Formeln. ClearFormats nur Format.',
    ),
    examples: [
      {
        title: L('Tre pulizie a confronto', 'Three clears compared', 'Tres limpiezas comparadas', 'Trois effacements comparés', 'Drei Löscharten im Vergleich'),
        code: `Sub DemoClear()
    Range("A1").Value = 10: Range("A1").Font.Bold = True
    Range("A2").Value = 20: Range("A2").Font.Bold = True
    Range("A3").Value = 30: Range("A3").Font.Bold = True
    Range("A1").ClearContents
    Range("A2").ClearFormats
    Range("A3").Clear
    Debug.Print "A1=[" & Range("A1").Value & "] bold=" & Range("A1").Font.Bold
    Debug.Print "A2=" & Range("A2").Value & " bold=" & Range("A2").Font.Bold
    Debug.Print "A3=[" & Range("A3").Value & "] bold=" & Range("A3").Font.Bold
End Sub`,
        result: L(
          'A1 vuota ma ancora grassetto. A2 = 20 non grassetto. A3 vuota e senza formato. Immediate Window conferma i tre stati.',
          'A1 blank but still bold. A2 = 20 not bold. A3 blank and unformatted. Immediate Window confirms the three states.',
          'A1 vacía pero aún negrita. A2 = 20 sin negrita. A3 vacía y sin formato. Immediate Window confirma los tres.',
          'A1 vide mais encore gras. A2 = 20 non gras. A3 vide et sans format. Immediate Window confirme les trois.',
          'A1 leer, aber noch fett. A2 = 20 nicht fett. A3 leer und unformatiert. Direktbereich bestätigt die drei Zustände.',
        ),
      },
    ],
    related: ['xl-range-insert-delete', 'xl-range-value', 'xl-range-font'],
  },
  {
    id: 'xl-range-autofill',
    name: 'Range.AutoFill',
    syntax: 'Range.AutoFill Destination, [Type]',
    scope: ['excel'],
    category: 'xl-range',
    subcategory: 'edit',
    description: L(
      'Prolunga una serie come il quadratino di riempimento. Destination deve includere l’origine. Type: xlFillSeries (2), xlFillCopy (1), xlFillDefault (0).',
      'Extends a series like the fill handle. Destination must include the source. Type: xlFillSeries (2), xlFillCopy (1), xlFillDefault (0).',
      'Prolonga una serie como el controlador de relleno. Destination debe incluir el origen. Type: xlFillSeries (2), xlFillCopy (1), xlFillDefault (0).',
      'Prolonge une série comme la poignée de recopie. Destination doit inclure la source. Type : xlFillSeries (2), xlFillCopy (1), xlFillDefault (0).',
      'Setzt eine Reihe fort wie der Ausfüllkästchen. Destination muss die Quelle enthalten. Type: xlFillSeries (2), xlFillCopy (1), xlFillDefault (0).',
    ),
    params: [
      { name: 'Destination', description: L('Range destinazione (include il seme).', 'Destination range (includes the seed).', 'Rango destino (incluye la semilla).', 'Plage destination (inclut la graine).', 'Zielbereich (enthält den Samen).') },
      { name: 'Type', optional: true, description: L('xlFillDefault, xlFillCopy, xlFillSeries, xlFillFormats, xlFillValues.', 'xlFillDefault, xlFillCopy, xlFillSeries, xlFillFormats, xlFillValues.', 'xlFillDefault, xlFillCopy, xlFillSeries, xlFillFormats, xlFillValues.', 'xlFillDefault, xlFillCopy, xlFillSeries, xlFillFormats, xlFillValues.', 'xlFillDefault, xlFillCopy, xlFillSeries, xlFillFormats, xlFillValues.') },
    ],
    examples: [
      {
        title: L('Serie 1,2,…,6', 'Series 1,2,…,6', 'Serie 1,2,…,6', 'Série 1,2,…,6', 'Reihe 1,2,…,6'),
        code: `Sub DemoAutoFill()
    Range("A1").Value = 1
    Range("A2").Value = 2
    Range("A1:A2").AutoFill Destination:=Range("A1:A6"), Type:=xlFillSeries
    Debug.Print Range("A6").Value
End Sub`,
        result: L(
          'A1:A6 = 1,2,3,4,5,6. Immediate Window: 6.',
          'A1:A6 = 1,2,3,4,5,6. Immediate Window: 6.',
          'A1:A6 = 1,2,3,4,5,6. Immediate Window: 6.',
          'A1:A6 = 1,2,3,4,5,6. Immediate Window : 6.',
          'A1:A6 = 1,2,3,4,5,6. Direktbereich: 6.',
        ),
      },
    ],
    related: ['xl-range-value', 'xl-range-offset-resize'],
  },
  {
    id: 'xl-range-calculate',
    name: 'Range.Calculate',
    syntax: 'Range.Calculate  |  Worksheet.Calculate  |  Application.Calculate',
    scope: ['excel'],
    category: 'xl-range',
    subcategory: 'edit',
    description: L(
      'Ricalcola solo quell’intervallo (utile in Calculation = Manual). Worksheet.Calculate fa il foglio; Application.Calculate tutta la sessione.',
      'Recalculates only that range (useful when Calculation = Manual). Worksheet.Calculate does the sheet; Application.Calculate the whole session.',
      'Recalcula solo ese rango (útil con Calculation = Manual). Worksheet.Calculate hace la hoja; Application.Calculate la sesión.',
      'Recalcule seulement cette plage (utile si Calculation = Manual). Worksheet.Calculate fait la feuille ; Application.Calculate la session.',
      'Berechnet nur diesen Bereich neu (nützlich bei Calculation = Manual). Worksheet.Calculate das Blatt; Application.Calculate die Sitzung.',
    ),
    examples: [
      {
        title: L('Ricalcola solo C1 in modalità manuale', 'Recalculate only C1 in manual mode', 'Recalcular solo C1 en modo manual', 'Recalculer seulement C1 en mode manuel', 'Nur C1 im manuellen Modus neu berechnen'),
        code: `Sub DemoRangeCalculate()
    Application.Calculation = xlCalculationManual
    Range("A1").Value = 10
    Range("B1").Value = 5
    Range("C1").Formula = "=A1*B1"
    Range("D1").Formula = "=A1+B1"
    Range("C1").Calculate
    Debug.Print "C1=" & Range("C1").Value & " D1=" & Range("D1").Text
    Application.Calculation = xlCalculationAutomatic
End Sub`,
        result: L(
          'C1 = 50. D1 resta 0 o vuoto finché non ricalcoli tutto. Immediate Window: C1=50 D1=0. Poi Calculation torna Automatic.',
          'C1 = 50. D1 stays 0 or blank until a full recalc. Immediate Window: C1=50 D1=0. Then Calculation returns to Automatic.',
          'C1 = 50. D1 sigue 0 o vacío hasta un recálculo total. Immediate Window: C1=50 D1=0. Luego Calculation vuelve a Automatic.',
          'C1 = 50. D1 reste 0 ou vide jusqu’au recalcul global. Immediate Window : C1=50 D1=0. Puis Calculation redevient Automatic.',
          'C1 = 50. D1 bleibt 0 oder leer bis zur Vollberechnung. Direktbereich: C1=50 D1=0. Dann Calculation wieder Automatic.',
        ),
      },
    ],
    related: ['xl-calculation', 'xl-range-formula'],
  },
  {
    id: 'xl-union-intersect',
    name: 'Union / Intersect',
    syntax: 'Union(Range1, Range2, ...)  |  Intersect(Range1, Range2, ...)',
    scope: ['excel'],
    category: 'xl-range',
    subcategory: 'edit',
    description: L(
      'Union unisce aree (anche non contigue). Intersect è la sovrapposizione; restituisce Nothing se non si toccano. Entrambe sono funzioni di Application.',
      'Union joins areas (including non-contiguous). Intersect is the overlap; it returns Nothing if they do not touch. Both are Application functions.',
      'Union une áreas (también no contiguas). Intersect es la superposición; Nothing si no se tocan. Ambas son funciones de Application.',
      'Union unit des zones (y compris non contiguës). Intersect est le chevauchement ; Nothing s’ils ne se touchent pas. Toutes deux sont des fonctions Application.',
      'Union vereint Bereiche (auch nicht zusammenhängend). Intersect ist die Überlappung; Nothing, wenn sie sich nicht berühren. Beides Application-Funktionen.',
    ),
    examples: [
      {
        title: L('Colora unione e scrivi l’intersezione', 'Color the union and write the intersection', 'Colorear la unión y escribir la intersección', 'Colorer l’union et écrire l’intersection', 'Vereinigung färben und Schnittmenge schreiben'),
        code: `Sub DemoUnionIntersect()
    Dim u As Range, i As Range
    Set u = Union(Range("A1:B2"), Range("D1"))
    u.Interior.Color = RGB(200, 255, 200)
    Set i = Intersect(Range("A1:C3"), Range("B2:D5"))
    If Not i Is Nothing Then
        i.Value = "hit"
        Debug.Print "intersect=" & i.Address
    End If
    Debug.Print "union=" & u.Address
End Sub`,
        result: L(
          'A1:B2 e D1 verdi. B2:C3 = hit. Immediate Window: intersect=$B$2:$C$3 union=$A$1:$B$2,$D$1.',
          'A1:B2 and D1 green. B2:C3 = hit. Immediate Window: intersect=$B$2:$C$3 union=$A$1:$B$2,$D$1.',
          'A1:B2 y D1 verdes. B2:C3 = hit. Immediate Window: intersect=$B$2:$C$3 union=$A$1:$B$2,$D$1.',
          'A1:B2 et D1 verts. B2:C3 = hit. Immediate Window : intersect=$B$2:$C$3 union=$A$1:$B$2,$D$1.',
          'A1:B2 und D1 grün. B2:C3 = hit. Direktbereich: intersect=$B$2:$C$3 union=$A$1:$B$2,$D$1.',
        ),
      },
    ],
    related: ['xl-range-specialcells', 'xl-worksheet-change', 'xl-range-copy-cut'],
  },
  {
    id: 'xl-range-find',
    name: 'Range.Find / FindNext',
    syntax: 'Range.Find(What, [After], [LookIn], [LookAt], ...)  |  Range.FindNext([After])',
    scope: ['excel'],
    category: 'xl-analyze',
    subcategory: 'find',
    description: L(
      'Cerca What e restituisce la prima cella (Nothing se assente). FindNext continua dal punto After. Specifica sempre LookIn/LookAt/MatchCase: Find ricorda le ultime impostazioni della UI.',
      'Searches for What and returns the first cell (Nothing if missing). FindNext continues from After. Always specify LookIn/LookAt/MatchCase: Find remembers the last UI settings.',
      'Busca What y devuelve la primera celda (Nothing si falta). FindNext sigue desde After. Especifica siempre LookIn/LookAt/MatchCase.',
      'Cherche What et renvoie la première cellule (Nothing si absent). FindNext continue depuis After. Spécifiez toujours LookIn/LookAt/MatchCase.',
      'Sucht What und liefert die erste Zelle (Nothing wenn fehlend). FindNext macht bei After weiter. LookIn/LookAt/MatchCase immer setzen.',
    ),
    params: [
      { name: 'What', description: L('Testo o numero da trovare.', 'Text or number to find.', 'Texto o número a buscar.', 'Texte ou nombre à trouver.', 'Zu suchender Text oder Zahl.') },
      { name: 'LookIn', optional: true, description: L('xlValues (-4163) o xlFormulas (-4123).', 'xlValues (-4163) or xlFormulas (-4123).', 'xlValues (-4163) o xlFormulas (-4123).', 'xlValues (-4163) ou xlFormulas (-4123).', 'xlValues (-4163) oder xlFormulas (-4123).') },
      { name: 'LookAt', optional: true, description: L('xlWhole (1) o xlPart (2).', 'xlWhole (1) or xlPart (2).', 'xlWhole (1) o xlPart (2).', 'xlWhole (1) ou xlPart (2).', 'xlWhole (1) oder xlPart (2).') },
    ],
    examples: [
      {
        title: L('Elenca tutte le “apple”', 'List every “apple”', 'Listar todas las “apple”', 'Lister toutes les « apple »', 'Alle „apple“ auflisten'),
        code: `Sub DemoFind()
    Dim c As Range, first As String
    Range("A1").Value = "apple": Range("A2").Value = "pear": Range("A3").Value = "apple"
    Set c = Range("A1:A3").Find(What:="apple", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If c Is Nothing Then Exit Sub
    first = c.Address
    Do
        c.Interior.Color = RGB(255, 255, 150)
        Debug.Print c.Address
        Set c = Range("A1:A3").FindNext(c)
    Loop While Not c Is Nothing And c.Address <> first
End Sub`,
        result: L(
          'A1 e A3 gialle. Immediate Window: $A$1 poi $A$3. A2 resta pear senza colore.',
          'A1 and A3 yellow. Immediate Window: $A$1 then $A$3. A2 stays pear with no color.',
          'A1 y A3 amarillas. Immediate Window: $A$1 luego $A$3. A2 sigue pear sin color.',
          'A1 et A3 jaunes. Immediate Window : $A$1 puis $A$3. A2 reste pear sans couleur.',
          'A1 und A3 gelb. Direktbereich: $A$1, dann $A$3. A2 bleibt pear ohne Farbe.',
        ),
      },
    ],
    notes: L(
      'Ferma il loop confrontando c.Address con il primo indirizzo, altrimenti FindNext gira all’infinito.',
      'Stop the loop by comparing c.Address to the first address, or FindNext circles forever.',
      'Detén el bucle comparando c.Address con la primera dirección o FindNext gira sin fin.',
      'Arrêtez la boucle en comparant c.Address à la première adresse, sinon FindNext tourne en boucle.',
      'Schleife stoppen, indem c.Address mit der ersten Adresse verglichen wird, sonst kreist FindNext endlos.',
    ),
    related: ['xl-range-replace', 'xl-wsfn-match', 'xl-range-autofilter'],
  },
  {
    id: 'xl-range-replace',
    name: 'Range.Replace',
    syntax: 'Range.Replace What, Replacement, [LookAt], [SearchOrder], [MatchCase], ...',
    scope: ['excel'],
    category: 'xl-analyze',
    subcategory: 'find',
    description: L(
      'Sostituisce tutte le occorrenze di What con Replacement nell’intervallo. Restituisce True se almeno una sostituzione è avvenuta.',
      'Replaces every occurrence of What with Replacement in the range. Returns True if at least one replacement happened.',
      'Sustituye todas las ocurrencias de What por Replacement. Devuelve True si hubo al menos un cambio.',
      'Remplace toutes les occurrences de What par Replacement. Renvoie True s’il y a eu au moins un changement.',
      'Ersetzt jedes Vorkommen von What durch Replacement. Liefert True, wenn mindestens ein Ersatz erfolgte.',
    ),
    examples: [
      {
        title: L('Sostituisci TBD con Done', 'Replace TBD with Done', 'Sustituir TBD por Done', 'Remplacer TBD par Done', 'TBD durch Done ersetzen'),
        code: `Sub DemoReplace()
    Range("A1").Value = "TBD start"
    Range("A2").Value = "ok"
    Range("A3").Value = "TBD end"
    Range("A1:A3").Replace What:="TBD", Replacement:="Done", LookAt:=xlPart
    Debug.Print Range("A1").Value & " | " & Range("A3").Value
End Sub`,
        result: L(
          'A1 = Done start, A2 = ok, A3 = Done end. Immediate Window: Done start | Done end.',
          'A1 = Done start, A2 = ok, A3 = Done end. Immediate Window: Done start | Done end.',
          'A1 = Done start, A2 = ok, A3 = Done end. Immediate Window: Done start | Done end.',
          'A1 = Done start, A2 = ok, A3 = Done end. Immediate Window : Done start | Done end.',
          'A1 = Done start, A2 = ok, A3 = Done end. Direktbereich: Done start | Done end.',
        ),
      },
    ],
    related: ['xl-range-find', 'xl-range-value'],
  },
  {
    id: 'xl-range-sort',
    name: 'Range.Sort',
    syntax: 'Range.Sort Key1, [Order1], [Key2], [Type], [Order2], [Key3], [Order3], [Header], ...',
    scope: ['excel'],
    category: 'xl-analyze',
    subcategory: 'find',
    description: L(
      'Ordina il blocco su una o più chiavi. Header:=xlYes (1) esclude la prima riga. Order: xlAscending (1), xlDescending (2).',
      'Sorts the block on one or more keys. Header:=xlYes (1) excludes the first row. Order: xlAscending (1), xlDescending (2).',
      'Ordena el bloque por una o más claves. Header:=xlYes (1) excluye la primera fila. Order: xlAscending (1), xlDescending (2).',
      'Trie le bloc sur une ou plusieurs clés. Header:=xlYes (1) exclut la première ligne. Order : xlAscending (1), xlDescending (2).',
      'Sortiert den Block nach einer oder mehreren Schlüsseln. Header:=xlYes (1) schließt die erste Zeile aus. Order: xlAscending (1), xlDescending (2).',
    ),
    examples: [
      {
        title: L('Ordina nomi in A', 'Sort names in A', 'Ordenar nombres en A', 'Trier les noms en A', 'Namen in A sortieren'),
        code: `Sub DemoSort()
    Range("A1").Value = "Name": Range("B1").Value = "Qty"
    Range("A2").Value = "Zoe": Range("B2").Value = 3
    Range("A3").Value = "Ann": Range("B3").Value = 9
    Range("A1:B3").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
    Debug.Print Range("A2").Value & "=" & Range("B2").Value
End Sub`,
        result: L(
          'A2 = Ann / B2 = 9, A3 = Zoe / B3 = 3. Intestazione resta in riga 1. Immediate Window: Ann=9.',
          'A2 = Ann / B2 = 9, A3 = Zoe / B3 = 3. Header stays on row 1. Immediate Window: Ann=9.',
          'A2 = Ann / B2 = 9, A3 = Zoe / B3 = 3. El encabezado sigue en la fila 1. Immediate Window: Ann=9.',
          'A2 = Ann / B2 = 9, A3 = Zoe / B3 = 3. L’en-tête reste ligne 1. Immediate Window : Ann=9.',
          'A2 = Ann / B2 = 9, A3 = Zoe / B3 = 3. Kopfzeile bleibt Zeile 1. Direktbereich: Ann=9.',
        ),
      },
    ],
    related: ['xl-range-autofilter', 'xl-range-currentregion'],
  },
  {
    id: 'xl-range-autofilter',
    name: 'Range.AutoFilter',
    syntax: 'Range.AutoFilter [Field], [Criteria1], [Operator], [Criteria2], [VisibleDropDown]',
    scope: ['excel'],
    category: 'xl-analyze',
    subcategory: 'find',
    description: L(
      'Applica o toglie il filtro automatico. Field è l’indice colonna 1-based rispetto al range. Senza argomenti fa da interruttore. ShowAllData toglie i criteri senza togliere le frecce.',
      'Applies or clears AutoFilter. Field is the 1-based column index within the range. With no arguments it toggles. ShowAllData clears criteria without removing the arrows.',
      'Aplica o quita AutoFilter. Field es el índice de columna base 1. Sin argumentos conmuta. ShowAllData quita criterios sin quitar flechas.',
      'Applique ou ôte AutoFilter. Field est l’index de colonne base 1. Sans arguments, bascule. ShowAllData ôte les critères sans ôter les flèches.',
      'Setzt oder entfernt AutoFilter. Field ist der 1-basierte Spaltenindex. Ohne Argumente schaltet um. ShowAllData löscht Kriterien ohne die Pfeile.',
    ),
    examples: [
      {
        title: L('Filtra la colonna Region = West', 'Filter Region column = West', 'Filtrar columna Region = West', 'Filtrer la colonne Region = West', 'Spalte Region = West filtern'),
        code: `Sub DemoAutoFilter()
    Range("A1").Value = "Region": Range("B1").Value = "Amt"
    Range("A2").Value = "West": Range("B2").Value = 10
    Range("A3").Value = "East": Range("B3").Value = 4
    Range("A4").Value = "West": Range("B4").Value = 7
    Range("A1:B4").AutoFilter Field:=1, Criteria1:="West"
    Debug.Print "AutoFilterMode=" & ActiveSheet.AutoFilterMode
    Debug.Print "visible rows=" & Range("A2:A4").SpecialCells(xlCellTypeVisible).Count
End Sub`,
        result: L(
          'Frecce sul titolo. Righe East nascoste. Immediate Window: AutoFilterMode=True visible rows=2 (A2 e A4).',
          'Arrows on the header. East rows hidden. Immediate Window: AutoFilterMode=True visible rows=2 (A2 and A4).',
          'Flechas en el encabezado. Filas East ocultas. Immediate Window: AutoFilterMode=True visible rows=2.',
          'Flèches sur l’en-tête. Lignes East masquées. Immediate Window : AutoFilterMode=True visible rows=2.',
          'Pfeile auf der Kopfzeile. East-Zeilen ausgeblendet. Direktbereich: AutoFilterMode=True visible rows=2.',
        ),
      },
    ],
    notes: L(
      'ActiveSheet.AutoFilterMode = False toglie il filtro. Prima di SpecialCells(xlCellTypeVisible) verifica che ci siano celle visibili.',
      'ActiveSheet.AutoFilterMode = False removes the filter. Before SpecialCells(xlCellTypeVisible) check that visible cells exist.',
      'ActiveSheet.AutoFilterMode = False quita el filtro. Antes de SpecialCells(xlCellTypeVisible) comprueba que haya visibles.',
      'ActiveSheet.AutoFilterMode = False ôte le filtre. Avant SpecialCells(xlCellTypeVisible) vérifiez qu’il reste des cellules visibles.',
      'ActiveSheet.AutoFilterMode = False entfernt den Filter. Vor SpecialCells(xlCellTypeVisible) prüfen, ob sichtbare Zellen existieren.',
    ),
    related: ['xl-range-sort', 'xl-range-specialcells', 'xl-range-find'],
  },
  {
    id: 'xl-range-autofit',
    name: 'Range.AutoFit',
    syntax: 'Range.Columns.AutoFit  |  Range.Rows.AutoFit  |  Columns("A:C").AutoFit',
    scope: ['excel'],
    category: 'xl-analyze',
    subcategory: 'find',
    description: L(
      'Adatta larghezza colonne o altezza righe al contenuto. Chiama AutoFit su Columns o Rows, non sul Range di celle misto.',
      'Fits column widths or row heights to the content. Call AutoFit on Columns or Rows, not on a mixed cell Range.',
      'Ajusta anchos de columna o altos de fila al contenido. Llama AutoFit en Columns o Rows, no en un Range mixto.',
      'Ajuste largeurs de colonnes ou hauteurs de lignes au contenu. Appelez AutoFit sur Columns ou Rows, pas sur un Range mixte.',
      'Passt Spaltenbreiten oder Zeilenhöhen an den Inhalt an. AutoFit auf Columns oder Rows aufrufen, nicht auf gemischte Range.',
    ),
    examples: [
      {
        title: L('Allarga la colonna A al testo', 'Widen column A to the text', 'Ensachar la columna A al texto', 'Élargir la colonne A au texte', 'Spalte A an den Text anpassen'),
        code: `Sub DemoAutoFit()
    Range("A1").Value = "A very long header that needs space"
    Columns("A").ColumnWidth = 4
    Debug.Print "before=" & Columns("A").ColumnWidth
    Columns("A").AutoFit
    Debug.Print "after=" & Columns("A").ColumnWidth
End Sub`,
        result: L(
          'La colonna A passa da larghezza 4 a circa 35+ e il testo è interamente visibile. Immediate Window: before=4 after≈35.',
          'Column A grows from width 4 to about 35+ and the text is fully visible. Immediate Window: before=4 after≈35.',
          'La columna A pasa de ancho 4 a unos 35+ y el texto se ve entero. Immediate Window: before=4 after≈35.',
          'La colonne A passe de 4 à environ 35+ et le texte est entièrement visible. Immediate Window : before=4 after≈35.',
          'Spalte A wächst von Breite 4 auf etwa 35+ und der Text ist vollständig sichtbar. Direktbereich: before=4 after≈35.',
        ),
      },
    ],
    related: ['xl-range-text-format', 'xl-entire-row-column', 'xl-pagesetup-print'],
  },
  {
    id: 'xl-range-font',
    name: 'Range.Font',
    syntax: 'Range.Font.Name / Size / Bold / Italic / Color / Underline',
    scope: ['excel'],
    category: 'xl-analyze',
    subcategory: 'format',
    description: L(
      'Oggetto Font della cella: nome, dimensione, grassetto, corsivo, colore RGB e sottolineatura. Null in lettura significa “misto” su un range multi-cella.',
      'Cell Font object: name, size, bold, italic, RGB color and underline. Null on read means mixed values across a multi-cell range.',
      'Objeto Font: nombre, tamaño, negrita, cursiva, color RGB y subrayado. Null en lectura significa mixto en un rango.',
      'Objet Font : nom, taille, gras, italique, couleur RGB et soulignement. Null à la lecture = mixte sur une plage.',
      'Font-Objekt: Name, Größe, Fett, Kursiv, RGB-Farbe und Unterstreichung. Null beim Lesen bedeutet gemischt.',
    ),
    examples: [
      {
        title: L('Titolo rosso in Calibri 14', 'Red Calibri 14 title', 'Título rojo Calibri 14', 'Titre rouge Calibri 14', 'Roter Calibri-14-Titel'),
        code: `Sub DemoFont()
    Range("A1").Value = "Report"
    With Range("A1").Font
        .Name = "Calibri"
        .Size = 14
        .Bold = True
        .Color = RGB(192, 0, 0)
    End With
    Debug.Print Range("A1").Font.Name & " " & Range("A1").Font.Size & " bold=" & Range("A1").Font.Bold
End Sub`,
        result: L(
          'A1 = Report, Calibri 14 grassetto rosso. Immediate Window: Calibri 14 bold=True.',
          'A1 = Report, Calibri 14 bold red. Immediate Window: Calibri 14 bold=True.',
          'A1 = Report, Calibri 14 negrita rojo. Immediate Window: Calibri 14 bold=True.',
          'A1 = Report, Calibri 14 gras rouge. Immediate Window : Calibri 14 bold=True.',
          'A1 = Report, Calibri 14 fett rot. Direktbereich: Calibri 14 bold=True.',
        ),
      },
    ],
    related: ['xl-range-interior', 'xl-range-borders', 'xl-dialogs'],
  },
  {
    id: 'xl-range-interior',
    name: 'Range.Interior',
    syntax: 'Range.Interior.Color / ColorIndex / Pattern',
    scope: ['excel'],
    category: 'xl-analyze',
    subcategory: 'format',
    description: L(
      'Sfondo della cella. Color accetta RGB; ColorIndex usa la palette (xlNone = -4142 per nessuno). Pattern (xlSolid = 1) per campiture.',
      'Cell background. Color takes RGB; ColorIndex uses the palette (xlNone = -4142 for none). Pattern (xlSolid = 1) for fills.',
      'Fondo de la celda. Color acepta RGB; ColorIndex usa la paleta (xlNone = -4142). Pattern (xlSolid = 1) para tramas.',
      'Arrière-plan. Color accepte RGB ; ColorIndex utilise la palette (xlNone = -4142). Pattern (xlSolid = 1) pour les motifs.',
      'Zellenhintergrund. Color nimmt RGB; ColorIndex nutzt die Palette (xlNone = -4142). Pattern (xlSolid = 1) für Füllmuster.',
    ),
    examples: [
      {
        title: L('Evidenzia A1 in giallo', 'Highlight A1 in yellow', 'Resaltar A1 en amarillo', 'Surligner A1 en jaune', 'A1 gelb hervorheben'),
        code: `Sub DemoInterior()
    Range("A1").Value = "flag"
    Range("A1").Interior.Color = RGB(255, 255, 0)
    Range("A1").Interior.Pattern = xlSolid
    Debug.Print "Color=" & Range("A1").Interior.Color
End Sub`,
        result: L(
          'A1 = flag su sfondo giallo. Immediate Window: Color=65535 (RGB 255,255,0).',
          'A1 = flag on a yellow background. Immediate Window: Color=65535 (RGB 255,255,0).',
          'A1 = flag con fondo amarillo. Immediate Window: Color=65535 (RGB 255,255,0).',
          'A1 = flag sur fond jaune. Immediate Window : Color=65535 (RGB 255,255,0).',
          'A1 = flag auf gelbem Hintergrund. Direktbereich: Color=65535 (RGB 255,255,0).',
        ),
      },
    ],
    related: ['xl-range-font', 'xl-range-borders', 'xl-range-clear'],
  },
  {
    id: 'xl-range-borders',
    name: 'Range.Borders',
    syntax: 'Range.Borders[(xlBorderIndex)].LineStyle / Weight / Color',
    scope: ['excel'],
    category: 'xl-analyze',
    subcategory: 'format',
    description: L(
      'Bordi del range. Senza indice imposti tutti i lati. Indici: xlEdgeLeft (7), xlEdgeTop (8), xlEdgeBottom (9), xlEdgeRight (10), xlInsideHorizontal (12), xlInsideVertical (11).',
      'Range borders. With no index you set every side. Indexes: xlEdgeLeft (7), xlEdgeTop (8), xlEdgeBottom (9), xlEdgeRight (10), xlInsideHorizontal (12), xlInsideVertical (11).',
      'Bordes del rango. Sin índice se aplican todos. Índices: xlEdgeLeft (7), xlEdgeTop (8), xlEdgeBottom (9), xlEdgeRight (10), xlInsideHorizontal (12), xlInsideVertical (11).',
      'Bordures. Sans index, tous les côtés. Index : xlEdgeLeft (7), xlEdgeTop (8), xlEdgeBottom (9), xlEdgeRight (10), xlInsideHorizontal (12), xlInsideVertical (11).',
      'Rahmen. Ohne Index alle Seiten. Indizes: xlEdgeLeft (7), xlEdgeTop (8), xlEdgeBottom (9), xlEdgeRight (10), xlInsideHorizontal (12), xlInsideVertical (11).',
    ),
    examples: [
      {
        title: L('Griglia sottile e fondo spesso', 'Thin grid and thick bottom', 'Cuadrícula fina y fondo grueso', 'Grille fine et bas épais', 'Dünnes Gitter und dicker Boden'),
        code: `Sub DemoBorders()
    Range("A1:B2").Value = 1
    With Range("A1:B2").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
    Range("A1:B2").Borders(xlEdgeBottom).Weight = xlMedium
    Debug.Print "bottom weight=" & Range("A1:B2").Borders(xlEdgeBottom).Weight
End Sub`,
        result: L(
          'A1:B2 con griglia nera; bordo inferiore più spesso (xlMedium = -4138). Immediate Window: bottom weight=-4138.',
          'A1:B2 with a black grid; thicker bottom edge (xlMedium = -4138). Immediate Window: bottom weight=-4138.',
          'A1:B2 con cuadrícula negra; borde inferior más grueso (xlMedium = -4138). Immediate Window: bottom weight=-4138.',
          'A1:B2 avec grille noire ; bord inférieur plus épais (xlMedium = -4138). Immediate Window : bottom weight=-4138.',
          'A1:B2 mit schwarzem Gitter; unterer Rand dicker (xlMedium = -4138). Direktbereich: bottom weight=-4138.',
        ),
      },
    ],
    related: ['xl-range-font', 'xl-range-interior', 'xl-range-merge'],
  },
  {
    id: 'xl-range-merge',
    name: 'Range.Merge / UnMerge',
    syntax: 'Range.Merge [Across]  |  Range.UnMerge  |  Range.MergeCells = Boolean',
    scope: ['excel'],
    category: 'xl-analyze',
    subcategory: 'format',
    description: L(
      'Merge unisce le celle (resta il valore in alto a sinistra). UnMerge le separa. Across:=True unisce riga per riga. Evita Merge in tabelle da analizzare: rompe Sort e Find.',
      'Merge joins cells (the top-left value remains). UnMerge splits them. Across:=True merges row by row. Avoid Merge in data tables: it breaks Sort and Find.',
      'Merge une celdas (queda el valor superior izquierdo). UnMerge las separa. Across:=True une fila a fila. Evita Merge en tablas de datos.',
      'Merge fusionne (la valeur haut-gauche reste). UnMerge sépare. Across:=True fusionne ligne par ligne. Évitez Merge dans les tables de données.',
      'Merge verbindet Zellen (oben links bleibt). UnMerge trennt. Across:=True verbindet zeilenweise. Merge in Datentabellen vermeiden.',
    ),
    examples: [
      {
        title: L('Unisci un titolo e poi separalo', 'Merge a title then unmerge it', 'Combinar un título y luego separar', 'Fusionner un titre puis le défusionner', 'Titel verbinden und wieder trennen'),
        code: `Sub DemoMerge()
    Range("A1").Value = "Q1 report"
    Range("A1:C1").Merge
    Debug.Print "merged=" & Range("A1").MergeCells & " addr=" & Range("A1").MergeArea.Address
    Range("A1").UnMerge
    Debug.Print "after=" & Range("A1").MergeCells
End Sub`,
        result: L(
          'A1:C1 unito con Q1 report, poi di nuovo tre celle. Immediate Window: merged=True addr=$A$1:$C$1 after=False. B1 e C1 vuote.',
          'A1:C1 merged with Q1 report, then three cells again. Immediate Window: merged=True addr=$A$1:$C$1 after=False. B1 and C1 blank.',
          'A1:C1 combinado con Q1 report, luego tres celdas. Immediate Window: merged=True addr=$A$1:$C$1 after=False.',
          'A1:C1 fusionné avec Q1 report, puis trois cellules. Immediate Window : merged=True addr=$A$1:$C$1 after=False.',
          'A1:C1 verbunden mit Q1 report, dann wieder drei Zellen. Direktbereich: merged=True addr=$A$1:$C$1 after=False.',
        ),
      },
    ],
    related: ['xl-range-display', 'xl-range-value', 'xl-range-borders'],
  },
  {
    id: 'xl-range-display',
    name: 'Range.Hidden / WrapText / HorizontalAlignment',
    syntax: 'Range.EntireRow.Hidden  |  Range.WrapText  |  Range.HorizontalAlignment',
    scope: ['excel'],
    category: 'xl-analyze',
    subcategory: 'format',
    description: L(
      'Hidden nasconde riga o colonna intera (non una singola cella). WrapText spezza il testo. HorizontalAlignment: xlLeft (-4131), xlCenter (-4108), xlRight (-4152).',
      'Hidden hides a whole row or column (not a single cell). WrapText wraps text. HorizontalAlignment: xlLeft (-4131), xlCenter (-4108), xlRight (-4152).',
      'Hidden oculta fila o columna entera (no una celda). WrapText ajusta el texto. HorizontalAlignment: xlLeft (-4131), xlCenter (-4108), xlRight (-4152).',
      'Hidden masque une ligne ou colonne entière (pas une cellule). WrapText renvoie à la ligne. HorizontalAlignment : xlLeft (-4131), xlCenter (-4108), xlRight (-4152).',
      'Hidden blendet eine ganze Zeile oder Spalte aus (nicht eine Zelle). WrapText umbricht Text. HorizontalAlignment: xlLeft (-4131), xlCenter (-4108), xlRight (-4152).',
    ),
    examples: [
      {
        title: L('Centra, vai a capo e nascondi la riga 3', 'Center, wrap and hide row 3', 'Centrar, ajustar y ocultar la fila 3', 'Centrer, renvoyer et masquer la ligne 3', 'Zentrieren, umbrechen und Zeile 3 ausblenden'),
        code: `Sub DemoDisplay()
    Range("A1").Value = "Hello world from VBA"
    Range("A1").WrapText = True
    Range("A1").HorizontalAlignment = xlCenter
    Rows(1).AutoFit
    Range("A3").Value = "hidden row"
    Rows(3).Hidden = True
    Debug.Print "align=" & Range("A1").HorizontalAlignment & " wrap=" & Range("A1").WrapText & " r3=" & Rows(3).Hidden
End Sub`,
        result: L(
          'A1 centrato e a capo, riga 1 più alta. Riga 3 nascosta. Immediate Window: align=-4108 wrap=True r3=True.',
          'A1 centered and wrapped, row 1 taller. Row 3 hidden. Immediate Window: align=-4108 wrap=True r3=True.',
          'A1 centrado y ajustado, fila 1 más alta. Fila 3 oculta. Immediate Window: align=-4108 wrap=True r3=True.',
          'A1 centré et renvoyé, ligne 1 plus haute. Ligne 3 masquée. Immediate Window : align=-4108 wrap=True r3=True.',
          'A1 zentriert und umbrochen, Zeile 1 höher. Zeile 3 ausgeblendet. Direktbereich: align=-4108 wrap=True r3=True.',
        ),
      },
    ],
    related: ['xl-range-merge', 'xl-entire-row-column', 'xl-range-autofit'],
  },
  {
    id: 'xl-chartobjects',
    name: 'ChartObjects.Add / ChartType / SetSourceData',
    syntax: 'Worksheet.ChartObjects.Add(Left, Top, Width, Height)  |  Chart.SetSourceData Source  |  Chart.ChartType',
    scope: ['excel'],
    category: 'xl-analyze',
    subcategory: 'charts',
    description: L(
      'ChartObjects.Add crea un grafico incorporato (misure in punti). SetSourceData collega il range. ChartType: xlColumnClustered (51), xlLine (4), xlPie (5).',
      'ChartObjects.Add creates an embedded chart (sizes in points). SetSourceData binds the range. ChartType: xlColumnClustered (51), xlLine (4), xlPie (5).',
      'ChartObjects.Add crea un gráfico incrustado (tamaños en puntos). SetSourceData enlaza el rango. ChartType: xlColumnClustered (51), xlLine (4), xlPie (5).',
      'ChartObjects.Add crée un graphique incorporé (tailles en points). SetSourceData lie la plage. ChartType : xlColumnClustered (51), xlLine (4), xlPie (5).',
      'ChartObjects.Add erzeugt ein eingebettetes Diagramm (Größen in Punkt). SetSourceData bindet den Bereich. ChartType: xlColumnClustered (51), xlLine (4), xlPie (5).',
    ),
    params: [
      { name: 'Left / Top / Width / Height', description: L('Posizione e dimensione in punti.', 'Position and size in points.', 'Posición y tamaño en puntos.', 'Position et taille en points.', 'Position und Größe in Punkt.') },
    ],
    examples: [
      {
        title: L('Istogramma da A1:B3', 'Column chart from A1:B3', 'Gráfico de columnas desde A1:B3', 'Histogramme depuis A1:B3', 'Säulendiagramm aus A1:B3'),
        code: `Sub DemoChart()
    Dim co As ChartObject
    Range("A1").Value = "Name": Range("B1").Value = "Qty"
    Range("A2").Value = "Ann": Range("B2").Value = 4
    Range("A3").Value = "Ben": Range("B3").Value = 7
    On Error Resume Next
    ActiveSheet.ChartObjects("DemoChart").Delete
    On Error GoTo 0
    Set co = ActiveSheet.ChartObjects.Add(Left:=200, Top:=20, Width:=280, Height:=180)
    co.Name = "DemoChart"
    co.Chart.SetSourceData Source:=Range("A1:B3")
    co.Chart.ChartType = xlColumnClustered
    Debug.Print "charts=" & ActiveSheet.ChartObjects.Count & " type=" & co.Chart.ChartType
End Sub`,
        result: L(
          'Compare un istogramma DemoChart a destra dei dati. Immediate Window: charts=1 type=51.',
          'A DemoChart column chart appears to the right of the data. Immediate Window: charts=1 type=51.',
          'Aparece un gráfico de columnas DemoChart a la derecha. Immediate Window: charts=1 type=51.',
          'Un histogramme DemoChart apparaît à droite. Immediate Window : charts=1 type=51.',
          'Ein DemoChart-Säulendiagramm erscheint rechts. Direktbereich: charts=1 type=51.',
        ),
      },
    ],
    related: ['xl-pivottables', 'xl-range-currentregion', 'xl-sheets-vs-worksheets'],
  },
  {
    id: 'xl-pivottables',
    name: 'PivotTables / RefreshTable',
    syntax: 'Workbook.PivotCaches.Create(...)  |  PivotCache.CreatePivotTable(...)  |  PivotTable.RefreshTable',
    scope: ['excel'],
    category: 'xl-analyze',
    subcategory: 'charts',
    description: L(
      'Una PivotTable nasce da un PivotCache (xlDatabase = 1). RefreshTable ricalcola dai dati sorgente. Worksheets("X").PivotTables("Nome") la recupera.',
      'A PivotTable is born from a PivotCache (xlDatabase = 1). RefreshTable recalculates from source data. Worksheets("X").PivotTables("Name") retrieves it.',
      'Una PivotTable nace de un PivotCache (xlDatabase = 1). RefreshTable recalcula desde el origen. Worksheets("X").PivotTables("Nombre") la obtiene.',
      'Une PivotTable naît d’un PivotCache (xlDatabase = 1). RefreshTable recalcule depuis la source. Worksheets("X").PivotTables("Nom") la retrouve.',
      'Eine PivotTable entsteht aus einem PivotCache (xlDatabase = 1). RefreshTable berechnet aus den Quelldaten neu. Worksheets("X").PivotTables("Name") holt sie.',
    ),
    examples: [
      {
        title: L('Crea una pivot e aggiornala', 'Create a pivot and refresh it', 'Crear una dinámica y actualizarla', 'Créer un tableau croisé et l’actualiser', 'Pivot erstellen und aktualisieren'),
        code: `Sub DemoPivot()
    Dim pc As PivotCache, pt As PivotTable
    Range("A1").Value = "Reg": Range("B1").Value = "Amt"
    Range("A2").Value = "West": Range("B2").Value = 10
    Range("A3").Value = "West": Range("B3").Value = 5
    Range("A4").Value = "East": Range("B4").Value = 3
    On Error Resume Next
    ActiveSheet.PivotTables("PivotDemo").TableRange2.Clear
    On Error GoTo 0
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Range("A1:B4"))
    Set pt = pc.CreatePivotTable(TableDestination:=Range("D1"), TableName:="PivotDemo")
    pt.AddFields RowFields:="Reg"
    pt.AddDataField pt.PivotFields("Amt"), "Sum of Amt", xlSum
    Range("B3").Value = 50
    pt.RefreshTable
    Debug.Print "West=" & pt.GetPivotData("Amt", "Reg", "West")
End Sub`,
        result: L(
          'Pivot in D1: East 3, West 60 dopo Refresh (10+50). Immediate Window: West=60. La UI mostra la tabella pivot.',
          'Pivot at D1: East 3, West 60 after Refresh (10+50). Immediate Window: West=60. The UI shows the pivot table.',
          'Dinámica en D1: East 3, West 60 tras Refresh (10+50). Immediate Window: West=60.',
          'Tableau croisé en D1 : East 3, West 60 après Refresh (10+50). Immediate Window : West=60.',
          'Pivot bei D1: East 3, West 60 nach Refresh (10+50). Direktbereich: West=60.',
        ),
      },
    ],
    notes: L(
      'CreatePivotTable fallisce se D1 è occupato o se il nome PivotDemo esiste già: pulisci prima TableRange2.',
      'CreatePivotTable fails if D1 is occupied or PivotDemo already exists: clear TableRange2 first.',
      'CreatePivotTable falla si D1 está ocupado o PivotDemo ya existe: limpia TableRange2 antes.',
      'CreatePivotTable échoue si D1 est occupé ou si PivotDemo existe : effacez TableRange2 d’abord.',
      'CreatePivotTable schlägt fehl, wenn D1 belegt ist oder PivotDemo existiert: zuerst TableRange2 leeren.',
    ),
    related: ['xl-chartobjects', 'xl-range-sort', 'xl-range-currentregion'],
  },
  {
    id: 'xl-workbook-open-event',
    name: 'Workbook_Open / Workbook_BeforeClose',
    syntax: 'Private Sub Workbook_Open()  |  Private Sub Workbook_BeforeClose(Cancel As Boolean)',
    scope: ['excel'],
    category: 'xl-events',
    subcategory: 'events',
    description: L(
      'Eventi del modulo ThisWorkbook (non di un modulo standard). Open parte all’apertura. BeforeClose può impostare Cancel = True per bloccare la chiusura.',
      'Events of the ThisWorkbook module (not a standard module). Open runs on open. BeforeClose can set Cancel = True to block closing.',
      'Eventos del módulo ThisWorkbook (no un módulo estándar). Open corre al abrir. BeforeClose puede poner Cancel = True.',
      'Événements du module ThisWorkbook (pas un module standard). Open s’exécute à l’ouverture. BeforeClose peut mettre Cancel = True.',
      'Ereignisse des ThisWorkbook-Moduls (kein Standardmodul). Open läuft beim Öffnen. BeforeClose kann Cancel = True setzen.',
    ),
    params: [
      { name: 'Cancel', optional: true, description: L('True annulla la chiusura in BeforeClose.', 'True cancels close in BeforeClose.', 'True cancela el cierre en BeforeClose.', 'True annule la fermeture dans BeforeClose.', 'True bricht das Schließen in BeforeClose ab.') },
    ],
    examples: [
      {
        title: L('Timbro in A1 all’apertura (ThisWorkbook)', 'Stamp A1 on open (ThisWorkbook)', 'Sello en A1 al abrir (ThisWorkbook)', 'Horodatage en A1 à l’ouverture (ThisWorkbook)', 'Stempel in A1 beim Öffnen (ThisWorkbook)'),
        code: `' Place in ThisWorkbook, not in a standard module.
Private Sub Workbook_Open()
    Worksheets(1).Range("A1").Value = "opened " & Format(Now, "yyyy-mm-dd hh:nn")
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Worksheets(1).Range("A2").Value = "close attempt"
    ' Cancel = True  ' uncomment to block close
End Sub`,
        result: L(
          'All’apertura A1 = opened 2026-09-05 15:18. Prima di chiudere A2 = close attempt. Se Cancel = True la cartella resta aperta.',
          'On open A1 = opened 2026-09-05 15:18. Before close A2 = close attempt. If Cancel = True the workbook stays open.',
          'Al abrir A1 = opened 2026-09-05 15:18. Antes de cerrar A2 = close attempt. Si Cancel = True el libro sigue abierto.',
          'À l’ouverture A1 = opened 2026-09-05 15:18. Avant fermeture A2 = close attempt. Si Cancel = True le classeur reste ouvert.',
          'Beim Öffnen A1 = opened 2026-09-05 15:18. Vor dem Schließen A2 = close attempt. Bei Cancel = True bleibt die Mappe offen.',
        ),
      },
    ],
    notes: L(
      'Incolla queste Sub in ThisWorkbook. EnableEvents = False le sopprime. UserInterfaceOnly della protezione va riassegnato in Open.',
      'Paste these Subs into ThisWorkbook. EnableEvents = False suppresses them. Protect UserInterfaceOnly must be re-applied in Open.',
      'Pega estas Sub en ThisWorkbook. EnableEvents = False las suprime. UserInterfaceOnly hay que reponerlo en Open.',
      'Collez ces Sub dans ThisWorkbook. EnableEvents = False les coupe. UserInterfaceOnly doit être rétabli dans Open.',
      'Diese Subs in ThisWorkbook einfügen. EnableEvents = False unterdrückt sie. UserInterfaceOnly in Open neu setzen.',
    ),
    related: ['xl-enableevents', 'xl-worksheet-activate-event', 'xl-worksheet-protect'],
  },
  {
    id: 'xl-worksheet-change',
    name: 'Worksheet_Change',
    syntax: 'Private Sub Worksheet_Change(ByVal Target As Range)',
    scope: ['excel'],
    category: 'xl-events',
    subcategory: 'events',
    description: L(
      'Parte dopo ogni modifica di valore/formula sul foglio (manuale o VBA). Target è l’intervallo cambiato. Disabilita EnableEvents prima di scrivere, altrimenti ricorsione.',
      'Fires after every value/formula change on the sheet (manual or VBA). Target is the changed range. Disable EnableEvents before writing, or you recurse.',
      'Se dispara tras cada cambio de valor/fórmula (manual o VBA). Target es el rango cambiado. Desactiva EnableEvents antes de escribir.',
      'Se déclenche après chaque changement de valeur/formule (manuel ou VBA). Target est la plage changée. Coupez EnableEvents avant d’écrire.',
      'Läuft nach jeder Wert-/Formeländerung (manuell oder VBA). Target ist der geänderte Bereich. EnableEvents vor dem Schreiben aus.',
    ),
    params: [
      { name: 'Target', description: L('Celle appena modificate.', 'Cells just changed.', 'Celdas recién cambiadas.', 'Cellules venant d’être changées.', 'Gerade geänderte Zellen.') },
    ],
    examples: [
      {
        title: L('Se cambia A1, timbra B1 (modulo foglio)', 'If A1 changes, stamp B1 (sheet module)', 'Si cambia A1, sella B1 (módulo de hoja)', 'Si A1 change, horodater B1 (module feuille)', 'Wenn A1 sich ändert, B1 stempeln (Blattmodul)'),
        code: `' Place in the worksheet module (right-click the tab → View Code).
Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range("A1")) Is Nothing Then Exit Sub
    Application.EnableEvents = False
    Me.Range("B1").Value = "changed " & Format(Now, "hh:nn:ss")
    Application.EnableEvents = True
End Sub`,
        result: L(
          'Digitando 10 in A1, B1 diventa changed 15:18:02. Modifiche fuori da A1 non toccano B1.',
          'Typing 10 in A1 sets B1 to changed 15:18:02. Edits outside A1 leave B1 alone.',
          'Al escribir 10 en A1, B1 pasa a changed 15:18:02. Cambios fuera de A1 no tocan B1.',
          'Saisir 10 en A1 met B1 à changed 15:18:02. Les éditions hors A1 laissent B1.',
          'Eingabe 10 in A1 setzt B1 auf changed 15:18:02. Änderungen außerhalb A1 lassen B1.',
        ),
      },
    ],
    related: ['xl-enableevents', 'xl-union-intersect', 'xl-worksheet-selectionchange'],
  },
  {
    id: 'xl-worksheet-selectionchange',
    name: 'Worksheet_SelectionChange',
    syntax: 'Private Sub Worksheet_SelectionChange(ByVal Target As Range)',
    scope: ['excel'],
    category: 'xl-events',
    subcategory: 'events',
    description: L(
      'Parte a ogni nuova selezione sul foglio. Target è la selezione. Non scrivere celle senza cautela: può lanciare Change e rallentare la UI.',
      'Fires on every new selection on the sheet. Target is the selection. Do not write cells carelessly: it can fire Change and slow the UI.',
      'Se dispara en cada nueva selección. Target es la selección. No escribas celdas a la ligera: puede lanzar Change y ralentizar la IU.',
      'Se déclenche à chaque nouvelle sélection. Target est la sélection. N’écrivez pas à la légère : cela peut lancer Change et ralentir l’UI.',
      'Läuft bei jeder neuen Auswahl. Target ist die Auswahl. Nicht unbedacht Zellen schreiben: kann Change auslösen und die UI bremsen.',
    ),
    examples: [
      {
        title: L('Mostra l’indirizzo in Z1 (modulo foglio)', 'Show the address in Z1 (sheet module)', 'Mostrar la dirección en Z1 (módulo de hoja)', 'Afficher l’adresse en Z1 (module feuille)', 'Adresse in Z1 zeigen (Blattmodul)'),
        code: `' Place in the worksheet module.
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Application.EnableEvents = False
    Me.Range("Z1").Value = Target.Address(False, False)
    Application.EnableEvents = True
End Sub`,
        result: L(
          'Clic su C5: Z1 = C5. Selezione A1:B2: Z1 = A1:B2. La cella Z1 si aggiorna a ogni clic.',
          'Click C5: Z1 = C5. Select A1:B2: Z1 = A1:B2. Z1 updates on every click.',
          'Clic en C5: Z1 = C5. Selección A1:B2: Z1 = A1:B2. Z1 se actualiza en cada clic.',
          'Clic sur C5 : Z1 = C5. Sélection A1:B2 : Z1 = A1:B2. Z1 se met à jour à chaque clic.',
          'Klick auf C5: Z1 = C5. Auswahl A1:B2: Z1 = A1:B2. Z1 aktualisiert sich bei jedem Klick.',
        ),
      },
    ],
    related: ['xl-worksheet-change', 'xl-cells-range-activecell', 'xl-enableevents'],
  },
  {
    id: 'xl-worksheet-activate-event',
    name: 'Worksheet_Activate',
    syntax: 'Private Sub Worksheet_Activate()',
    scope: ['excel'],
    category: 'xl-events',
    subcategory: 'events',
    description: L(
      'Parte quando l’utente (o VBA) attiva quel foglio. Collocalo nel modulo del foglio. Non parte all’apertura della cartella sul foglio già attivo: usa anche Workbook_Open.',
      'Fires when the user (or VBA) activates that sheet. Put it in the sheet module. It does not fire on workbook open for the already-active sheet: also use Workbook_Open.',
      'Se dispara al activar esa hoja. Ponlo en el módulo de la hoja. No corre al abrir sobre la hoja ya activa: usa también Workbook_Open.',
      'Se déclenche à l’activation de cette feuille. Placez-le dans le module feuille. Ne part pas à l’ouverture sur la feuille déjà active : utilisez aussi Workbook_Open.',
      'Läuft, wenn der Benutzer (oder VBA) dieses Blatt aktiviert. Ins Blattmodul legen. Beim Öffnen der schon aktiven Mappe nicht: zusätzlich Workbook_Open.',
    ),
    examples: [
      {
        title: L('Contatore visite in A1 (modulo foglio)', 'Visit counter in A1 (sheet module)', 'Contador de visitas en A1 (módulo de hoja)', 'Compteur de visites en A1 (module feuille)', 'Besuchszähler in A1 (Blattmodul)'),
        code: `' Place in the worksheet module.
Private Sub Worksheet_Activate()
    Me.Range("A1").Value = CLng(Val(Me.Range("A1").Value)) + 1
    Me.Range("B1").Value = "activated " & Format(Now, "hh:nn:ss")
End Sub`,
        result: L(
          'Ogni volta che passi a quel foglio A1 aumenta di 1 e B1 mostra l’orario. Immediate Window non è usato; l’effetto è sulla griglia.',
          'Each time you switch to that sheet A1 increases by 1 and B1 shows the time. Immediate Window unused; the effect is on the grid.',
          'Cada vez que cambias a esa hoja A1 aumenta 1 y B1 muestra la hora. El efecto está en la cuadrícula.',
          'Chaque fois que vous activez cette feuille A1 augmente de 1 et B1 affiche l’heure. Effet sur la grille.',
          'Jedes Mal, wenn du zu diesem Blatt wechselst, steigt A1 um 1 und B1 zeigt die Uhrzeit. Effekt im Raster.',
        ),
      },
    ],
    related: ['xl-worksheet-activate', 'xl-workbook-open-event', 'xl-enableevents'],
  },
  {
    id: 'xl-pagesetup-print',
    name: 'PageSetup / PrintOut / PrintPreview',
    syntax: 'Worksheet.PageSetup  |  Range.PrintOut  |  Workbook.PrintPreview',
    scope: ['excel'],
    category: 'xl-events',
    subcategory: 'print',
    description: L(
      'PageSetup imposta orientamento, area di stampa, piè di pagina. PrintOut stampa un range o un foglio. PrintPreview apre l’anteprima (Workbook o Worksheet).',
      'PageSetup sets orientation, print area, footer. PrintOut prints a range or sheet. PrintPreview opens the preview (Workbook or Worksheet).',
      'PageSetup define orientación, área de impresión, pie. PrintOut imprime un rango u hoja. PrintPreview abre la vista previa.',
      'PageSetup définit orientation, zone d’impression, pied. PrintOut imprime une plage ou feuille. PrintPreview ouvre l’aperçu.',
      'PageSetup setzt Ausrichtung, Druckbereich, Fußzeile. PrintOut druckt einen Bereich oder ein Blatt. PrintPreview öffnet die Vorschau.',
    ),
    examples: [
      {
        title: L('Landscape su A1:C20 e anteprima', 'Landscape on A1:C20 and preview', 'Horizontal en A1:C20 y vista previa', 'Paysage sur A1:C20 et aperçu', 'Querformat auf A1:C20 und Vorschau'),
        code: `Sub DemoPageSetup()
    Range("A1").Value = "Print demo"
    Range("A2").Value = 123
    With ActiveSheet.PageSetup
        .Orientation = xlLandscape
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintArea = "$A$1:$C$20"
        .CenterFooter = "Page &P of &N"
    End With
    Debug.Print "area=" & ActiveSheet.PageSetup.PrintArea & " orient=" & ActiveSheet.PageSetup.Orientation
    ActiveSheet.PrintPreview
End Sub`,
        result: L(
          'Immediate Window: area=$A$1:$C$20 orient=2 (xlLandscape). Si apre l’anteprima di stampa del foglio. PrintOut invierebbe alla stampante.',
          'Immediate Window: area=$A$1:$C$20 orient=2 (xlLandscape). Print preview of the sheet opens. PrintOut would send to the printer.',
          'Immediate Window: area=$A$1:$C$20 orient=2 (xlLandscape). Se abre la vista previa. PrintOut enviaría a la impresora.',
          'Immediate Window : area=$A$1:$C$20 orient=2 (xlLandscape). L’aperçu s’ouvre. PrintOut enverrait à l’imprimante.',
          'Direktbereich: area=$A$1:$C$20 orient=2 (xlLandscape). Seitenansicht öffnet sich. PrintOut würde an den Drucker senden.',
        ),
      },
    ],
    related: ['xl-range-autofit', 'xl-workbook-save', 'xl-dialogs'],
  },
  {
    id: 'xl-names-add',
    name: 'Names.Add',
    syntax: 'Workbook.Names.Add Name, [RefersTo]  |  Range.Name = String',
    scope: ['excel'],
    category: 'xl-events',
    subcategory: 'print',
    description: L(
      'Crea un nome definito (ambito cartella). RefersTo è una formula inglese ("=Sheet1!$A$1:$B$10" o "=0.22"). Range.Name assegna un nome a un intervallo.',
      'Creates a defined name (workbook scope). RefersTo is an English formula ("=Sheet1!$A$1:$B$10" or "=0.22"). Range.Name assigns a name to a range.',
      'Crea un nombre definido (ámbito de libro). RefersTo es una fórmula en inglés. Range.Name asigna un nombre a un rango.',
      'Crée un nom défini (portée classeur). RefersTo est une formule anglaise. Range.Name affecte un nom à une plage.',
      'Erzeugt einen definierten Namen (Mappenbereich). RefersTo ist eine englische Formel. Range.Name weist einem Bereich einen Namen zu.',
    ),
    params: [
      { name: 'Name', description: L('Identificatore (niente spazi; non come riferimento cella).', 'Identifier (no spaces; not a cell reference).', 'Identificador (sin espacios; no una referencia de celda).', 'Identifiant (pas d’espaces ; pas une référence de cellule).', 'Bezeichner (keine Leerzeichen; keine Zelladresse).') },
      { name: 'RefersTo', optional: true, description: L('Formula =... che il nome rappresenta.', 'Formula =... that the name represents.', 'Fórmula =... que representa el nombre.', 'Formule =... que le nom représente.', 'Formel =..., die der Name darstellt.') },
    ],
    examples: [
      {
        title: L('Nome VAT e nome di range', 'VAT name and a range name', 'Nombre VAT y nombre de rango', 'Nom VAT et nom de plage', 'Name VAT und Bereichsname'),
        code: `Sub DemoNamesAdd()
    On Error Resume Next
    ThisWorkbook.Names("VAT").Delete
    ThisWorkbook.Names("SalesBlock").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="VAT", RefersTo:="=0.22"
    Range("A1").Value = 100
    Range("B1").Formula = "=A1*VAT"
    Range("A1:B1").Name = "SalesBlock"
    Debug.Print "VAT=" & Evaluate("VAT") & " B1=" & Range("B1").Value & " " & Range("SalesBlock").Address
End Sub`,
        result: L(
          'B1 = 22. Gestione nomi elenca VAT e SalesBlock. Immediate Window: VAT=0.22 B1=22 $A$1:$B$1.',
          'B1 = 22. Name Manager lists VAT and SalesBlock. Immediate Window: VAT=0.22 B1=22 $A$1:$B$1.',
          'B1 = 22. El Administrador de nombres lista VAT y SalesBlock. Immediate Window: VAT=0.22 B1=22 $A$1:$B$1.',
          'B1 = 22. Le Gestionnaire de noms liste VAT et SalesBlock. Immediate Window : VAT=0.22 B1=22 $A$1:$B$1.',
          'B1 = 22. Namens-Manager listet VAT und SalesBlock. Direktbereich: VAT=0.22 B1=22 $A$1:$B$1.',
        ),
      },
    ],
    related: ['xl-cells-range-activecell', 'xl-range-address-row-col', 'xl-range-formula'],
  },
  {
    id: 'xl-hyperlinks-add',
    name: 'Hyperlinks.Add',
    syntax: 'Worksheet.Hyperlinks.Add Anchor, Address, [SubAddress], [ScreenTip], [TextToDisplay]',
    scope: ['excel'],
    category: 'xl-events',
    subcategory: 'print',
    description: L(
      'Aggiunge un collegamento a una cella. Address è URL o file; SubAddress è un segnalibro interno ("\'Sheet2\'!A1"). TextToDisplay è il testo visibile.',
      'Adds a link to a cell. Address is a URL or file; SubAddress is an internal bookmark ("\'Sheet2\'!A1"). TextToDisplay is the visible text.',
      'Añade un vínculo a una celda. Address es URL o archivo; SubAddress es un marcador interno. TextToDisplay es el texto visible.',
      'Ajoute un lien à une cellule. Address est une URL ou un fichier ; SubAddress est un signet interne. TextToDisplay est le texte visible.',
      'Fügt einer Zelle einen Link hinzu. Address ist URL oder Datei; SubAddress ist ein internes Lesezeichen. TextToDisplay ist der sichtbare Text.',
    ),
    params: [
      { name: 'Anchor', description: L('Range cliccabile.', 'Clickable range.', 'Rango clicable.', 'Plage cliquable.', 'Anklickbarer Bereich.') },
      { name: 'Address', description: L('URL, percorso file, o "" se solo SubAddress.', 'URL, file path, or "" for SubAddress only.', 'URL, ruta, o "" si solo SubAddress.', 'URL, chemin, ou "" si SubAddress seul.', 'URL, Dateipfad oder "" nur mit SubAddress.') },
      { name: 'SubAddress', optional: true, description: L('Riferimento interno foglio!cella.', 'Internal sheet!cell reference.', 'Referencia interna hoja!celda.', 'Référence interne feuille!cellule.', 'Interne Blatt!Zelle-Referenz.') },
    ],
    examples: [
      {
        title: L('Link web e salto interno', 'Web link and internal jump', 'Vínculo web y salto interno', 'Lien web et saut interne', 'Weblink und interner Sprung'),
        code: `Sub DemoHyperlinks()
    ActiveSheet.Hyperlinks.Add Anchor:=Range("A1"), _
        Address:="https://example.com", TextToDisplay:="Open site"
    ActiveSheet.Hyperlinks.Add Anchor:=Range("A2"), _
        Address:="", SubAddress:="'" & Worksheets(1).Name & "'!A1", _
        TextToDisplay:="Go to A1"
    Debug.Print "links=" & ActiveSheet.Hyperlinks.Count & " A1=" & Range("A1").Value
End Sub`,
        result: L(
          'A1 = Open site (blu, sottolineato) verso example.com. A2 = Go to A1 salta a A1. Immediate Window: links=2 A1=Open site.',
          'A1 = Open site (blue, underlined) to example.com. A2 = Go to A1 jumps to A1. Immediate Window: links=2 A1=Open site.',
          'A1 = Open site (azul, subrayado) a example.com. A2 = Go to A1 salta a A1. Immediate Window: links=2 A1=Open site.',
          'A1 = Open site (bleu, souligné) vers example.com. A2 = Go to A1 saute vers A1. Immediate Window : links=2 A1=Open site.',
          'A1 = Open site (blau, unterstrichen) zu example.com. A2 = Go to A1 springt zu A1. Direktbereich: links=2 A1=Open site.',
        ),
      },
    ],
    related: ['xl-names-add', 'xl-range-value', 'xl-worksheet-activate'],
  },
  {
    id: 'xl-comments',
    name: 'Comments / Notes',
    syntax: 'Range.AddComment [Text]  |  Range.Comment.Text  |  Range.AddCommentThreaded [Text]',
    scope: ['excel'],
    category: 'xl-events',
    subcategory: 'print',
    description: L(
      'Le Note classiche sono l’oggetto Comment (triangolino rosso). AddCommentThreaded crea un commento moderno conversazionale (Excel 365). Delete / Comment.Delete rimuove la nota.',
      'Classic Notes are the Comment object (red triangle). AddCommentThreaded creates a modern threaded comment (Excel 365). Delete / Comment.Delete removes the note.',
      'Las Notas clásicas son el objeto Comment (triángulo rojo). AddCommentThreaded crea un comentario moderno (Excel 365). Delete quita la nota.',
      'Les Notes classiques sont l’objet Comment (triangle rouge). AddCommentThreaded crée un commentaire moderne (Excel 365). Delete ôte la note.',
      'Klassische Notizen sind das Comment-Objekt (rotes Dreieck). AddCommentThreaded erzeugt einen modernen Kommentar (Excel 365). Delete entfernt die Notiz.',
    ),
    examples: [
      {
        title: L('Aggiungi una nota classica su A1', 'Add a classic note on A1', 'Añadir una nota clásica en A1', 'Ajouter une note classique sur A1', 'Klassische Notiz auf A1 setzen'),
        code: `Sub DemoComment()
    Range("A1").Value = 15
    On Error Resume Next
    Range("A1").Comment.Delete
    On Error GoTo 0
    Range("A1").AddComment "Check this value"
    Range("A1").Comment.Visible = True
    Debug.Print Range("A1").Comment.Text
End Sub`,
        result: L(
          'A1 = 15 con nota visibile “Check this value” e triangolino. Immediate Window: Check this value.',
          'A1 = 15 with a visible note “Check this value” and a red triangle. Immediate Window: Check this value.',
          'A1 = 15 con nota visible “Check this value” y triángulo. Immediate Window: Check this value.',
          'A1 = 15 avec note visible « Check this value » et triangle. Immediate Window : Check this value.',
          'A1 = 15 mit sichtbarer Notiz „Check this value“ und Dreieck. Direktbereich: Check this value.',
        ),
      },
    ],
    notes: L(
      'AddComment fallisce se la cella ha già una nota: eliminala prima. AddCommentThreaded richiede un build Microsoft 365 recente.',
      'AddComment fails if the cell already has a note: delete it first. AddCommentThreaded needs a recent Microsoft 365 build.',
      'AddComment falla si ya hay una nota: bórrala antes. AddCommentThreaded requiere Microsoft 365 reciente.',
      'AddComment échoue s’il y a déjà une note : supprimez-la d’abord. AddCommentThreaded exige Microsoft 365 récent.',
      'AddComment schlägt fehl, wenn schon eine Notiz existiert: zuerst löschen. AddCommentThreaded braucht aktuelles Microsoft 365.',
    ),
    related: ['xl-range-value', 'xl-range-font', 'xl-range-clear'],
  },
]