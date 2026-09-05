import { L, type Command } from '../types'

export const accessCommands: Command[] = [
  {
    id: 'ac-current-db',
    name: 'CurrentDb',
    syntax: 'CurrentDb As DAO.Database',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'current',
    description: L(
      'Restituisce l’oggetto DAO.Database del file .accdb/.mdb aperto.',
      'Returns the DAO.Database object of the open .accdb/.mdb file.',
      'Devuelve el objeto DAO.Database del archivo .accdb/.mdb abierto.',
      'Renvoie l’objet DAO.Database du fichier .accdb/.mdb ouvert.',
      'Gibt das DAO.Database-Objekt der geöffneten .accdb/.mdb-Datei zurück.',
    ),
    examples: [
      {
        title: L('Nome file e numero tabelle', 'File name and table count', 'Nombre de archivo y número de tablas', 'Nom du fichier et nombre de tables', 'Dateiname und Tabellenanzahl'),
        code: `Sub EsempioCurrentDb()
    Dim db As DAO.Database
    Set db = CurrentDb
    Debug.Print db.Name
    Debug.Print db.TableDefs.Count
End Sub`,
        result: L(
          'Finestra Immediata: C:\\Dati\\Vendite.accdb  poi  14',
          'Immediate Window: C:\\Dati\\Vendite.accdb  then  14',
          'Ventana Inmediato: C:\\Dati\\Vendite.accdb  luego  14',
          'Fenêtre Exécution : C:\\Dati\\Vendite.accdb  puis  14',
          'Direktbereich: C:\\Dati\\Vendite.accdb  dann  14',
        ),
      },
    ],
    notes: L(
      'Preferisci CurrentDb a DBEngine(0)(0): è l’istanza cached del database corrente.',
      'Prefer CurrentDb over DBEngine(0)(0): it is the cached instance of the current database.',
      'Prefiere CurrentDb a DBEngine(0)(0): es la instancia en caché de la base actual.',
      'Préférez CurrentDb à DBEngine(0)(0) : c’est l’instance en cache de la base courante.',
      'CurrentDb ist DBEngine(0)(0) vorzuziehen: es ist die zwischengespeicherte Instanz der aktuellen Datenbank.',
    ),
    related: ['ac-dbengine', 'ac-open-database', 'ac-dao-open-recordset', 'ac-dao-querydef-tabledef'],
  },
  {
    id: 'ac-dbengine',
    name: 'DBEngine',
    syntax: 'DBEngine As DAO.DBEngine',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'current',
    description: L(
      'Motore DAO: workspace, database aperti e operazioni d’insieme.',
      'DAO engine: workspaces, open databases and engine-wide operations.',
      'Motor DAO: workspaces, bases abiertas y operaciones globales.',
      'Moteur DAO : workspaces, bases ouvertes et opérations globales.',
      'DAO-Engine: Workspaces, geöffnete Datenbanken und engineweite Operationen.',
    ),
    examples: [
      {
        title: L('Workspace predefinito e Idle', 'Default workspace and Idle', 'Workspace predeterminado e Idle', 'Workspace par défaut et Idle', 'Standard-Workspace und Idle'),
        code: `Sub EsempioDBEngine()
    Debug.Print DBEngine.Version
    Debug.Print DBEngine.Workspaces(0).Name
    DBEngine.Idle dbRefreshCache
    Debug.Print "cache aggiornata"
End Sub`,
        result: L(
          'Finestra Immediata: 14.0  poi  #Default Workspace#  poi  cache aggiornata',
          'Immediate Window: 14.0  then  #Default Workspace#  then  cache aggiornata',
          'Ventana Inmediato: 14.0  luego  #Default Workspace#  luego  cache aggiornata',
          'Fenêtre Exécution : 14.0  puis  #Default Workspace#  puis  cache aggiornata',
          'Direktbereich: 14.0  dann  #Default Workspace#  dann  cache aggiornata',
        ),
      },
    ],
    notes: L(
      'DBEngine(0)(0) è il primo database del primo workspace; in Access coincide di solito con CurrentDb ma non è la stessa referenza cached.',
      'DBEngine(0)(0) is the first database of the first workspace; in Access it usually matches CurrentDb but is not the same cached reference.',
      'DBEngine(0)(0) es la primera base del primer workspace; en Access suele coincidir con CurrentDb pero no es la misma referencia en caché.',
      'DBEngine(0)(0) est la première base du premier workspace ; dans Access elle correspond souvent à CurrentDb sans être la même référence en cache.',
      'DBEngine(0)(0) ist die erste Datenbank des ersten Workspace; in Access entspricht sie meist CurrentDb, ist aber nicht dieselbe Cache-Referenz.',
    ),
    related: ['ac-current-db', 'ac-open-database'],
  },
  {
    id: 'ac-open-database',
    name: 'OpenDatabase',
    syntax: 'DBEngine.OpenDatabase(Name, [Options], [ReadOnly], [Connect]) As DAO.Database',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'current',
    description: L(
      'Apre un altro file Access (o un’origine ODBC) come DAO.Database.',
      'Opens another Access file (or an ODBC source) as a DAO.Database.',
      'Abre otro archivo Access (o un origen ODBC) como DAO.Database.',
      'Ouvre un autre fichier Access (ou une source ODBC) en tant que DAO.Database.',
      'Öffnet eine andere Access-Datei (oder eine ODBC-Quelle) als DAO.Database.',
    ),
    params: [
      { name: 'Name', description: L('Percorso .accdb/.mdb o stringa ODBC.', 'Path to .accdb/.mdb or ODBC string.', 'Ruta .accdb/.mdb o cadena ODBC.', 'Chemin .accdb/.mdb ou chaîne ODBC.', 'Pfad zur .accdb/.mdb oder ODBC-Zeichenfolge.') },
      { name: 'Options', optional: true, description: L('True = esclusivo.', 'True = exclusive.', 'True = exclusivo.', 'True = exclusif.', 'True = exklusiv.') },
      { name: 'ReadOnly', optional: true, description: L('True = sola lettura.', 'True = read-only.', 'True = solo lectura.', 'True = lecture seule.', 'True = schreibgeschützt.') },
      { name: 'Connect', optional: true, description: L('Parametri di connessione (password, ODBC).', 'Connection parameters (password, ODBC).', 'Parámetros de conexión (contraseña, ODBC).', 'Paramètres de connexion (mot de passe, ODBC).', 'Verbindungsparameter (Kennwort, ODBC).') },
    ],
    examples: [
      {
        title: L('Leggere una tabella da un secondo .accdb', 'Read a table from a second .accdb', 'Leer una tabla de un segundo .accdb', 'Lire une table d’un second .accdb', 'Tabelleende aus einer zweiten .accdb lesen'),
        code: `Sub EsempioOpenDatabase()
    Dim dbEsterno As DAO.Database
    Dim rs As DAO.Recordset
    Set dbEsterno = DBEngine.OpenDatabase("C:\\Dati\\Archivio.accdb")
    Set rs = dbEsterno.OpenRecordset("Clienti")
    Debug.Print rs!RagioneSociale
    rs.Close
    dbEsterno.Close
End Sub`,
        result: L(
          'Finestra Immediata: Bianchi Srl  (primo record di Archivio.accdb / Clienti)',
          'Immediate Window: Bianchi Srl  (first record of Archivio.accdb / Clienti)',
          'Ventana Inmediato: Bianchi Srl  (primer registro de Archivio.accdb / Clienti)',
          'Fenêtre Exécution : Bianchi Srl  (premier enregistrement de Archivio.accdb / Clienti)',
          'Direktbereich: Bianchi Srl  (erster Datensatz von Archivio.accdb / Clienti)',
        ),
      },
    ],
    notes: L(
      'Chiudi sempre il Database esterno: non è CurrentDb e resta bloccato finché non chiami Close.',
      'Always close the external Database: it is not CurrentDb and stays locked until you call Close.',
      'Cierra siempre la base externa: no es CurrentDb y permanece bloqueada hasta Close.',
      'Fermez toujours la base externe : ce n’est pas CurrentDb et elle reste verrouillée jusqu’à Close.',
      'Schließen Sie die externe Datenbank immer: sie ist nicht CurrentDb und bleibt bis Close gesperrt.',
    ),
    related: ['ac-current-db', 'ac-dbengine', 'ac-dao-open-recordset'],
  },
  {
    id: 'ac-current-project',
    name: 'CurrentProject',
    syntax: 'CurrentProject As CurrentProject',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'current',
    description: L(
      'Progetto Access: percorso, nome, elenchi di maschere/report e stato di attendibilità.',
      'Access project: path, name, form/report lists and trust state.',
      'Proyecto Access: ruta, nombre, listas de formularios/informes y confianza.',
      'Projet Access : chemin, nom, listes de formulaires/états et état de confiance.',
      'Access-Projekt: Pfad, Name, Formular-/Berichtslisten und Vertrauensstatus.',
    ),
    examples: [
      {
        title: L('Percorso, nome e conteggio maschere', 'Path, name and form count', 'Ruta, nombre y recuento de formularios', 'Chemin, nom et nombre de formulaires', 'Pfad, Name und Formularanzahl'),
        code: `Sub EsempioCurrentProject()
    Debug.Print CurrentProject.Name
    Debug.Print CurrentProject.Path
    Debug.Print CurrentProject.AllForms.Count
    Debug.Print CurrentProject.IsTrusted
End Sub`,
        result: L(
          'Finestra Immediata: Vendite.accdb  |  C:\\Dati  |  6  |  True',
          'Immediate Window: Vendite.accdb  |  C:\\Dati  |  6  |  True',
          'Ventana Inmediato: Vendite.accdb  |  C:\\Dati  |  6  |  True',
          'Fenêtre Exécution : Vendite.accdb  |  C:\\Dati  |  6  |  True',
          'Direktbereich: Vendite.accdb  |  C:\\Dati  |  6  |  True',
        ),
      },
    ],
    related: ['ac-current-db', 'ac-current-user', 'ac-code-context-object'],
  },
  {
    id: 'ac-current-user',
    name: 'CurrentUser',
    syntax: 'CurrentUser() As String',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'current',
    description: L(
      'Nome dell’utente Access collegato (sicurezza a livello utente).',
      'Name of the signed-in Access user (user-level security).',
      'Nombre del usuario de Access conectado (seguridad por usuario).',
      'Nom de l’utilisateur Access connecté (sécurité par utilisateur).',
      'Name des angemeldeten Access-Benutzers (benutzerbezogene Sicherheit).',
    ),
    examples: [
      {
        title: L('Stampare l’utente corrente', 'Print the current user', 'Imprimir el usuario actual', 'Afficher l’utilisateur courant', 'Aktuellen Benutzer ausgeben'),
        code: `Sub EsempioCurrentUser()
    Debug.Print CurrentUser()
    If CurrentUser() = "Admin" Then
        Debug.Print "account predefinito"
    End If
End Sub`,
        result: L(
          'Finestra Immediata: Admin  poi  account predefinito  (file senza sicurezza utenti)',
          'Immediate Window: Admin  then  account predefinito  (file without user-level security)',
          'Ventana Inmediato: Admin  luego  account predefinito  (archivo sin seguridad de usuarios)',
          'Fenêtre Exécution : Admin  puis  account predefinito  (fichier sans sécurité utilisateurs)',
          'Direktbereich: Admin  dann  account predefinito  (Datei ohne Benutzersicherheit)',
        ),
      },
    ],
    notes: L(
      'Nei file .accdb moderni senza workgroup il valore è quasi sempre Admin. Per il login Windows usa Environ("USERNAME").',
      'In modern .accdb files without a workgroup the value is almost always Admin. For the Windows login use Environ("USERNAME").',
      'En archivos .accdb modernos sin workgroup el valor es casi siempre Admin. Para el inicio de Windows usa Environ("USERNAME").',
      'Dans les .accdb modernes sans workgroup la valeur est presque toujours Admin. Pour le login Windows utilisez Environ("USERNAME").',
      'In modernen .accdb-Dateien ohne Workgroup ist der Wert fast immer Admin. Für die Windows-Anmeldung Environ("USERNAME") verwenden.',
    ),
    related: ['ac-current-project', 'ac-current-db'],
  },
  {
    id: 'ac-code-context-object',
    name: 'Application.CodeContextObject',
    syntax: 'Application.CodeContextObject As Object',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'current',
    description: L(
      'Oggetto (maschera, report o modulo) in cui è in esecuzione il codice.',
      'Object (form, report or module) in which the code is running.',
      'Objeto (formulario, informe o módulo) donde se ejecuta el código.',
      'Objet (formulaire, état ou module) dans lequel le code s’exécute.',
      'Objekt (Formular, Bericht oder Modul), in dem der Code läuft.',
    ),
    examples: [
      {
        title: L('Routine condivisa che legge il chiamante', 'Shared routine that reads the caller', 'Rutina compartida que lee al llamador', 'Routine partagée qui lit l’appelant', 'Gemeinsame Routine, die den Aufrufer liest'),
        code: `Sub LogContesto()
    Dim ctx As Object
    Set ctx = Application.CodeContextObject
    Debug.Print TypeName(ctx)
    Debug.Print ctx.Name
End Sub

' Da modulo della maschera frmClienti:
Private Sub Form_Load()
    LogContesto
End Sub`,
        result: L(
          'Finestra Immediata: Form  poi  frmClienti',
          'Immediate Window: Form  then  frmClienti',
          'Ventana Inmediato: Form  luego  frmClienti',
          'Fenêtre Exécution : Form  puis  frmClienti',
          'Direktbereich: Form  dann  frmClienti',
        ),
      },
    ],
    notes: L(
      'A differenza di Me, funziona anche da una routine pubblica in un modulo standard chiamata dalla maschera.',
      'Unlike Me, this also works from a public routine in a standard module called by the form.',
      'A diferencia de Me, también funciona desde una rutina pública en un módulo estándar llamada por el formulario.',
      'Contrairement à Me, cela fonctionne aussi depuis une routine publique d’un module standard appelée par le formulaire.',
      'Im Gegensatz zu Me funktioniert das auch aus einer öffentlichen Routine in einem Standardmodul, die das Formular aufruft.',
    ),
    related: ['ac-me', 'ac-current-project', 'ac-screen-active'],
  },
  {
    id: 'ac-nz',
    name: 'Nz',
    syntax: 'Nz(Value, [ValueIfNull])',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'current',
    description: L(
      'Sostituisce Null con un valore di default (0, stringa vuota o altro).',
      'Replaces Null with a default value (0, empty string or other).',
      'Sustituye Null por un valor predeterminado (0, cadena vacía u otro).',
      'Remplace Null par une valeur par défaut (0, chaîne vide ou autre).',
      'Ersetzt Null durch einen Standardwert (0, leere Zeichenfolge oder anderes).',
    ),
    params: [
      { name: 'Value', description: L('Espressione che può essere Null.', 'Expression that may be Null.', 'Expresión que puede ser Null.', 'Expression pouvant être Null.', 'Ausdruck, der Null sein kann.') },
      { name: 'ValueIfNull', optional: true, description: L('Sostituto se Value è Null. Default: 0 o "".', 'Replacement if Value is Null. Default: 0 or "".', 'Sustituto si Value es Null. Predeterminado: 0 o "".', 'Remplacement si Value est Null. Défaut : 0 ou "".', 'Ersatz, wenn Value Null ist. Standard: 0 oder "".') },
    ],
    examples: [
      {
        title: L('DLookup che può non trovare la riga', 'DLookup that may find no row', 'DLookup que puede no hallar la fila', 'DLookup qui peut ne trouver aucune ligne', 'DLookup, das keine Zeile findet'),
        code: `Sub EsempioNz()
    Dim tot As Currency
    tot = Nz(DLookup("Sconto", "Clienti", "Codice=104"), 0)
    Debug.Print tot
    Debug.Print Nz(Null, "n.d.")
End Sub`,
        result: L(
          'Finestra Immediata: 0  poi  n.d.  (nessun cliente 104, secondo Nz restituisce "n.d.")',
          'Immediate Window: 0  then  n.d.  (no customer 104; second Nz returns "n.d.")',
          'Ventana Inmediato: 0  luego  n.d.  (no hay cliente 104; el segundo Nz devuelve "n.d.")',
          'Fenêtre Exécution : 0  puis  n.d.  (pas de client 104 ; le second Nz renvoie "n.d.")',
          'Direktbereich: 0  dann  n.d.  (kein Kunde 104; zweites Nz liefert "n.d.")',
        ),
      },
    ],
    notes: L(
      'Senza secondo argomento Nz(Null) su un contesto numerico diventa 0, su testo diventa stringa vuota.',
      'Without a second argument Nz(Null) becomes 0 in a numeric context and an empty string in a text context.',
      'Sin segundo argumento Nz(Null) es 0 en contexto numérico y cadena vacía en texto.',
      'Sans second argument Nz(Null) devient 0 en contexte numérique et une chaîne vide en texte.',
      'Ohne zweites Argument wird Nz(Null) im Zahlenkontext 0 und im Textkontext eine leere Zeichenfolge.',
    ),
    related: ['ac-dlookup', 'ac-eval', 'ac-dsum'],
  },
  {
    id: 'ac-eval',
    name: 'Eval',
    syntax: 'Eval(StringExpr)',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'current',
    description: L(
      'Valuta una stringa come espressione Access (funzioni di dominio, operatori, date).',
      'Evaluates a string as an Access expression (domain functions, operators, dates).',
      'Evalúa una cadena como expresión de Access (dominio, operadores, fechas).',
      'Évalue une chaîne comme expression Access (domaine, opérateurs, dates).',
      'Wertet eine Zeichenfolge als Access-Ausdruck aus (Domänenfunktionen, Operatoren, Datum).',
    ),
    params: [
      { name: 'StringExpr', description: L('Espressione in testo, come nella finestra Criteri.', 'Text expression, as in a Criteria box.', 'Expresión en texto, como en Criterios.', 'Expression texte, comme dans Critères.', 'Textausdruck, wie im Kriterienfeld.') },
    ],
    examples: [
      {
        title: L('Calcolo e funzione di dominio da stringa', 'Calculation and domain function from a string', 'Cálculo y función de dominio desde una cadena', 'Calcul et fonction de domaine depuis une chaîne', 'Berechnung und Domänenfunktion aus einer Zeichenfolge'),
        code: `Sub EsempioEval()
    Debug.Print Eval("2+2")
    Debug.Print Eval("Date()")
    Debug.Print Eval("DCount(\"*\",\"Clienti\")")
End Sub`,
        result: L(
          'Finestra Immediata: 4  poi  05/09/2026  poi  128',
          'Immediate Window: 4  then  9/5/2026  then  128',
          'Ventana Inmediato: 4  luego  05/09/2026  luego  128',
          'Fenêtre Exécution : 4  puis  05/09/2026  puis  128',
          'Direktbereich: 4  dann  05.09.2026  dann  128',
        ),
      },
    ],
    notes: L(
      'Le virgolette interne si raddoppiano. Non passare input utente non filtrato: Eval esegue l’espressione.',
      'Double inner quotes. Do not pass unfiltered user input: Eval runs the expression.',
      'Duplica las comillas internas. No pases entrada de usuario sin filtrar: Eval ejecuta la expresión.',
      'Doublez les guillemets internes. Ne passez pas d’entrée utilisateur non filtrée : Eval exécute l’expression.',
      'Innere Anführungszeichen verdoppeln. Keine ungefilterte Benutzereingabe übergeben: Eval führt den Ausdruck aus.',
    ),
    related: ['ac-nz', 'ac-dcount', 'ac-dlookup'],
  },
  {
    id: 'ac-docmd-open-form',
    name: 'DoCmd.OpenForm',
    syntax: 'DoCmd.OpenForm FormName, [View], [FilterName], [WhereCondition], [DataMode], [WindowMode], [OpenArgs]',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'docmd',
    description: L(
      'Apre una maschera in visualizzazione, filtro e modalità dati scelti.',
      'Opens a form in the chosen view, filter and data mode.',
      'Abre un formulario en la vista, filtro y modo de datos elegidos.',
      'Ouvre un formulaire dans la vue, le filtre et le mode de données choisis.',
      'Öffnet ein Formular in der gewählten Ansicht, Filterung und Datenmodus.',
    ),
    params: [
      { name: 'FormName', description: L('Nome della maschera.', 'Form name.', 'Nombre del formulario.', 'Nom du formulaire.', 'Formularname.') },
      { name: 'View', optional: true, description: L('acNormal, acDesign, acPreview, acFormDS, acLayout.', 'acNormal, acDesign, acPreview, acFormDS, acLayout.', 'acNormal, acDesign, acPreview, acFormDS, acLayout.', 'acNormal, acDesign, acPreview, acFormDS, acLayout.', 'acNormal, acDesign, acPreview, acFormDS, acLayout.') },
      { name: 'WhereCondition', optional: true, description: L('Clausola WHERE senza la parola WHERE.', 'WHERE clause without the word WHERE.', 'Cláusula WHERE sin la palabra WHERE.', 'Clause WHERE sans le mot WHERE.', 'WHERE-Klausel ohne das Wort WHERE.') },
      { name: 'DataMode', optional: true, description: L('acFormAdd, acFormEdit, acFormReadOnly.', 'acFormAdd, acFormEdit, acFormReadOnly.', 'acFormAdd, acFormEdit, acFormReadOnly.', 'acFormAdd, acFormEdit, acFormReadOnly.', 'acFormAdd, acFormEdit, acFormReadOnly.') },
      { name: 'WindowMode', optional: true, description: L('acWindowNormal, acDialog, acHidden, acIcon.', 'acWindowNormal, acDialog, acHidden, acIcon.', 'acWindowNormal, acDialog, acHidden, acIcon.', 'acWindowNormal, acDialog, acHidden, acIcon.', 'acWindowNormal, acDialog, acHidden, acIcon.') },
      { name: 'OpenArgs', optional: true, description: L('Stringa letta da Me.OpenArgs nella maschera.', 'String read from Me.OpenArgs in the form.', 'Cadena leída en Me.OpenArgs del formulario.', 'Chaîne lue via Me.OpenArgs dans le formulaire.', 'Zeichenfolge, die das Formular in Me.OpenArgs liest.') },
    ],
    examples: [
      {
        title: L('Aprire frmClienti sul codice 12', 'Open frmClienti on code 12', 'Abrir frmClienti en el código 12', 'Ouvrir frmClienti sur le code 12', 'frmClienti für Code 12 öffnen'),
        code: `Sub ApriCliente12()
    DoCmd.OpenForm "frmClienti", acNormal, , "Codice=12", acFormEdit, acWindowNormal, "daMenu"
End Sub`,
        result: L(
          'Si apre frmClienti sul record Codice=12 in modifica. Me.OpenArgs = "daMenu".',
          'frmClienti opens on the Codice=12 record in edit mode. Me.OpenArgs = "daMenu".',
          'Se abre frmClienti en el registro Codice=12 en edición. Me.OpenArgs = "daMenu".',
          'frmClienti s’ouvre sur l’enregistrement Codice=12 en modification. Me.OpenArgs = "daMenu".',
          'frmClienti öffnet den Datensatz Codice=12 im Bearbeiten-Modus. Me.OpenArgs = "daMenu".',
        ),
      },
    ],
    related: ['ac-docmd-close', 'ac-open-args', 'ac-forms-bang', 'ac-me', 'ac-docmd-open-report'],
  },
  {
    id: 'ac-docmd-close',
    name: 'DoCmd.Close',
    syntax: 'DoCmd.Close [ObjectType], [ObjectName], [Save]',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'docmd',
    description: L(
      'Chiude una maschera, un report o un altro oggetto Access.',
      'Closes a form, report or another Access object.',
      'Cierra un formulario, un informe u otro objeto de Access.',
      'Ferme un formulaire, un état ou un autre objet Access.',
      'Schließt ein Formular, einen Bericht oder ein anderes Access-Objekt.',
    ),
    params: [
      { name: 'ObjectType', optional: true, description: L('acForm, acReport, acTable, acQuery, acMacro, acModule.', 'acForm, acReport, acTable, acQuery, acMacro, acModule.', 'acForm, acReport, acTable, acQuery, acMacro, acModule.', 'acForm, acReport, acTable, acQuery, acMacro, acModule.', 'acForm, acReport, acTable, acQuery, acMacro, acModule.') },
      { name: 'ObjectName', optional: true, description: L('Nome oggetto. Omesso: chiude l’oggetto attivo.', 'Object name. Omitted: closes the active object.', 'Nombre del objeto. Omitido: cierra el objeto activo.', 'Nom de l’objet. Omise : ferme l’objet actif.', 'Objektname. Weggelassen: schließt das aktive Objekt.') },
      { name: 'Save', optional: true, description: L('acSaveYes, acSaveNo, acSavePrompt.', 'acSaveYes, acSaveNo, acSavePrompt.', 'acSaveYes, acSaveNo, acSavePrompt.', 'acSaveYes, acSaveNo, acSavePrompt.', 'acSaveYes, acSaveNo, acSavePrompt.') },
    ],
    examples: [
      {
        title: L('Chiudere la maschera senza salvare il layout', 'Close the form without saving layout', 'Cerrar el formulario sin guardar el diseño', 'Fermer le formulaire sans enregistrer la disposition', 'Formular schließen ohne Layout zu speichern'),
        code: `Sub ChiudiClienti()
    DoCmd.Close acForm, "frmClienti", acSaveNo
End Sub`,
        result: L(
          'frmClienti scompare. Eventuali modifiche ai dati già salvate restano; il layout non viene scritto.',
          'frmClienti disappears. Already-saved data edits remain; the layout is not written.',
          'frmClienti desaparece. Los datos ya guardados permanecen; el diseño no se escribe.',
          'frmClienti disparaît. Les données déjà enregistrées restent ; la disposition n’est pas écrite.',
          'frmClienti verschwindet. Bereits gespeicherte Daten bleiben; das Layout wird nicht geschrieben.',
        ),
      },
    ],
    related: ['ac-docmd-open-form', 'ac-docmd-quit', 'ac-me-undo-dirty'],
  },
  {
    id: 'ac-docmd-open-report',
    name: 'DoCmd.OpenReport',
    syntax: 'DoCmd.OpenReport ReportName, [View], [FilterName], [WhereCondition], [WindowMode], [OpenArgs]',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'docmd',
    description: L(
      'Apre o stampa un report. acViewNormal invia alla stampante predefinita.',
      'Opens or prints a report. acViewNormal sends it to the default printer.',
      'Abre o imprime un informe. acViewNormal lo envía a la impresora predeterminada.',
      'Ouvre ou imprime un état. acViewNormal l’envoie à l’imprimante par défaut.',
      'Öffnet oder druckt einen Bericht. acViewNormal sendet ihn an den Standarddrucker.',
    ),
    params: [
      { name: 'ReportName', description: L('Nome del report.', 'Report name.', 'Nombre del informe.', 'Nom de l’état.', 'Berichtsname.') },
      { name: 'View', optional: true, description: L('acViewNormal (stampa), acViewPreview, acViewDesign, acViewLayout.', 'acViewNormal (print), acViewPreview, acViewDesign, acViewLayout.', 'acViewNormal (impresión), acViewPreview, acViewDesign, acViewLayout.', 'acViewNormal (impression), acViewPreview, acViewDesign, acViewLayout.', 'acViewNormal (Druck), acViewPreview, acViewDesign, acViewLayout.') },
      { name: 'WhereCondition', optional: true, description: L('Filtro sui record del report.', 'Filter on the report records.', 'Filtro de los registros del informe.', 'Filtre sur les enregistrements de l’état.', 'Filter der Berichtsdatensätze.') },
    ],
    examples: [
      {
        title: L('Stampare gli ordini di un cliente', 'Print one customer’s orders', 'Imprimir los pedidos de un cliente', 'Imprimer les commandes d’un client', 'Bestellungen eines Kunden drucken'),
        code: `Sub StampaOrdiniCliente()
    DoCmd.OpenReport "rptOrdini", acViewNormal, , "IDCliente=12"
End Sub`,
        result: L(
          'rptOrdini viene inviato alla stampante con i soli ordini del cliente 12. Non resta aperto a video.',
          'rptOrdini is sent to the printer with only customer 12’s orders. It does not stay open on screen.',
          'rptOrdini se envía a la impresora solo con los pedidos del cliente 12. No permanece abierto en pantalla.',
          'rptOrdini est envoyé à l’imprimante avec uniquement les commandes du client 12. Il ne reste pas ouvert à l’écran.',
          'rptOrdini geht an den Drucker, nur mit Bestellungen von Kunde 12. Er bleibt nicht am Bildschirm geöffnet.',
        ),
      },
    ],
    notes: L(
      'Per l’anteprima usa acViewPreview (vedi il comando dedicato nei report).',
      'For preview use acViewPreview (see the dedicated reports command).',
      'Para la vista previa usa acViewPreview (véase el comando dedicado en informes).',
      'Pour l’aperçu utilisez acViewPreview (voir la commande dédiée dans les états).',
      'Für die Vorschau acViewPreview verwenden (siehe den eigenen Befehl unter Berichte).',
    ),
    related: ['ac-open-report-preview', 'ac-reports-bang', 'ac-output-to', 'ac-docmd-open-form'],
  },
  {
    id: 'ac-docmd-open-query',
    name: 'DoCmd.OpenQuery',
    syntax: 'DoCmd.OpenQuery QueryName, [View], [DataMode]',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'docmd',
    description: L(
      'Apre una query in foglio dati, struttura o anteprima. Le action query vengono eseguite.',
      'Opens a query in datasheet, design or preview. Action queries are executed.',
      'Abre una consulta en hoja de datos, diseño o vista previa. Las de acción se ejecutan.',
      'Ouvre une requête en feuille, création ou aperçu. Les requêtes Action s’exécutent.',
      'Öffnet eine Abfrage in Datenblatt, Entwurf oder Vorschau. Aktionsabfragen werden ausgeführt.',
    ),
    params: [
      { name: 'QueryName', description: L('Nome della query salvata.', 'Saved query name.', 'Nombre de la consulta guardada.', 'Nom de la requête enregistrée.', 'Name der gespeicherten Abfrage.') },
      { name: 'View', optional: true, description: L('acViewNormal, acViewDesign, acViewPreview.', 'acViewNormal, acViewDesign, acViewPreview.', 'acViewNormal, acViewDesign, acViewPreview.', 'acViewNormal, acViewDesign, acViewPreview.', 'acViewNormal, acViewDesign, acViewPreview.') },
      { name: 'DataMode', optional: true, description: L('acEdit, acAdd, acReadOnly.', 'acEdit, acAdd, acReadOnly.', 'acEdit, acAdd, acReadOnly.', 'acEdit, acAdd, acReadOnly.', 'acEdit, acAdd, acReadOnly.') },
    ],
    examples: [
      {
        title: L('Aprire qryClientiMilano in sola lettura', 'Open qryClientiMilano read-only', 'Abrir qryClientiMilano de solo lectura', 'Ouvrir qryClientiMilano en lecture seule', 'qryClientiMilano schreibgeschützt öffnen'),
        code: `Sub ApriQueryMilano()
    DoCmd.OpenQuery "qryClientiMilano", acViewNormal, acReadOnly
End Sub`,
        result: L(
          'Si apre il foglio dati di qryClientiMilano (es. 17 righe, Citta=Milano). Le celle non sono modificabili.',
          'The qryClientiMilano datasheet opens (e.g. 17 rows, Citta=Milano). Cells are not editable.',
          'Se abre la hoja de qryClientiMilano (p. ej. 17 filas, Citta=Milano). Las celdas no se pueden editar.',
          'La feuille qryClientiMilano s’ouvre (p. ex. 17 lignes, Citta=Milano). Les cellules ne sont pas modifiables.',
          'Das Datenblatt qryClientiMilano öffnet sich (z. B. 17 Zeilen, Citta=Milano). Zellen sind nicht editierbar.',
        ),
      },
    ],
    notes: L(
      'Una query di comando (UPDATE/DELETE) con OpenQuery viene eseguita e chiede conferma se SetWarnings è True.',
      'An action query (UPDATE/DELETE) opened with OpenQuery runs and prompts if SetWarnings is True.',
      'Una consulta de acción (UPDATE/DELETE) con OpenQuery se ejecuta y pregunta si SetWarnings es True.',
      'Une requête Action (UPDATE/DELETE) avec OpenQuery s’exécute et demande confirmation si SetWarnings est True.',
      'Eine Aktionsabfrage (UPDATE/DELETE) mit OpenQuery wird ausgeführt und fragt nach, wenn SetWarnings True ist.',
    ),
    related: ['ac-docmd-open-table', 'ac-docmd-run-sql', 'ac-dao-querydef-tabledef', 'ac-docmd-set-warnings'],
  },
  {
    id: 'ac-docmd-open-table',
    name: 'DoCmd.OpenTable',
    syntax: 'DoCmd.OpenTable TableName, [View], [DataMode]',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'docmd',
    description: L(
      'Apre una tabella in foglio dati, struttura o anteprima di stampa.',
      'Opens a table in datasheet, design or print preview.',
      'Abre una tabla en hoja de datos, diseño o vista previa de impresión.',
      'Ouvre une table en feuille de données, création ou aperçu avant impression.',
      'Öffnet eine Tabelle in Datenblatt, Entwurf oder Druckvorschau.',
    ),
    params: [
      { name: 'TableName', description: L('Nome della tabella.', 'Table name.', 'Nombre de la tabla.', 'Nom de la table.', 'Tabellenname.') },
      { name: 'View', optional: true, description: L('acViewNormal, acViewDesign, acViewPreview.', 'acViewNormal, acViewDesign, acViewPreview.', 'acViewNormal, acViewDesign, acViewPreview.', 'acViewNormal, acViewDesign, acViewPreview.', 'acViewNormal, acViewDesign, acViewPreview.') },
      { name: 'DataMode', optional: true, description: L('acEdit, acAdd, acReadOnly.', 'acEdit, acAdd, acReadOnly.', 'acEdit, acAdd, acReadOnly.', 'acEdit, acAdd, acReadOnly.', 'acEdit, acAdd, acReadOnly.') },
    ],
    examples: [
      {
        title: L('Aprire Prodotti in modifica', 'Open Prodotti for editing', 'Abrir Prodotti para editar', 'Ouvrir Prodotti en modification', 'Prodotti zum Bearbeiten öffnen'),
        code: `Sub ApriProdotti()
    DoCmd.OpenTable "Prodotti", acViewNormal, acEdit
End Sub`,
        result: L(
          'Foglio dati Prodotti (es. 54 articoli). Il record selector è sul primo; le celle sono modificabili.',
          'Prodotti datasheet (e.g. 54 items). Record selector is on the first row; cells are editable.',
          'Hoja Prodotti (p. ej. 54 artículos). El selector está en la primera fila; las celdas se pueden editar.',
          'Feuille Prodotti (p. ex. 54 articles). Le sélecteur est sur la première ligne ; les cellules sont modifiables.',
          'Datenblatt Prodotti (z. B. 54 Artikel). Der Datensatzmarkierer steht auf der ersten Zeile; Zellen sind editierbar.',
        ),
      },
    ],
    related: ['ac-docmd-open-query', 'ac-docmd-select-object', 'ac-dao-open-recordset'],
  },
  {
    id: 'ac-docmd-run-sql',
    name: 'DoCmd.RunSQL',
    syntax: 'DoCmd.RunSQL SQLStatement, [UseTransaction]',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'docmd',
    description: L(
      'Esegue un’istruzione SQL di comando (INSERT, UPDATE, DELETE, SELECT INTO).',
      'Runs an action SQL statement (INSERT, UPDATE, DELETE, SELECT INTO).',
      'Ejecuta una instrucción SQL de acción (INSERT, UPDATE, DELETE, SELECT INTO).',
      'Exécute une instruction SQL Action (INSERT, UPDATE, DELETE, SELECT INTO).',
      'Führt eine SQL-Aktionsanweisung aus (INSERT, UPDATE, DELETE, SELECT INTO).',
    ),
    params: [
      { name: 'SQLStatement', description: L('SQL di comando, non una SELECT di sola lettura.', 'Action SQL, not a read-only SELECT.', 'SQL de acción, no un SELECT de solo lectura.', 'SQL Action, pas un SELECT en lecture seule.', 'Aktions-SQL, kein schreibgeschütztes SELECT.') },
      { name: 'UseTransaction', optional: true, description: L('True (default) usa una transazione.', 'True (default) uses a transaction.', 'True (predeterminado) usa una transacción.', 'True (défaut) utilise une transaction.', 'True (Standard) verwendet eine Transaktion.') },
    ],
    examples: [
      {
        title: L('Aumentare i prezzi del 5%', 'Raise prices by 5%', 'Subir los precios un 5%', 'Augmenter les prix de 5 %', 'Preise um 5 % erhöhen'),
        code: `Sub AumentaPrezzi()
    DoCmd.SetWarnings False
    DoCmd.RunSQL "UPDATE Prodotti SET PrezzoUnitario = PrezzoUnitario * 1.05 WHERE Categoria='Bevande'"
    DoCmd.SetWarnings True
    Debug.Print DCount("*", "Prodotti", "Categoria='Bevande'") & " righe aggiornate"
End Sub`,
        result: L(
          'I 9 prodotti Bevande hanno il prezzo * 1,05 (es. 2,00 → 2,10). Nessuna finestra di conferma. Immediata: 9 righe aggiornate.',
          'The 9 Bevande products now have price * 1.05 (e.g. 2.00 → 2.10). No confirm dialog. Immediate: 9 righe aggiornate.',
          'Los 9 productos Bevande quedan con precio * 1,05 (p. ej. 2,00 → 2,10). Sin confirmación. Inmediato: 9 righe aggiornate.',
          'Les 9 produits Bevande ont le prix * 1,05 (p. ex. 2,00 → 2,10). Pas de confirmation. Exécution : 9 righe aggiornate.',
          'Die 9 Bevande-Produkte haben Preis * 1,05 (z. B. 2,00 → 2,10). Kein Bestätigungsdialog. Direktbereich: 9 righe aggiornate.',
        ),
      },
    ],
    notes: L(
      'CurrentDb.Execute sql, dbFailOnError è preferibile: restituisce RecordsAffected e solleva errore senza SetWarnings.',
      'CurrentDb.Execute sql, dbFailOnError is preferable: it returns RecordsAffected and raises errors without SetWarnings.',
      'CurrentDb.Execute sql, dbFailOnError es preferible: devuelve RecordsAffected y lanza error sin SetWarnings.',
      'CurrentDb.Execute sql, dbFailOnError est préférable : il renvoie RecordsAffected et lève une erreur sans SetWarnings.',
      'CurrentDb.Execute sql, dbFailOnError ist vorzuziehen: liefert RecordsAffected und löst Fehler ohne SetWarnings aus.',
    ),
    related: ['ac-docmd-set-warnings', 'ac-ado-execute', 'ac-dcount', 'ac-dao-querydef-tabledef'],
  },
  {
    id: 'ac-docmd-goto-record',
    name: 'DoCmd.GoToRecord',
    syntax: 'DoCmd.GoToRecord [ObjectType], [ObjectName], [Record], [Offset]',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'docmd',
    description: L(
      'Sposta il record corrente di una maschera o di un foglio dati.',
      'Moves the current record of a form or datasheet.',
      'Mueve el registro actual de un formulario o hoja de datos.',
      'Déplace l’enregistrement courant d’un formulaire ou d’une feuille.',
      'Bewegt den aktuellen Datensatz eines Formulars oder Datenblatts.',
    ),
    params: [
      { name: 'ObjectType', optional: true, description: L('acDataForm, acDataTable, acDataQuery, acActiveDataObject.', 'acDataForm, acDataTable, acDataQuery, acActiveDataObject.', 'acDataForm, acDataTable, acDataQuery, acActiveDataObject.', 'acDataForm, acDataTable, acDataQuery, acActiveDataObject.', 'acDataForm, acDataTable, acDataQuery, acActiveDataObject.') },
      { name: 'Record', optional: true, description: L('acFirst, acLast, acNext, acPrevious, acGoTo, acNewRec.', 'acFirst, acLast, acNext, acPrevious, acGoTo, acNewRec.', 'acFirst, acLast, acNext, acPrevious, acGoTo, acNewRec.', 'acFirst, acLast, acNext, acPrevious, acGoTo, acNewRec.', 'acFirst, acLast, acNext, acPrevious, acGoTo, acNewRec.') },
      { name: 'Offset', optional: true, description: L('Numero record se Record = acGoTo.', 'Record number if Record = acGoTo.', 'Número de registro si Record = acGoTo.', 'Numéro d’enregistrement si Record = acGoTo.', 'Datensatznummer wenn Record = acGoTo.') },
    ],
    examples: [
      {
        title: L('Andare all’ultimo e al nuovo record', 'Go to last and new record', 'Ir al último y al nuevo registro', 'Aller au dernier et au nouvel enregistrement', 'Zum letzten und zum neuen Datensatz'),
        code: `Sub NavigaClienti()
    DoCmd.OpenForm "frmClienti"
    DoCmd.GoToRecord acDataForm, "frmClienti", acLast
    Debug.Print Forms!frmClienti!Codice
    DoCmd.GoToRecord acDataForm, "frmClienti", acNewRec
    Debug.Print Forms!frmClienti.NewRecord
End Sub`,
        result: L(
          'Maschera sull’ultimo cliente (es. Codice=128). Poi record nuovo: NewRecord = True, campi vuoti.',
          'Form on the last customer (e.g. Codice=128). Then new record: NewRecord = True, empty fields.',
          'Formulario en el último cliente (p. ej. Codice=128). Luego registro nuevo: NewRecord = True, campos vacíos.',
          'Formulaire sur le dernier client (p. ex. Codice=128). Puis nouvel enregistrement : NewRecord = True, champs vides.',
          'Formular auf dem letzten Kunden (z. B. Codice=128). Dann neuer Datensatz: NewRecord = True, leere Felder.',
        ),
      },
    ],
    related: ['ac-docmd-find-record', 'ac-docmd-apply-filter', 'ac-forms-bang', 'ac-dao-move-ends'],
  },
  {
    id: 'ac-docmd-find-record',
    name: 'DoCmd.FindRecord',
    syntax: 'DoCmd.FindRecord FindWhat, [Match], [MatchCase], [Search], [SearchAsFormatted], [OnlyCurrentField], [FindFirst]',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'docmd',
    description: L(
      'Cerca un valore nel recordset della maschera o del foglio dati attivo.',
      'Searches for a value in the active form or datasheet recordset.',
      'Busca un valor en el recordset del formulario u hoja activos.',
      'Recherche une valeur dans le recordset du formulaire ou de la feuille actifs.',
      'Sucht einen Wert im Recordset des aktiven Formulars oder Datenblatts.',
    ),
    params: [
      { name: 'FindWhat', description: L('Valore da trovare.', 'Value to find.', 'Valor a buscar.', 'Valeur à trouver.', 'Zu suchender Wert.') },
      { name: 'Match', optional: true, description: L('acAnywhere, acEntire, acStart.', 'acAnywhere, acEntire, acStart.', 'acAnywhere, acEntire, acStart.', 'acAnywhere, acEntire, acStart.', 'acAnywhere, acEntire, acStart.') },
      { name: 'OnlyCurrentField', optional: true, description: L('acCurrent, acAll.', 'acCurrent, acAll.', 'acCurrent, acAll.', 'acCurrent, acAll.', 'acCurrent, acAll.') },
      { name: 'FindFirst', optional: true, description: L('True riparte dall’inizio.', 'True restarts from the beginning.', 'True reinicia desde el principio.', 'True repart du début.', 'True startet von vorn.') },
    ],
    examples: [
      {
        title: L('Trovare Bianchi nel campo RagioneSociale', 'Find Bianchi in RagioneSociale', 'Buscar Bianchi en RagioneSociale', 'Trouver Bianchi dans RagioneSociale', 'Bianchi in RagioneSociale finden'),
        code: `Sub TrovaBianchi()
    DoCmd.OpenForm "frmClienti"
    DoCmd.GoToControl "RagioneSociale"
    DoCmd.FindRecord "Bianchi", acStart, False, acDown, False, acCurrent, True
    Debug.Print Forms!frmClienti!Codice & " " & Forms!frmClienti!RagioneSociale
End Sub`,
        result: L(
          'frmClienti si posiziona su Codice=18, RagioneSociale=Bianchi Srl.',
          'frmClienti moves to Codice=18, RagioneSociale=Bianchi Srl.',
          'frmClienti se sitúa en Codice=18, RagioneSociale=Bianchi Srl.',
          'frmClienti se place sur Codice=18, RagioneSociale=Bianchi Srl.',
          'frmClienti steht auf Codice=18, RagioneSociale=Bianchi Srl.',
        ),
      },
    ],
    related: ['ac-docmd-goto-record', 'ac-dao-find', 'ac-docmd-apply-filter'],
  },
  {
    id: 'ac-docmd-apply-filter',
    name: 'DoCmd.ApplyFilter',
    syntax: 'DoCmd.ApplyFilter [FilterName], [WhereCondition], [ControlName]',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'docmd',
    description: L(
      'Applica un filtro o una query alla maschera o al report attivo.',
      'Applies a filter or query to the active form or report.',
      'Aplica un filtro o una consulta al formulario o informe activo.',
      'Applique un filtre ou une requête au formulaire ou à l’état actif.',
      'Wendet einen Filter oder eine Abfrage auf das aktive Formular oder den Bericht an.',
    ),
    params: [
      { name: 'FilterName', optional: true, description: L('Nome di una query o filtro salvato.', 'Name of a saved query or filter.', 'Nombre de una consulta o filtro guardado.', 'Nom d’une requête ou d’un filtre enregistré.', 'Name einer gespeicherten Abfrage oder eines Filters.') },
      { name: 'WhereCondition', optional: true, description: L('Criterio SQL senza WHERE.', 'SQL criterion without WHERE.', 'Criterio SQL sin WHERE.', 'Critère SQL sans WHERE.', 'SQL-Kriterium ohne WHERE.') },
    ],
    examples: [
      {
        title: L('Filtrare i clienti di Torino', 'Filter customers in Turin', 'Filtrar los clientes de Turín', 'Filtrer les clients de Turin', 'Kunden in Turin filtern'),
        code: `Sub FiltraTorino()
    DoCmd.OpenForm "frmClienti"
    DoCmd.ApplyFilter , "Citta='Torino'"
    Debug.Print Forms!frmClienti.Filter
    Debug.Print Forms!frmClienti.FilterOn
End Sub`,
        result: L(
          'frmClienti mostra solo Torino (es. 11 record). Filter = Citta=\'Torino\', FilterOn = True.',
          'frmClienti shows only Turin (e.g. 11 records). Filter = Citta=\'Torino\', FilterOn = True.',
          'frmClienti muestra solo Turín (p. ej. 11 registros). Filter = Citta=\'Torino\', FilterOn = True.',
          'frmClienti n’affiche que Turin (p. ex. 11 enregistrements). Filter = Citta=\'Torino\', FilterOn = True.',
          'frmClienti zeigt nur Turin (z. B. 11 Datensätze). Filter = Citta=\'Torino\', FilterOn = True.',
        ),
      },
    ],
    related: ['ac-docmd-show-all-records', 'ac-docmd-open-form', 'ac-dcount'],
  },
  {
    id: 'ac-docmd-set-warnings',
    name: 'DoCmd.SetWarnings',
    syntax: 'DoCmd.SetWarnings WarningsOn',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'docmd',
    description: L(
      'Attiva o disattiva i messaggi di conferma delle query di comando.',
      'Turns action-query confirmation messages on or off.',
      'Activa o desactiva los mensajes de confirmación de las consultas de acción.',
      'Active ou désactive les messages de confirmation des requêtes Action.',
      'Schaltet Bestätigungsmeldungen von Aktionsabfragen ein oder aus.',
    ),
    params: [
      { name: 'WarningsOn', description: L('True = avvisi visibili; False = silenziosi.', 'True = warnings visible; False = silent.', 'True = avisos visibles; False = silenciosos.', 'True = avertissements visibles ; False = silencieux.', 'True = Warnungen sichtbar; False = still.') },
    ],
    examples: [
      {
        title: L('DELETE senza finestra di conferma', 'DELETE with no confirm dialog', 'DELETE sin cuadro de confirmación', 'DELETE sans boîte de confirmation', 'DELETE ohne Bestätigungsdialog'),
        code: `Sub PulisciBozze()
    On Error GoTo Fine
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE FROM Ordini WHERE Stato='Bozza'"
Fine:
    DoCmd.SetWarnings True
    Debug.Print "avvisi ripristinati"
End Sub`,
        result: L(
          'Le 3 righe Bozza vengono eliminate senza popup. In Immediata: avvisi ripristinati. SetWarnings True è nel Fine: anche se c’è errore.',
          'The 3 Bozza rows are deleted with no popup. Immediate: avvisi ripristinati. SetWarnings True runs in Fine: even after an error.',
          'Las 3 filas Bozza se eliminan sin ventana. Inmediato: avvisi ripristinati. SetWarnings True corre en Fine: aunque haya error.',
          'Les 3 lignes Bozza sont supprimées sans popup. Exécution : avvisi ripristinati. SetWarnings True s’exécute dans Fine: même après erreur.',
          'Die 3 Bozza-Zeilen werden ohne Popup gelöscht. Direktbereich: avvisi ripristinati. SetWarnings True läuft in Fine: auch nach einem Fehler.',
        ),
      },
    ],
    notes: L(
      'Metti sempre SetWarnings True in un gestore di errore. Un False dimenticato nasconde gli avvisi per tutta la sessione.',
      'Always set SetWarnings True in an error handler. A leftover False hides warnings for the whole session.',
      'Pon siempre SetWarnings True en el gestor de errores. Un False olvidado oculta avisos toda la sesión.',
      'Remettez toujours SetWarnings True dans le gestionnaire d’erreur. Un False oublié masque les avis pour toute la session.',
      'SetWarnings True immer im Fehlerhandler zurücksetzen. Ein vergessenes False unterdrückt Warnungen die ganze Sitzung.',
    ),
    related: ['ac-docmd-run-sql', 'ac-docmd-ui-feedback', 'ac-ado-execute'],
  },
  {
    id: 'ac-docmd-ui-feedback',
    name: 'DoCmd.Hourglass / Echo / Beep',
    syntax: 'DoCmd.Hourglass HourglassOn | DoCmd.Echo EchoOn, [StatusBarText] | DoCmd.Beep',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'docmd',
    description: L(
      'Feedback UI: cursore clessidra, ridisegno schermo e beep di sistema.',
      'UI feedback: hourglass cursor, screen painting and system beep.',
      'Comentarios de IU: cursor de reloj de arena, repintado y beep del sistema.',
      'Retour UI : curseur sablier, dessin de l’écran et bip système.',
      'UI-Rückmeldung: Sanduhrzeiger, Bildschirmaktualisierung und Systemton.',
    ),
    params: [
      { name: 'HourglassOn', description: L('True = clessidra; False = puntatore normale.', 'True = hourglass; False = normal pointer.', 'True = reloj de arena; False = puntero normal.', 'True = sablier ; False = pointeur normal.', 'True = Sanduhr; False = normaler Zeiger.') },
      { name: 'EchoOn', description: L('False sospende l’aggiornamento della UI.', 'False suspends UI painting.', 'False suspende el repintado de la IU.', 'False suspend le dessin de l’UI.', 'False hält die UI-Aktualisierung an.') },
      { name: 'StatusBarText', optional: true, description: L('Testo in barra di stato se Echo è False.', 'Status-bar text when Echo is False.', 'Texto de la barra de estado si Echo es False.', 'Texte de la barre d’état si Echo est False.', 'Statusleistentext wenn Echo False ist.') },
    ],
    examples: [
      {
        title: L('Clessidra ed Echo durante un ciclo', 'Hourglass and Echo during a loop', 'Reloj de arena y Echo durante un bucle', 'Sablier et Echo pendant une boucle', 'Sanduhr und Echo während einer Schleife'),
        code: `Sub RicalcolaSconti()
    On Error GoTo Fine
    DoCmd.Hourglass True
    DoCmd.Echo False, "Ricalcolo sconti..."
    CurrentDb.Execute "UPDATE Clienti SET Sconto = 5 WHERE Citta='Milano'", dbFailOnError
Fine:
    DoCmd.Echo True
    DoCmd.Hourglass False
    DoCmd.Beep
    Debug.Print "fine"
End Sub`,
        result: L(
          'Durante l’UPDATE il mouse è clessidra e la UI non lampeggia. A fine lavoro: beep e Immediata = fine. 22 clienti Milano hanno Sconto=5.',
          'During the UPDATE the mouse is an hourglass and the UI does not flicker. When done: beep and Immediate = fine. 22 Milan customers have Sconto=5.',
          'Durante el UPDATE el ratón es reloj de arena y la IU no parpadea. Al terminar: beep e Inmediato = fine. 22 clientes de Milán tienen Sconto=5.',
          'Pendant l’UPDATE la souris est un sablier et l’UI ne clignote pas. À la fin : bip et Exécution = fine. 22 clients de Milan ont Sconto=5.',
          'Während des UPDATE ist die Maus eine Sanduhr und die UI flackert nicht. Danach: Signalton und Direktbereich = fine. 22 Mailänder Kunden haben Sconto=5.',
        ),
      },
    ],
    notes: L(
      'Echo False va sempre ripristinato: altrimenti Access sembra bloccato. Beep usa il suono predefinito di Windows.',
      'Always restore Echo True or Access looks frozen. Beep uses the default Windows sound.',
      'Restaura siempre Echo True o Access parece bloqueado. Beep usa el sonido predeterminado de Windows.',
      'Restaurez toujours Echo True sinon Access paraît figé. Beep utilise le son Windows par défaut.',
      'Echo True immer zurücksetzen, sonst wirkt Access eingefroren. Beep verwendet den Windows-Standardton.',
    ),
    related: ['ac-docmd-set-warnings', 'ac-docmd-run-sql', 'ac-docmd-quit'],
  },
  {
    id: 'ac-docmd-quit',
    name: 'DoCmd.Quit',
    syntax: 'DoCmd.Quit [Options]',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'docmd',
    description: L(
      'Chiude Microsoft Access e il database corrente.',
      'Closes Microsoft Access and the current database.',
      'Cierra Microsoft Access y la base de datos actual.',
      'Ferme Microsoft Access et la base de données courante.',
      'Beendet Microsoft Access und die aktuelle Datenbank.',
    ),
    params: [
      { name: 'Options', optional: true, description: L('acQuitSaveAll, acQuitPrompt, acQuitSaveNone.', 'acQuitSaveAll, acQuitPrompt, acQuitSaveNone.', 'acQuitSaveAll, acQuitPrompt, acQuitSaveNone.', 'acQuitSaveAll, acQuitPrompt, acQuitSaveNone.', 'acQuitSaveAll, acQuitPrompt, acQuitSaveNone.') },
    ],
    examples: [
      {
        title: L('Uscire salvando gli oggetti modificati', 'Exit saving changed objects', 'Salir guardando los objetos cambiados', 'Quitter en enregistrant les objets modifiés', 'Beenden und geänderte Objekte speichern'),
        code: `Sub EsciAccess()
    DoCmd.Quit acQuitSaveAll
End Sub`,
        result: L(
          'Access si chiude. Maschere/report/macro sporchi vengono salvati; i record Dirty delle maschere seguono le regole di salvataggio form.',
          'Access closes. Dirty forms/reports/macros are saved; Dirty form records follow the form’s save rules.',
          'Access se cierra. Los formularios/informes/macros sucios se guardan; los registros Dirty siguen las reglas del formulario.',
          'Access se ferme. Les formulaires/états/macros sales sont enregistrés ; les enregistrements Dirty suivent les règles du formulaire.',
          'Access wird beendet. Geänderte Formulare/Berichte/Makros werden gespeichert; Dirty-Datensätze folgen den Formularregeln.',
        ),
      },
    ],
    related: ['ac-docmd-close', 'ac-me-undo-dirty'],
  },
  {
    id: 'ac-docmd-window',
    name: 'DoCmd.Maximize / Minimize / Restore',
    syntax: 'DoCmd.Maximize | DoCmd.Minimize | DoCmd.Restore',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'docmd',
    description: L(
      'Massimizza, riduce a icona o ripristina la finestra dell’oggetto attivo.',
      'Maximizes, minimizes or restores the active object window.',
      'Maximiza, minimiza o restaura la ventana del objeto activo.',
      'Agrandit, réduit ou restaure la fenêtre de l’objet actif.',
      'Maximiert, minimiert oder stellt das Fenster des aktiven Objekts wieder her.',
    ),
    examples: [
      {
        title: L('Massimizzare poi ripristinare frmOrdini', 'Maximize then restore frmOrdini', 'Maximizar y luego restaurar frmOrdini', 'Agrandir puis restaurer frmOrdini', 'frmOrdini maximieren und wiederherstellen'),
        code: `Sub FinestraOrdini()
    DoCmd.OpenForm "frmOrdini"
    DoCmd.Maximize
    Debug.Print "massimizzata"
    DoCmd.Restore
    Debug.Print "ripristinata"
End Sub`,
        result: L(
          'frmOrdini occupa tutto lo spazio MDI, poi torna alle dimensioni precedenti. Immediata: massimizzata / ripristinata. Minimize la manderebbe sulla barra.',
          'frmOrdini fills the MDI workspace, then returns to its previous size. Immediate: massimizzata / ripristinata. Minimize would send it to the bar.',
          'frmOrdini llena el espacio MDI y luego vuelve a su tamaño anterior. Inmediato: massimizzata / ripristinata. Minimize la enviaría a la barra.',
          'frmOrdini remplit l’espace MDI puis reprend sa taille précédente. Exécution : massimizzata / ripristinata. Minimize l’enverrait dans la barre.',
          'frmOrdini füllt den MDI-Bereich und kehrt dann zur vorherigen Größe zurück. Direktbereich: massimizzata / ripristinata. Minimize legt es in die Leiste.',
        ),
      },
    ],
    notes: L(
      'In interfaccia a schede (Tabbed Documents) Maximize/Restore hanno poco effetto visivo; servono soprattutto in Overlapping Windows.',
      'In a tabbed document interface Maximize/Restore have little visual effect; they matter mainly with Overlapping Windows.',
      'En interfaz por pestañas Maximize/Restore tienen poco efecto; importan sobre todo con Overlapping Windows.',
      'En interface à onglets Maximize/Restore ont peu d’effet ; ils comptent surtout avec Overlapping Windows.',
      'Bei Registerkarten haben Maximize/Restore kaum sichtbaren Effekt; sie zählen vor allem bei überlappenden Fenstern.',
    ),
    related: ['ac-docmd-open-form', 'ac-docmd-select-object'],
  },
  {
    id: 'ac-docmd-select-object',
    name: 'DoCmd.SelectObject',
    syntax: 'DoCmd.SelectObject ObjectType, [ObjectName], [InDatabaseWindow]',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'docmd',
    description: L(
      'Seleziona un oggetto nel riquadro di spostamento o gli dà lo stato attivo se è aperto.',
      'Selects an object in the navigation pane or focuses it if it is open.',
      'Selecciona un objeto en el panel de navegación o le da el foco si está abierto.',
      'Sélectionne un objet dans le volet de navigation ou lui donne le focus s’il est ouvert.',
      'Wählt ein Objekt im Navigationsbereich oder gibt ihm den Fokus, wenn es geöffnet ist.',
    ),
    params: [
      { name: 'ObjectType', description: L('acTable, acQuery, acForm, acReport, acMacro, acModule.', 'acTable, acQuery, acForm, acReport, acMacro, acModule.', 'acTable, acQuery, acForm, acReport, acMacro, acModule.', 'acTable, acQuery, acForm, acReport, acMacro, acModule.', 'acTable, acQuery, acForm, acReport, acMacro, acModule.') },
      { name: 'ObjectName', optional: true, description: L('Nome oggetto.', 'Object name.', 'Nombre del objeto.', 'Nom de l’objet.', 'Objektname.') },
      { name: 'InDatabaseWindow', optional: true, description: L('True = evidenzialo nel riquadro anche se chiuso.', 'True = highlight it in the pane even if closed.', 'True = resáltalo en el panel aunque esté cerrado.', 'True = le surligner dans le volet même s’il est fermé.', 'True = im Bereich markieren, auch wenn geschlossen.') },
    ],
    examples: [
      {
        title: L('Evidenziare la tabella Clienti', 'Highlight the Clienti table', 'Resaltar la tabla Clienti', 'Surligner la table Clienti', 'Tabelle Clienti markieren'),
        code: `Sub SelezionaClienti()
    DoCmd.SelectObject acTable, "Clienti", True
End Sub`,
        result: L(
          'Nel riquadro di spostamento Clienti è selezionata (sfondo evidenziato), anche se la tabella non è aperta.',
          'In the navigation pane Clienti is selected (highlighted), even if the table is not open.',
          'En el panel de navegación Clienti queda seleccionada (resaltada), aunque la tabla no esté abierta.',
          'Dans le volet de navigation Clienti est sélectionnée (surlignée), même si la table n’est pas ouverte.',
          'Im Navigationsbereich ist Clienti ausgewählt (hervorgehoben), auch wenn die Tabelle nicht geöffnet ist.',
        ),
      },
    ],
    related: ['ac-docmd-open-table', 'ac-docmd-open-module', 'ac-docmd-window'],
  },
  {
    id: 'ac-docmd-run-macro',
    name: 'DoCmd.RunMacro',
    syntax: 'DoCmd.RunMacro MacroName, [RepeatCount], [RepeatExpression]',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'docmd',
    description: L(
      'Esegue una macro salvata (o una sottomacro con il punto).',
      'Runs a saved macro (or a submacro with a dot).',
      'Ejecuta una macro guardada (o una submacro con punto).',
      'Exécute une macro enregistrée (ou une sous-macro avec un point).',
      'Führt ein gespeichertes Makro aus (oder ein Submakro mit Punkt).',
    ),
    params: [
      { name: 'MacroName', description: L('Nome macro, es. mcrApriClienti o mcrMenu.Apri.', 'Macro name, e.g. mcrApriClienti or mcrMenu.Apri.', 'Nombre de macro, p. ej. mcrApriClienti o mcrMenu.Apri.', 'Nom de macro, p. ex. mcrApriClienti ou mcrMenu.Apri.', 'Makroname, z. B. mcrApriClienti oder mcrMenu.Apri.') },
      { name: 'RepeatCount', optional: true, description: L('Quante volte ripetere.', 'How many times to repeat.', 'Cuántas veces repetir.', 'Combien de fois répéter.', 'Wie oft wiederholen.') },
    ],
    examples: [
      {
        title: L('Lanciare la macro di apertura', 'Run the open macro', 'Ejecutar la macro de apertura', 'Lancer la macro d’ouverture', 'Öffnungs-Makro starten'),
        code: `Sub LanciaMacro()
    DoCmd.RunMacro "mcrApriClienti"
End Sub`,
        result: L(
          'La macro mcrApriClienti gira: se contiene OpenForm, si apre frmClienti come definito nella macro.',
          'Macro mcrApriClienti runs: if it contains OpenForm, frmClienti opens as the macro defines.',
          'La macro mcrApriClienti se ejecuta: si contiene OpenForm, se abre frmClienti como define la macro.',
          'La macro mcrApriClienti s’exécute : si elle contient OpenForm, frmClienti s’ouvre comme défini.',
          'Makro mcrApriClienti läuft: enthält es OpenForm, öffnet sich frmClienti wie im Makro definiert.',
        ),
      },
    ],
    related: ['ac-docmd-open-form', 'ac-docmd-open-module'],
  },
  {
    id: 'ac-docmd-open-module',
    name: 'DoCmd.OpenModule',
    syntax: 'DoCmd.OpenModule [ModuleName], [ProcedureName]',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'docmd',
    description: L(
      'Apre un modulo standard o di classe nell’editor VBA, opzionalmente su una procedura.',
      'Opens a standard or class module in the VBA editor, optionally on a procedure.',
      'Abre un módulo estándar o de clase en el editor VBA, opcionalmente en un procedimiento.',
      'Ouvre un module standard ou de classe dans l’éditeur VBA, éventuellement sur une procédure.',
      'Öffnet ein Standard- oder Klassenmodul im VBA-Editor, optional bei einer Prozedur.',
    ),
    params: [
      { name: 'ModuleName', optional: true, description: L('Nome modulo, es. modUtilita.', 'Module name, e.g. modUtilita.', 'Nombre del módulo, p. ej. modUtilita.', 'Nom du module, p. ex. modUtilita.', 'Modulname, z. B. modUtilita.') },
      { name: 'ProcedureName', optional: true, description: L('Sub o Function da mostrare.', 'Sub or Function to show.', 'Sub o Function a mostrar.', 'Sub ou Function à afficher.', 'Anzuzeigende Sub oder Function.') },
    ],
    examples: [
      {
        title: L('Aprire modUtilita sulla funzione Iva', 'Open modUtilita on function Iva', 'Abrir modUtilita en la función Iva', 'Ouvrir modUtilita sur la fonction Iva', 'modUtilita bei Funktion Iva öffnen'),
        code: `Sub ApriModuloIva()
    DoCmd.OpenModule "modUtilita", "Iva"
End Sub`,
        result: L(
          'Si apre l’editor VBA su modUtilita, cursore sulla Function Iva (o Sub Iva).',
          'The VBA editor opens on modUtilita, caret on Function Iva (or Sub Iva).',
          'Se abre el editor VBA en modUtilita, cursor en Function Iva (o Sub Iva).',
          'L’éditeur VBA s’ouvre sur modUtilita, curseur sur Function Iva (ou Sub Iva).',
          'Der VBA-Editor öffnet modUtilita, Cursor auf Function Iva (oder Sub Iva).',
        ),
      },
    ],
    related: ['ac-docmd-select-object', 'ac-docmd-run-macro'],
  },
  {
    id: 'ac-docmd-show-all-records',
    name: 'DoCmd.ShowAllRecords',
    syntax: 'DoCmd.ShowAllRecords',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'docmd',
    description: L(
      'Rimuove il filtro dalla maschera o dal foglio dati attivo.',
      'Removes the filter from the active form or datasheet.',
      'Quita el filtro del formulario u hoja de datos activo.',
      'Supprime le filtre du formulaire ou de la feuille actifs.',
      'Entfernt den Filter vom aktiven Formular oder Datenblatt.',
    ),
    examples: [
      {
        title: L('Togliere il filtro da frmClienti', 'Clear the filter on frmClienti', 'Quitar el filtro de frmClienti', 'Ôter le filtre de frmClienti', 'Filter von frmClienti entfernen'),
        code: `Sub MostraTuttiClienti()
    DoCmd.OpenForm "frmClienti"
    DoCmd.ApplyFilter , "Citta='Torino'"
    Debug.Print Forms!frmClienti.Recordset.RecordCount
    DoCmd.ShowAllRecords
    Debug.Print Forms!frmClienti.FilterOn
End Sub`,
        result: L(
          'Dopo il filtro: 11 record. Dopo ShowAllRecords: FilterOn = False, di nuovo tutti i 128 clienti.',
          'After the filter: 11 records. After ShowAllRecords: FilterOn = False, all 128 customers again.',
          'Tras el filtro: 11 registros. Tras ShowAllRecords: FilterOn = False, de nuevo los 128 clientes.',
          'Après le filtre : 11 enregistrements. Après ShowAllRecords : FilterOn = False, de nouveau les 128 clients.',
          'Nach dem Filter: 11 Datensätze. Nach ShowAllRecords: FilterOn = False, wieder alle 128 Kunden.',
        ),
      },
    ],
    related: ['ac-docmd-apply-filter', 'ac-docmd-requery', 'ac-me-requery'],
  },
  {
    id: 'ac-docmd-requery',
    name: 'DoCmd.Requery',
    syntax: 'DoCmd.Requery [ControlName]',
    scope: ['access'],
    category: 'ac-app',
    subcategory: 'docmd',
    description: L(
      'Riesegue l’origine dati della maschera attiva o di un suo controllo.',
      'Re-runs the data source of the active form or of one of its controls.',
      'Vuelve a ejecutar el origen del formulario activo o de uno de sus controles.',
      'Ré-exécute la source du formulaire actif ou de l’un de ses contrôles.',
      'Führt die Datenquelle des aktiven Formulars oder eines Steuerelements erneut aus.',
    ),
    params: [
      { name: 'ControlName', optional: true, description: L('Nome controllo (combo, listbox). Omesso: tutta la maschera.', 'Control name (combo, list box). Omitted: the whole form.', 'Nombre del control (combo, lista). Omitido: todo el formulario.', 'Nom du contrôle (liste). Omise : tout le formulaire.', 'Steuerelementname (Kombi, Liste). Weggelassen: das ganze Formular.') },
    ],
    examples: [
      {
        title: L('Aggiornare la maschera dopo un INSERT', 'Refresh the form after an INSERT', 'Actualizar el formulario tras un INSERT', 'Actualiser le formulaire après un INSERT', 'Formular nach INSERT aktualisieren'),
        code: `Sub DopoInserimento()
    CurrentDb.Execute "INSERT INTO Clienti (RagioneSociale, Citta) VALUES ('Neri Spa','Bologna')"
    DoCmd.OpenForm "frmClienti"
    DoCmd.Requery
    Debug.Print DCount("*", "Clienti")
End Sub`,
        result: L(
          'frmClienti ricarica i record: compare Neri Spa. Immediata: 129 (era 128).',
          'frmClienti reloads its records: Neri Spa appears. Immediate: 129 (was 128).',
          'frmClienti recarga los registros: aparece Neri Spa. Inmediato: 129 (era 128).',
          'frmClienti recharge les enregistrements : Neri Spa apparaît. Exécution : 129 (c’était 128).',
          'frmClienti lädt die Datensätze neu: Neri Spa erscheint. Direktbereich: 129 (vorher 128).',
        ),
      },
    ],
    related: ['ac-me-requery', 'ac-dao-requery', 'ac-docmd-show-all-records'],
  },
  {
    id: 'ac-dlookup',
    name: 'DLookup',
    syntax: 'DLookup(Expr, Domain, [Criteria])',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'domain',
    description: L(
      'Restituisce il valore di un campo dalla prima riga che soddisfa i criteri.',
      'Returns a field value from the first row that matches the criteria.',
      'Devuelve el valor de un campo de la primera fila que cumple los criterios.',
      'Renvoie la valeur d’un champ de la première ligne qui satisfait les critères.',
      'Gibt den Feldwert der ersten Zeile zurück, die die Kriterien erfüllt.',
    ),
    params: [
      { name: 'Expr', description: L('Campo o espressione, es. "RagioneSociale" o "Prezzo*Quantita".', 'Field or expression, e.g. "RagioneSociale" or "Prezzo*Quantita".', 'Campo o expresión, p. ej. "RagioneSociale" o "Prezzo*Quantita".', 'Champ ou expression, p. ex. "RagioneSociale" ou "Prezzo*Quantita".', 'Feld oder Ausdruck, z. B. "RagioneSociale" oder "Prezzo*Quantita".') },
      { name: 'Domain', description: L('Tabella o query.', 'Table or query.', 'Tabla o consulta.', 'Table ou requête.', 'Tabelle oder Abfrage.') },
      { name: 'Criteria', optional: true, description: L('WHERE senza la parola WHERE. Attenzione alle virgolette sulle stringhe.', 'WHERE without the word WHERE. Watch quotes around strings.', 'WHERE sin la palabra WHERE. Cuidado con las comillas.', 'WHERE sans le mot WHERE. Attention aux guillemets.', 'WHERE ohne das Wort WHERE. Anführungszeichen bei Text beachten.') },
    ],
    examples: [
      {
        title: L('Città del cliente 12', 'City of customer 12', 'Ciudad del cliente 12', 'Ville du client 12', 'Stadt von Kunde 12'),
        code: `Sub EsempioDLookup()
    Dim citta As Variant
    citta = DLookup("Citta", "Clienti", "Codice=12")
    Debug.Print Nz(citta, "(non trovato)")
    Debug.Print DLookup("RagioneSociale", "Clienti", "Email='info@bianchi.it'")
End Sub`,
        result: L(
          'Finestra Immediata: Milano  poi  Bianchi Srl',
          'Immediate Window: Milano  then  Bianchi Srl',
          'Ventana Inmediato: Milano  luego  Bianchi Srl',
          'Fenêtre Exécution : Milano  puis  Bianchi Srl',
          'Direktbereich: Milano  dann  Bianchi Srl',
        ),
      },
    ],
    notes: L(
      'Se nessuna riga corrisponde, DLookup restituisce Null: avvolgilo con Nz. Non è indicizzato come un Seek su chiave.',
      'If no row matches, DLookup returns Null: wrap it with Nz. It is not as fast as Seek on a key.',
      'Si ninguna fila coincide, DLookup devuelve Null: envuélvelo con Nz. No es tan rápido como Seek en una clave.',
      'Si aucune ligne ne correspond, DLookup renvoie Null : entourez-le de Nz. Ce n’est pas aussi rapide qu’un Seek sur clé.',
      'Ohne Treffer liefert DLookup Null: mit Nz umhüllen. Es ist nicht so schnell wie Seek auf einem Schlüssel.',
    ),
    related: ['ac-nz', 'ac-dcount', 'ac-dfirst', 'ac-dao-find', 'ac-dao-seek'],
  },
  {
    id: 'ac-dcount',
    name: 'DCount',
    syntax: 'DCount(Expr, Domain, [Criteria])',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'domain',
    description: L(
      'Conta le righe del dominio (tabella/query) che soddisfano i criteri.',
      'Counts the domain (table/query) rows that match the criteria.',
      'Cuenta las filas del dominio (tabla/consulta) que cumplen los criterios.',
      'Compte les lignes du domaine (table/requête) qui satisfont les critères.',
      'Zählt die Zeilen der Domäne (Tabelle/Abfrage), die die Kriterien erfüllen.',
    ),
    params: [
      { name: 'Expr', description: L('Di solito "*" oppure un campo da contare (ignora Null).', 'Usually "*" or a field to count (skips Null).', 'Suele ser "*" o un campo a contar (ignora Null).', 'En général "*" ou un champ à compter (ignore Null).', 'Meist "*" oder ein zu zählendes Feld (übergeht Null).') },
      { name: 'Domain', description: L('Tabella o query.', 'Table or query.', 'Tabla o consulta.', 'Table ou requête.', 'Tabelle oder Abfrage.') },
      { name: 'Criteria', optional: true, description: L('Filtro opzionale.', 'Optional filter.', 'Filtro opcional.', 'Filtre facultatif.', 'Optionaler Filter.') },
    ],
    examples: [
      {
        title: L('Clienti totali e di Milano', 'All customers and those in Milan', 'Clientes totales y de Milán', 'Clients totaux et ceux de Milan', 'Alle Kunden und die in Mailand'),
        code: `Sub EsempioDCount()
    Debug.Print DCount("*", "Clienti")
    Debug.Print DCount("*", "Clienti", "Citta='Milano'")
    Debug.Print DCount("Email", "Clienti")
End Sub`,
        result: L(
          'Finestra Immediata: 128  poi  22  poi  121  (7 clienti senza Email non sono contati)',
          'Immediate Window: 128  then  22  then  121  (7 customers without Email are not counted)',
          'Ventana Inmediato: 128  luego  22  luego  121  (7 clientes sin Email no se cuentan)',
          'Fenêtre Exécution : 128  puis  22  puis  121  (7 clients sans Email ne sont pas comptés)',
          'Direktbereich: 128  dann  22  dann  121  (7 Kunden ohne Email werden nicht gezählt)',
        ),
      },
    ],
    related: ['ac-dsum', 'ac-dlookup', 'ac-dao-recordcount'],
  },
  {
    id: 'ac-dsum',
    name: 'DSum',
    syntax: 'DSum(Expr, Domain, [Criteria])',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'domain',
    description: L(
      'Somma un’espressione numerica sul dominio filtrato.',
      'Sums a numeric expression over the filtered domain.',
      'Suma una expresión numérica sobre el dominio filtrado.',
      'Additionne une expression numérique sur le domaine filtré.',
      'Summiert einen numerischen Ausdruck über die gefilterte Domäne.',
    ),
    examples: [
      {
        title: L('Totale ordini 2026 del cliente 12', '2026 order total for customer 12', 'Total de pedidos 2026 del cliente 12', 'Total des commandes 2026 du client 12', 'Bestellsumme 2026 für Kunde 12'),
        code: `Sub EsempioDSum()
    Dim tot As Variant
    tot = DSum("Importo", "Ordini", "IDCliente=12 AND Year(DataOrdine)=2026")
    Debug.Print Nz(tot, 0)
    Debug.Print DSum("Quantita*PrezzoUnitario", "OrdiniDettaglio", "IDProdotto=7")
End Sub`,
        result: L(
          'Finestra Immediata: 18450  poi  960  (righe di OrdiniDettaglio del prodotto 7)',
          'Immediate Window: 18450  then  960  (OrdiniDettaglio rows for product 7)',
          'Ventana Inmediato: 18450  luego  960  (filas de OrdiniDettaglio del producto 7)',
          'Fenêtre Exécution : 18450  puis  960  (lignes OrdiniDettaglio du produit 7)',
          'Direktbereich: 18450  dann  960  (OrdiniDettaglio-Zeilen für Produkt 7)',
        ),
      },
    ],
    related: ['ac-davg', 'ac-dcount', 'ac-nz'],
  },
  {
    id: 'ac-davg',
    name: 'DAvg',
    syntax: 'DAvg(Expr, Domain, [Criteria])',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'domain',
    description: L(
      'Media aritmetica di un campo numerico nel dominio (ignora Null).',
      'Arithmetic mean of a numeric field in the domain (skips Null).',
      'Media aritmética de un campo numérico en el dominio (ignora Null).',
      'Moyenne arithmétique d’un champ numérique du domaine (ignore Null).',
      'Arithmetisches Mittel eines numerischen Feldes in der Domäne (übergeht Null).',
    ),
    examples: [
      {
        title: L('Prezzo medio per categoria', 'Average price by category', 'Precio medio por categoría', 'Prix moyen par catégorie', 'Durchschnittspreis nach Kategorie'),
        code: `Sub EsempioDAvg()
    Debug.Print Round(DAvg("PrezzoUnitario", "Prodotti", "Categoria='Bevande'"), 2)
    Debug.Print Round(DAvg("Importo", "Ordini"), 2)
End Sub`,
        result: L(
          'Finestra Immediata: 3,45  poi  612,8',
          'Immediate Window: 3.45  then  612.8',
          'Ventana Inmediato: 3,45  luego  612,8',
          'Fenêtre Exécution : 3,45  puis  612,8',
          'Direktbereich: 3,45  dann  612,8',
        ),
      },
    ],
    related: ['ac-dsum', 'ac-dmin', 'ac-dmax'],
  },
  {
    id: 'ac-dmin',
    name: 'DMin',
    syntax: 'DMin(Expr, Domain, [Criteria])',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'domain',
    description: L(
      'Valore minimo di un campo (numero, data o testo) nel dominio.',
      'Minimum value of a field (number, date or text) in the domain.',
      'Valor mínimo de un campo (número, fecha o texto) en el dominio.',
      'Valeur minimale d’un champ (nombre, date ou texte) dans le domaine.',
      'Minimalwert eines Feldes (Zahl, Datum oder Text) in der Domäne.',
    ),
    examples: [
      {
        title: L('Primo ordine e prezzo più basso', 'Earliest order and lowest price', 'Primer pedido y precio más bajo', 'Première commande et prix le plus bas', 'Früheste Bestellung und niedrigster Preis'),
        code: `Sub EsempioDMin()
    Debug.Print DMin("DataOrdine", "Ordini", "IDCliente=12")
    Debug.Print DMin("PrezzoUnitario", "Prodotti")
End Sub`,
        result: L(
          'Finestra Immediata: 14/01/2024  poi  0,85',
          'Immediate Window: 1/14/2024  then  0.85',
          'Ventana Inmediato: 14/01/2024  luego  0,85',
          'Fenêtre Exécution : 14/01/2024  puis  0,85',
          'Direktbereich: 14.01.2024  dann  0,85',
        ),
      },
    ],
    related: ['ac-dmax', 'ac-davg', 'ac-dfirst'],
  },
  {
    id: 'ac-dmax',
    name: 'DMax',
    syntax: 'DMax(Expr, Domain, [Criteria])',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'domain',
    description: L(
      'Valore massimo di un campo nel dominio.',
      'Maximum value of a field in the domain.',
      'Valor máximo de un campo en el dominio.',
      'Valeur maximale d’un champ dans le domaine.',
      'Maximalwert eines Feldes in der Domäne.',
    ),
    examples: [
      {
        title: L('Ultimo ordine e prossimo codice', 'Latest order and next code', 'Último pedido y siguiente código', 'Dernière commande et code suivant', 'Letzte Bestellung und nächster Code'),
        code: `Sub EsempioDMax()
    Debug.Print DMax("DataOrdine", "Ordini")
    Debug.Print Nz(DMax("Codice", "Clienti"), 0) + 1
End Sub`,
        result: L(
          'Finestra Immediata: 03/09/2026  poi  129  (prossimo Codice se la numerazione è lineare)',
          'Immediate Window: 9/3/2026  then  129  (next Codice if numbering is sequential)',
          'Ventana Inmediato: 03/09/2026  luego  129  (siguiente Codice si la numeración es lineal)',
          'Fenêtre Exécution : 03/09/2026  puis  129  (prochain Codice si la numérotation est linéaire)',
          'Direktbereich: 03.09.2026  dann  129  (nächster Codice bei fortlaufender Nummerierung)',
        ),
      },
    ],
    notes: L(
      'Per una chiave Autonumber usa piuttosto il record appena inserito (@@IDENTITY / Bookmark), non DMax+1 in concorrenza.',
      'For an Autonumber key prefer the just-inserted record (@@IDENTITY / Bookmark), not DMax+1 under concurrency.',
      'Para una clave Autonumber prefiere el registro recién insertado (@@IDENTITY / Bookmark), no DMax+1 en concurrencia.',
      'Pour une clé Autonumber préférez l’enregistrement venant d’être inséré (@@IDENTITY / Bookmark), pas DMax+1 en concurrence.',
      'Für Autonumber-Schlüssel den gerade eingefügten Datensatz (@@IDENTITY / Bookmark) nutzen, nicht DMax+1 bei Parallelzugriff.',
    ),
    related: ['ac-dmin', 'ac-dao-addnew', 'ac-dlookup'],
  },
  {
    id: 'ac-dfirst',
    name: 'DFirst',
    syntax: 'DFirst(Expr, Domain, [Criteria])',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'domain',
    description: L(
      'Valore della prima riga del dominio secondo l’ordine naturale della tabella/query.',
      'Value from the first row of the domain in the table/query’s natural order.',
      'Valor de la primera fila del dominio según el orden natural de la tabla/consulta.',
      'Valeur de la première ligne du domaine selon l’ordre naturel de la table/requête.',
      'Wert der ersten Zeile der Domäne in der natürlichen Reihenfolge der Tabelle/Abfrage.',
    ),
    examples: [
      {
        title: L('Primo cliente del foglio', 'First customer on the sheet', 'Primer cliente de la hoja', 'Premier client de la feuille', 'Erster Kunde im Blatt'),
        code: `Sub EsempioDFirst()
    Debug.Print DFirst("RagioneSociale", "Clienti")
    Debug.Print DFirst("RagioneSociale", "Clienti", "Citta='Milano'")
End Sub`,
        result: L(
          'Finestra Immediata: Alfa Logistica  (primo in ordine fisico)  poi  Bianchi Srl  (primo tra i milanesi)',
          'Immediate Window: Alfa Logistica  (first in physical order)  then  Bianchi Srl  (first among Milan rows)',
          'Ventana Inmediato: Alfa Logistica  (primero en orden físico)  luego  Bianchi Srl  (primero entre los de Milán)',
          'Fenêtre Exécution : Alfa Logistica  (premier en ordre physique)  puis  Bianchi Srl  (premier parmi les Milanais)',
          'Direktbereich: Alfa Logistica  (erster in physischer Reihenfolge)  dann  Bianchi Srl  (erster unter den Mailänder Zeilen)',
        ),
      },
    ],
    notes: L(
      'DFirst/DLast non ordinano: l’ordine dipende dall’indice clustered o dalla query. Per min/max usa DMin/DMax.',
      'DFirst/DLast do not sort: order depends on the clustered index or the query. For min/max use DMin/DMax.',
      'DFirst/DLast no ordenan: el orden depende del índice clustered o de la consulta. Para mín/máx usa DMin/DMax.',
      'DFirst/DLast ne trient pas : l’ordre dépend de l’index clustered ou de la requête. Pour min/max utilisez DMin/DMax.',
      'DFirst/DLast sortieren nicht: die Reihenfolge hängt vom Clustered Index oder der Abfrage ab. Für Min/Max DMin/DMax verwenden.',
    ),
    related: ['ac-dlast', 'ac-dmin', 'ac-dlookup'],
  },
  {
    id: 'ac-dlast',
    name: 'DLast',
    syntax: 'DLast(Expr, Domain, [Criteria])',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'domain',
    description: L(
      'Valore dell’ultima riga del dominio secondo l’ordine naturale.',
      'Value from the last row of the domain in natural order.',
      'Valor de la última fila del dominio según el orden natural.',
      'Valeur de la dernière ligne du domaine selon l’ordre naturel.',
      'Wert der letzten Zeile der Domäne in natürlicher Reihenfolge.',
    ),
    examples: [
      {
        title: L('Ultimo ordine inserito', 'Last order entered', 'Último pedido introducido', 'Dernière commande saisie', 'Zuletzt erfasste Bestellung'),
        code: `Sub EsempioDLast()
    Debug.Print DLast("Codice", "Ordini")
    Debug.Print DLast("Importo", "Ordini", "IDCliente=12")
End Sub`,
        result: L(
          'Finestra Immediata: 9041  poi  350  (importo dell’ultimo ordine fisico del cliente 12, non per data)',
          'Immediate Window: 9041  then  350  (amount of customer 12’s last physical order, not by date)',
          'Ventana Inmediato: 9041  luego  350  (importe del último pedido físico del cliente 12, no por fecha)',
          'Fenêtre Exécution : 9041  puis  350  (montant de la dernière commande physique du client 12, pas par date)',
          'Direktbereich: 9041  dann  350  (Betrag der letzten physischen Bestellung von Kunde 12, nicht nach Datum)',
        ),
      },
    ],
    related: ['ac-dfirst', 'ac-dmax', 'ac-dao-move-ends'],
  },
  {
    id: 'ac-dao-open-recordset',
    name: 'OpenRecordset',
    syntax: 'Database.OpenRecordset(Name, [Type], [Options], [LockEdit]) As DAO.Recordset',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'dao',
    description: L(
      'Apre un Recordset DAO su tabella, query o SQL.',
      'Opens a DAO Recordset on a table, query or SQL string.',
      'Abre un Recordset DAO sobre tabla, consulta o SQL.',
      'Ouvre un Recordset DAO sur une table, une requête ou du SQL.',
      'Öffnet ein DAO-Recordset auf Tabelle, Abfrage oder SQL.',
    ),
    params: [
      { name: 'Name', description: L('Nome tabella/query oppure SQL SELECT.', 'Table/query name or SELECT SQL.', 'Nombre de tabla/consulta o SQL SELECT.', 'Nom de table/requête ou SQL SELECT.', 'Tabellen-/Abfragename oder SELECT-SQL.') },
      { name: 'Type', optional: true, description: L('dbOpenTable, dbOpenDynaset, dbOpenSnapshot, dbOpenForwardOnly.', 'dbOpenTable, dbOpenDynaset, dbOpenSnapshot, dbOpenForwardOnly.', 'dbOpenTable, dbOpenDynaset, dbOpenSnapshot, dbOpenForwardOnly.', 'dbOpenTable, dbOpenDynaset, dbOpenSnapshot, dbOpenForwardOnly.', 'dbOpenTable, dbOpenDynaset, dbOpenSnapshot, dbOpenForwardOnly.') },
      { name: 'Options', optional: true, description: L('dbReadOnly, dbSeeChanges, dbAppendOnly…', 'dbReadOnly, dbSeeChanges, dbAppendOnly…', 'dbReadOnly, dbSeeChanges, dbAppendOnly…', 'dbReadOnly, dbSeeChanges, dbAppendOnly…', 'dbReadOnly, dbSeeChanges, dbAppendOnly…') },
      { name: 'LockEdit', optional: true, description: L('dbOptimistic, dbPessimistic.', 'dbOptimistic, dbPessimistic.', 'dbOptimistic, dbPessimistic.', 'dbOptimistic, dbPessimistic.', 'dbOptimistic, dbPessimistic.') },
    ],
    examples: [
      {
        title: L('Dynaset sui clienti di Milano', 'Dynaset of Milan customers', 'Dynaset de clientes de Milán', 'Dynaset des clients de Milan', 'Dynaset der Mailänder Kunden'),
        code: `Sub EsempioOpenRecordset()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT Codice, RagioneSociale FROM Clienti WHERE Citta='Milano'", _
        dbOpenDynaset)
    Debug.Print rs.Fields.Count
    Debug.Print rs!Codice & " " & rs!RagioneSociale
    rs.Close
End Sub`,
        result: L(
          'Recordset: 2 campi. Primo record: 12 Bianchi Srl. Tipo = dynaset (aggiornabile).',
          'Recordset: 2 fields. First record: 12 Bianchi Srl. Type = dynaset (updatable).',
          'Recordset: 2 campos. Primer registro: 12 Bianchi Srl. Tipo = dynaset (actualizable).',
          'Recordset : 2 champs. Premier enregistrement : 12 Bianchi Srl. Type = dynaset (modifiable).',
          'Recordset: 2 Felder. Erster Datensatz: 12 Bianchi Srl. Typ = Dynaset (aktualisierbar).',
        ),
      },
    ],
    notes: L(
      'dbOpenTable è obbligatorio per Seek e richiede una tabella locale, non SQL né query. Le tabelle collegate usano dbOpenDynaset.',
      'dbOpenTable is required for Seek and needs a local table, not SQL or a query. Linked tables use dbOpenDynaset.',
      'dbOpenTable es obligatorio para Seek y requiere una tabla local, no SQL ni consulta. Las vinculadas usan dbOpenDynaset.',
      'dbOpenTable est obligatoire pour Seek et exige une table locale, pas du SQL ni une requête. Les tables liées utilisent dbOpenDynaset.',
      'dbOpenTable ist für Seek nötig und verlangt eine lokale Tabelle, kein SQL und keine Abfrage. Verknüpfte Tabellen nutzen dbOpenDynaset.',
    ),
    related: ['ac-current-db', 'ac-dao-fields', 'ac-dao-clone-close', 'ac-dao-seek', 'ac-ado-recordset'],
  },
  {
    id: 'ac-dao-move-ends',
    name: 'MoveFirst / MoveLast',
    syntax: 'Recordset.MoveFirst | Recordset.MoveLast',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'dao',
    description: L(
      'Salta al primo o all’ultimo record del Recordset.',
      'Jumps to the first or last record of the Recordset.',
      'Salta al primer o al último registro del Recordset.',
      'Saute au premier ou au dernier enregistrement du Recordset.',
      'Springt zum ersten oder letzten Datensatz des Recordsets.',
    ),
    examples: [
      {
        title: L('Primo e ultimo ordine', 'First and last order', 'Primer y último pedido', 'Première et dernière commande', 'Erste und letzte Bestellung'),
        code: `Sub EsempioMoveEnds()
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("Ordini", dbOpenSnapshot)
    rs.MoveFirst
    Debug.Print rs!Codice
    rs.MoveLast
    Debug.Print rs!Codice
    Debug.Print rs.RecordCount
    rs.Close
End Sub`,
        result: L(
          'Finestra Immediata: 8001  (primo)  poi  9041  (ultimo)  poi  1041  (RecordCount valido dopo MoveLast)',
          'Immediate Window: 8001  (first)  then  9041  (last)  then  1041  (RecordCount valid after MoveLast)',
          'Ventana Inmediato: 8001  (primero)  luego  9041  (último)  luego  1041  (RecordCount válido tras MoveLast)',
          'Fenêtre Exécution : 8001  (premier)  puis  9041  (dernier)  puis  1041  (RecordCount valable après MoveLast)',
          'Direktbereich: 8001  (erster)  dann  9041  (letzter)  dann  1041  (RecordCount nach MoveLast gültig)',
        ),
      },
    ],
    notes: L(
      'Su un dynaset/snapshot RecordCount è completo solo dopo MoveLast. Su recordset vuoto MoveFirst/Last generano errore 3021.',
      'On a dynaset/snapshot RecordCount is complete only after MoveLast. On an empty recordset MoveFirst/Last raise error 3021.',
      'En un dynaset/snapshot RecordCount es completo solo tras MoveLast. En un recordset vacío MoveFirst/Last lanzan el error 3021.',
      'Sur un dynaset/snapshot RecordCount n’est complet qu’après MoveLast. Sur un recordset vide MoveFirst/Last lèvent l’erreur 3021.',
      'Bei Dynaset/Snapshot ist RecordCount erst nach MoveLast vollständig. Bei leerem Recordset lösen MoveFirst/Last Fehler 3021 aus.',
    ),
    related: ['ac-dao-move-step', 'ac-dao-bof-eof', 'ac-dao-recordcount'],
  },
  {
    id: 'ac-dao-move-step',
    name: 'MoveNext / MovePrevious',
    syntax: 'Recordset.MoveNext | Recordset.MovePrevious',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'dao',
    description: L(
      'Avanza o arretra di un record. Oltre i bordi imposta EOF o BOF.',
      'Moves one record forward or back. Past the edges it sets EOF or BOF.',
      'Avanza o retrocede un registro. Más allá de los bordes activa EOF o BOF.',
      'Avance ou recule d’un enregistrement. Au-delà des bords, EOF ou BOF est True.',
      'Geht einen Datensatz vor oder zurück. Hinter den Rändern wird EOF oder BOF gesetzt.',
    ),
    examples: [
      {
        title: L('Ciclo su tutti i prodotti', 'Loop all products', 'Bucle por todos los productos', 'Boucle sur tous les produits', 'Schleife über alle Produkte'),
        code: `Sub EsempioMoveNext()
    Dim rs As DAO.Recordset
    Dim n As Long
    Set rs = CurrentDb.OpenRecordset("SELECT Codice FROM Prodotti", dbOpenSnapshot)
    Do Until rs.EOF
        n = n + 1
        rs.MoveNext
    Loop
    Debug.Print n
    rs.MovePrevious
    Debug.Print rs!Codice
    rs.Close
End Sub`,
        result: L(
          'Dopo il ciclo EOF è True e n = 54. MovePrevious torna sull’ultimo: Codice=54.',
          'After the loop EOF is True and n = 54. MovePrevious returns to the last row: Codice=54.',
          'Tras el bucle EOF es True y n = 54. MovePrevious vuelve al último: Codice=54.',
          'Après la boucle EOF est True et n = 54. MovePrevious revient au dernier : Codice=54.',
          'Nach der Schleife ist EOF True und n = 54. MovePrevious geht zur letzten Zeile: Codice=54.',
        ),
      },
    ],
    related: ['ac-dao-move-ends', 'ac-dao-bof-eof', 'ac-dao-open-recordset'],
  },
  {
    id: 'ac-dao-addnew',
    name: 'AddNew',
    syntax: 'Recordset.AddNew',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'dao',
    description: L(
      'Prepara un nuovo record in buffer. I valori diventano permanenti solo con Update.',
      'Prepares a new record in the copy buffer. Values become permanent only with Update.',
      'Prepara un registro nuevo en el búfer. Los valores son permanentes solo con Update.',
      'Prépare un nouvel enregistrement dans le tampon. Les valeurs deviennent permanentes avec Update.',
      'Bereitet einen neuen Datensatz im Puffer vor. Werte werden erst mit Update dauerhaft.',
    ),
    examples: [
      {
        title: L('Inserire un cliente', 'Insert a customer', 'Insertar un cliente', 'Insérer un client', 'Kunden einfügen'),
        code: `Sub EsempioAddNew()
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("Clienti", dbOpenDynaset)
    rs.AddNew
    rs!RagioneSociale = "Verdi Group"
    rs!Citta = "Parma"
    rs!Email = "info@verdi.it"
    rs.Update
    rs.Bookmark = rs.LastModified
    Debug.Print rs!Codice & " " & rs!RagioneSociale
    rs.Close
End Sub`,
        result: L(
          'Nuova riga in Clienti. Dopo Update+LastModified: 129 Verdi Group. Senza Update il buffer si perde.',
          'New row in Clienti. After Update+LastModified: 129 Verdi Group. Without Update the buffer is discarded.',
          'Nueva fila en Clienti. Tras Update+LastModified: 129 Verdi Group. Sin Update se pierde el búfer.',
          'Nouvelle ligne dans Clienti. Après Update+LastModified : 129 Verdi Group. Sans Update le tampon est perdu.',
          'Neue Zeile in Clienti. Nach Update+LastModified: 129 Verdi Group. Ohne Update geht der Puffer verloren.',
        ),
      },
    ],
    related: ['ac-dao-update', 'ac-dao-edit', 'ac-dao-delete', 'ac-dmax'],
  },
  {
    id: 'ac-dao-edit',
    name: 'Edit',
    syntax: 'Recordset.Edit',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'dao',
    description: L(
      'Mette il record corrente in modalità modifica (buffer di copia).',
      'Puts the current record into edit mode (copy buffer).',
      'Pone el registro actual en modo edición (búfer de copia).',
      'Passe l’enregistrement courant en mode modification (tampon copie).',
      'Versetzt den aktuellen Datensatz in den Bearbeiten-Modus (Kopierpuffer).',
    ),
    examples: [
      {
        title: L('Cambiare la città del cliente 12', 'Change city of customer 12', 'Cambiar la ciudad del cliente 12', 'Changer la ville du client 12', 'Stadt von Kunde 12 ändern'),
        code: `Sub EsempioEdit()
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("Clienti", dbOpenDynaset)
    rs.FindFirst "Codice=12"
    rs.Edit
    rs!Citta = "Bergamo"
    rs.Update
    Debug.Print rs!Codice & " " & rs!Citta
    rs.Close
End Sub`,
        result: L(
          'Il record 12 passa da Milano a Bergamo. Senza Edit, l’assegnazione a rs!Citta solleva errore 3020.',
          'Record 12 changes from Milan to Bergamo. Without Edit, assigning rs!Citta raises error 3020.',
          'El registro 12 pasa de Milán a Bérgamo. Sin Edit, asignar rs!Citta lanza el error 3020.',
          'L’enregistrement 12 passe de Milan à Bergame. Sans Edit, l’affectation à rs!Citta lève l’erreur 3020.',
          'Datensatz 12 wechselt von Mailand nach Bergamo. Ohne Edit löst die Zuweisung an rs!Citta Fehler 3020 aus.',
        ),
      },
    ],
    related: ['ac-dao-update', 'ac-dao-addnew', 'ac-dao-find'],
  },
  {
    id: 'ac-dao-update',
    name: 'Update',
    syntax: 'Recordset.Update [MultiUserUpdate], [UpdateType]',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'dao',
    description: L(
      'Scrive sul disco il record in buffer dopo AddNew o Edit.',
      'Writes the buffered record to disk after AddNew or Edit.',
      'Escribe en disco el registro del búfer tras AddNew o Edit.',
      'Écrit sur disque l’enregistrement du tampon après AddNew ou Edit.',
      'Schreibt den gepufferten Datensatz nach AddNew oder Edit auf den Datenträger.',
    ),
    examples: [
      {
        title: L('Annullare un Edit con CancelUpdate', 'Cancel an Edit with CancelUpdate', 'Cancelar un Edit con CancelUpdate', 'Annuler un Edit avec CancelUpdate', 'Edit mit CancelUpdate verwerfen'),
        code: `Sub EsempioUpdate()
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("Prodotti", dbOpenDynaset)
    rs.FindFirst "Codice=7"
    Debug.Print rs!PrezzoUnitario
    rs.Edit
    rs!PrezzoUnitario = 99
    rs.CancelUpdate
    Debug.Print rs!PrezzoUnitario
    rs.Close
End Sub`,
        result: L(
          'Immediata: 2,5  poi  2,5. CancelUpdate scarta il 99; Update al suo posto avrebbe salvato 99.',
          'Immediate: 2.5  then  2.5. CancelUpdate discards 99; Update instead would have saved 99.',
          'Inmediato: 2,5  luego  2,5. CancelUpdate descarta el 99; Update habría guardado 99.',
          'Exécution : 2,5  puis  2,5. CancelUpdate jette le 99 ; Update à la place aurait enregistré 99.',
          'Direktbereich: 2,5  dann  2,5. CancelUpdate verwirft 99; Update hätte 99 gespeichert.',
        ),
      },
    ],
    related: ['ac-dao-addnew', 'ac-dao-edit', 'ac-me-undo-dirty'],
  },
  {
    id: 'ac-dao-delete',
    name: 'Delete',
    syntax: 'Recordset.Delete',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'dao',
    description: L(
      'Elimina il record corrente. Devi poi spostarti: il puntatore resta su un record eliminato.',
      'Deletes the current record. You must then move: the pointer stays on a deleted record.',
      'Elimina el registro actual. Luego debes moverte: el puntero queda en un registro eliminado.',
      'Supprime l’enregistrement courant. Il faut ensuite se déplacer : le pointeur reste sur un enregistrement supprimé.',
      'Löscht den aktuellen Datensatz. Danach weiterbewegen: der Zeiger bleibt auf einem gelöschten Datensatz.',
    ),
    examples: [
      {
        title: L('Cancellare le bozze', 'Delete draft rows', 'Borrar los borradores', 'Supprimer les brouillons', 'Entwürfe löschen'),
        code: `Sub EsempioDelete()
    Dim rs As DAO.Recordset
    Dim n As Long
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM Ordini WHERE Stato='Bozza'", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Delete
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    Debug.Print n
End Sub`,
        result: L(
          '3 ordini Bozza eliminati. Immediata: 3. DCount("*","Ordini","Stato=\'Bozza\'") ora è 0.',
          '3 Bozza orders deleted. Immediate: 3. DCount("*","Ordini","Stato=\'Bozza\'") is now 0.',
          '3 pedidos Bozza eliminados. Inmediato: 3. DCount("*","Ordini","Stato=\'Bozza\'") ahora es 0.',
          '3 commandes Bozza supprimées. Exécution : 3. DCount("*","Ordini","Stato=\'Bozza\'") vaut maintenant 0.',
          '3 Bozza-Bestellungen gelöscht. Direktbereich: 3. DCount("*","Ordini","Stato=\'Bozza\'") ist jetzt 0.',
        ),
      },
    ],
    related: ['ac-dao-addnew', 'ac-docmd-run-sql', 'ac-dao-move-step'],
  },
  {
    id: 'ac-dao-find',
    name: 'FindFirst / FindNext',
    syntax: 'Recordset.FindFirst Criteria | Recordset.FindNext Criteria',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'dao',
    description: L(
      'Cerca nel dynaset/snapshot. FindFirst riparte dall’inizio, FindNext continua. NoMatch dice se ha fallito.',
      'Searches a dynaset/snapshot. FindFirst starts at the beginning, FindNext continues. NoMatch reports failure.',
      'Busca en un dynaset/snapshot. FindFirst empieza al inicio, FindNext sigue. NoMatch indica el fallo.',
      'Recherche dans un dynaset/snapshot. FindFirst repart du début, FindNext continue. NoMatch signale l’échec.',
      'Sucht im Dynaset/Snapshot. FindFirst startet vorn, FindNext macht weiter. NoMatch zeigt Fehlschlag.',
    ),
    params: [
      { name: 'Criteria', description: L('Espressione come in una WHERE, es. "Citta=\'Milano\'".', 'Expression like a WHERE, e.g. "Citta=\'Milano\'".', 'Expresión como WHERE, p. ej. "Citta=\'Milano\'".', 'Expression comme un WHERE, p. ex. "Citta=\'Milano\'".', 'Ausdruck wie WHERE, z. B. "Citta=\'Milano\'".') },
    ],
    examples: [
      {
        title: L('Tutti i clienti di Milano', 'All customers in Milan', 'Todos los clientes de Milán', 'Tous les clients de Milan', 'Alle Kunden in Mailand'),
        code: `Sub EsempioFind()
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("Clienti", dbOpenDynaset)
    rs.FindFirst "Citta='Milano'"
    Do Until rs.NoMatch
        Debug.Print rs!Codice & " " & rs!RagioneSociale
        rs.FindNext "Citta='Milano'"
    Loop
    rs.Close
End Sub`,
        result: L(
          'Immediata, una riga per cliente: 12 Bianchi Srl  /  31 Rossi Spa  /  …  (22 righe). Poi NoMatch = True.',
          'Immediate, one line per customer: 12 Bianchi Srl  /  31 Rossi Spa  /  …  (22 lines). Then NoMatch = True.',
          'Inmediato, una línea por cliente: 12 Bianchi Srl  /  31 Rossi Spa  /  …  (22 líneas). Luego NoMatch = True.',
          'Exécution, une ligne par client : 12 Bianchi Srl  /  31 Rossi Spa  /  …  (22 lignes). Puis NoMatch = True.',
          'Direktbereich, eine Zeile pro Kunde: 12 Bianchi Srl  /  31 Rossi Spa  /  …  (22 Zeilen). Dann NoMatch = True.',
        ),
      },
    ],
    notes: L(
      'Find* non usa l’indice come Seek: su tabelle grandi filtra con SQL nel OpenRecordset. Non funziona su dbOpenTable.',
      'Find* does not use the index like Seek: on large tables filter with SQL in OpenRecordset. It does not work on dbOpenTable.',
      'Find* no usa el índice como Seek: en tablas grandes filtra con SQL en OpenRecordset. No funciona en dbOpenTable.',
      'Find* n’utilise pas l’index comme Seek : sur les grandes tables, filtrez en SQL dans OpenRecordset. Inutile sur dbOpenTable.',
      'Find* nutzt den Index nicht wie Seek: große Tabellen per SQL in OpenRecordset filtern. Funktioniert nicht mit dbOpenTable.',
    ),
    related: ['ac-dao-seek', 'ac-dlookup', 'ac-docmd-find-record', 'ac-dao-bof-eof'],
  },
  {
    id: 'ac-dao-seek',
    name: 'Seek',
    syntax: 'Recordset.Seek Comparison, Key1, [Key2], [Key3]…',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'dao',
    description: L(
      'Ricerca per indice su un Recordset dbOpenTable. Imposta prima la proprietà Index.',
      'Index search on a dbOpenTable Recordset. Set the Index property first.',
      'Búsqueda por índice en un Recordset dbOpenTable. Asigna antes la propiedad Index.',
      'Recherche par index sur un Recordset dbOpenTable. Définissez d’abord la propriété Index.',
      'Indexsuche auf einem dbOpenTable-Recordset. Zuerst die Eigenschaft Index setzen.',
    ),
    params: [
      { name: 'Comparison', description: L('"=", ">=", ">", "<=", "<".', '"=", ">=", ">", "<=", "<".', '"=", ">=", ">", "<=", "<".', '"=", ">=", ">", "<=", "<".', '"=", ">=", ">", "<=", "<".') },
      { name: 'Key1', description: L('Valore della prima colonna dell’indice.', 'Value of the first index column.', 'Valor de la primera columna del índice.', 'Valeur de la première colonne de l’index.', 'Wert der ersten Indexspalte.') },
    ],
    examples: [
      {
        title: L('Seek sulla chiave primaria', 'Seek on the primary key', 'Seek en la clave primaria', 'Seek sur la clé primaire', 'Seek auf dem Primärschlüssel'),
        code: `Sub EsempioSeek()
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("Clienti", dbOpenTable)
    rs.Index = "PrimaryKey"
    rs.Seek "=", 12
    If rs.NoMatch Then
        Debug.Print "mancante"
    Else
        Debug.Print rs!RagioneSociale
    End If
    rs.Close
End Sub`,
        result: L(
          'Finestra Immediata: Bianchi Srl  (indice PrimaryKey, Codice=12). Su tabella collegata Seek non è disponibile.',
          'Immediate Window: Bianchi Srl  (PrimaryKey index, Codice=12). Seek is not available on a linked table.',
          'Ventana Inmediato: Bianchi Srl  (índice PrimaryKey, Codice=12). Seek no está disponible en una tabla vinculada.',
          'Fenêtre Exécution : Bianchi Srl  (index PrimaryKey, Codice=12). Seek n’est pas disponible sur une table liée.',
          'Direktbereich: Bianchi Srl  (PrimaryKey-Index, Codice=12). Seek ist auf einer verknüpften Tabelle nicht verfügbar.',
        ),
      },
    ],
    related: ['ac-dao-find', 'ac-dao-open-recordset', 'ac-dlookup'],
  },
  {
    id: 'ac-dao-clone-close',
    name: 'Clone / Recordset.Close',
    syntax: 'Set rs2 = Recordset.Clone | Recordset.Close',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'dao',
    description: L(
      'Clone crea un secondo cursore sullo stesso set (Bookmark compatibili). Close rilascia il Recordset.',
      'Clone creates a second cursor on the same set (compatible Bookmarks). Close releases the Recordset.',
      'Clone crea un segundo cursor sobre el mismo conjunto (Bookmark compatibles). Close libera el Recordset.',
      'Clone crée un second curseur sur le même ensemble (signets compatibles). Close libère le Recordset.',
      'Clone erzeugt einen zweiten Cursor auf derselben Menge (kompatible Bookmarks). Close gibt das Recordset frei.',
    ),
    examples: [
      {
        title: L('Segnalibro condiviso e chiusura', 'Shared bookmark and close', 'Marcador compartido y cierre', 'Signet partagé et fermeture', 'Gemeinsames Bookmark und Schließen'),
        code: `Sub EsempioCloneClose()
    Dim rs As DAO.Recordset
    Dim clone As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("Clienti", dbOpenDynaset)
    rs.FindFirst "Codice=12"
    Set clone = rs.Clone
    clone.Bookmark = rs.Bookmark
    Debug.Print clone!RagioneSociale
    clone.Close
    rs.Close
    Debug.Print "chiusi"
End Sub`,
        result: L(
          'Il clone è sullo stesso cliente: Bianchi Srl. Poi entrambi Close; Immediata: chiusi. Non usare il Recordset dopo Close.',
          'The clone is on the same customer: Bianchi Srl. Then both Close; Immediate: chiusi. Do not use the Recordset after Close.',
          'El clon está en el mismo cliente: Bianchi Srl. Luego ambos Close; Inmediato: chiusi. No uses el Recordset después de Close.',
          'Le clone est sur le même client : Bianchi Srl. Puis les deux Close ; Exécution : chiusi. N’utilisez plus le Recordset après Close.',
          'Der Klon steht auf demselben Kunden: Bianchi Srl. Dann beide Close; Direktbereich: chiusi. Recordset nach Close nicht weiter verwenden.',
        ),
      },
    ],
    related: ['ac-dao-open-recordset', 'ac-dao-find', 'ac-dao-fields'],
  },
  {
    id: 'ac-dao-bof-eof',
    name: 'BOF / EOF',
    syntax: 'Recordset.BOF As Boolean | Recordset.EOF As Boolean',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'dao',
    description: L(
      'BOF è True prima del primo record, EOF dopo l’ultimo. Entrambi True = set vuoto.',
      'BOF is True before the first record, EOF after the last. Both True = empty set.',
      'BOF es True antes del primero, EOF después del último. Ambos True = conjunto vacío.',
      'BOF est True avant le premier, EOF après le dernier. Les deux True = ensemble vide.',
      'BOF ist True vor dem ersten Datensatz, EOF nach dem letzten. Beide True = leere Menge.',
    ),
    examples: [
      {
        title: L('Test set vuoto e fine ciclo', 'Empty-set test and loop end', 'Prueba de conjunto vacío y fin de bucle', 'Test d’ensemble vide et fin de boucle', 'Leere Menge prüfen und Schleifenende'),
        code: `Sub EsempioBofEof()
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM Clienti WHERE Codice=0", dbOpenSnapshot)
    Debug.Print rs.BOF
    Debug.Print rs.EOF
    rs.Close
    Set rs = CurrentDb.OpenRecordset("Clienti", dbOpenSnapshot)
    rs.MoveLast
    rs.MoveNext
    Debug.Print rs.EOF
    rs.Close
End Sub`,
        result: L(
          'Immediata: True  True  (nessun Codice=0)  poi  True  (oltre l’ultimo cliente).',
          'Immediate: True  True  (no Codice=0)  then  True  (past the last customer).',
          'Inmediato: True  True  (no hay Codice=0)  luego  True  (más allá del último cliente).',
          'Exécution : True  True  (aucun Codice=0)  puis  True  (au-delà du dernier client).',
          'Direktbereich: True  True  (kein Codice=0)  dann  True  (hinter dem letzten Kunden).',
        ),
      },
    ],
    related: ['ac-dao-move-step', 'ac-dao-move-ends', 'ac-dao-find'],
  },
  {
    id: 'ac-dao-recordcount',
    name: 'RecordCount',
    syntax: 'Recordset.RecordCount As Long',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'dao',
    description: L(
      'Numero di record visitati. Su dynaset/snapshot chiama MoveLast per il totale reale.',
      'Number of records visited. On a dynaset/snapshot call MoveLast for the true total.',
      'Número de registros visitados. En dynaset/snapshot llama MoveLast para el total real.',
      'Nombre d’enregistrements visités. Sur dynaset/snapshot appelez MoveLast pour le total réel.',
      'Anzahl der besuchten Datensätze. Bei Dynaset/Snapshot MoveLast für die echte Gesamtzahl aufrufen.',
    ),
    examples: [
      {
        title: L('Conteggio prima e dopo MoveLast', 'Count before and after MoveLast', 'Recuento antes y después de MoveLast', 'Comptage avant et après MoveLast', 'Zählung vor und nach MoveLast'),
        code: `Sub EsempioRecordCount()
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("Clienti", dbOpenDynaset)
    Debug.Print rs.RecordCount
    rs.MoveLast
    Debug.Print rs.RecordCount
    rs.Close
End Sub`,
        result: L(
          'Finestra Immediata: 1  (appena aperto)  poi  128  (dopo MoveLast). dbOpenTable dà subito il totale.',
          'Immediate Window: 1  (just opened)  then  128  (after MoveLast). dbOpenTable gives the total immediately.',
          'Ventana Inmediato: 1  (recién abierto)  luego  128  (tras MoveLast). dbOpenTable da el total al instante.',
          'Fenêtre Exécution : 1  (vient d’ouvrir)  puis  128  (après MoveLast). dbOpenTable donne le total tout de suite.',
          'Direktbereich: 1  (gerade geöffnet)  dann  128  (nach MoveLast). dbOpenTable liefert die Gesamtzahl sofort.',
        ),
      },
    ],
    related: ['ac-dao-move-ends', 'ac-dcount', 'ac-dao-open-recordset'],
  },
  {
    id: 'ac-dao-fields',
    name: 'Fields',
    syntax: 'Recordset.Fields[(NameOrIndex)] As DAO.Field',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'dao',
    description: L(
      'Collezione dei campi: nome, valore, tipo e attributi del record corrente.',
      'Field collection: name, value, type and attributes of the current record.',
      'Colección de campos: nombre, valor, tipo y atributos del registro actual.',
      'Collection des champs : nom, valeur, type et attributs de l’enregistrement courant.',
      'Feldsammlung: Name, Wert, Typ und Attribute des aktuellen Datensatzes.',
    ),
    examples: [
      {
        title: L('Elencare nome e valore', 'List name and value', 'Listar nombre y valor', 'Lister nom et valeur', 'Name und Wert auflisten'),
        code: `Sub EsempioFields()
    Dim rs As DAO.Recordset
    Dim fld As DAO.Field
    Set rs = CurrentDb.OpenRecordset("SELECT Codice, RagioneSociale, Citta FROM Clienti WHERE Codice=12")
    For Each fld In rs.Fields
        Debug.Print fld.Name & "=" & fld.Value
    Next
    Debug.Print rs.Fields("Citta").Type
    rs.Close
End Sub`,
        result: L(
          'Immediata: Codice=12  /  RagioneSociale=Bianchi Srl  /  Citta=Milano  /  10  (dbText). Equivalente: rs!Citta.',
          'Immediate: Codice=12  /  RagioneSociale=Bianchi Srl  /  Citta=Milano  /  10  (dbText). Equivalent: rs!Citta.',
          'Inmediato: Codice=12  /  RagioneSociale=Bianchi Srl  /  Citta=Milano  /  10  (dbText). Equivalente: rs!Citta.',
          'Exécution : Codice=12  /  RagioneSociale=Bianchi Srl  /  Citta=Milano  /  10  (dbText). Équivalent : rs!Citta.',
          'Direktbereich: Codice=12  /  RagioneSociale=Bianchi Srl  /  Citta=Milano  /  10  (dbText). Entsprechung: rs!Citta.',
        ),
      },
    ],
    related: ['ac-dao-open-recordset', 'ac-dao-edit', 'ac-forms-bang'],
  },
  {
    id: 'ac-dao-requery',
    name: 'Recordset.Requery',
    syntax: 'Recordset.Requery [NewQueryDef]',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'dao',
    description: L(
      'Riesegue la query del Recordset e si riposiziona sul primo record.',
      'Re-runs the Recordset query and moves back to the first record.',
      'Vuelve a ejecutar la consulta del Recordset y vuelve al primer registro.',
      'Ré-exécute la requête du Recordset et se replace sur le premier enregistrement.',
      'Führt die Recordset-Abfrage erneut aus und stellt auf den ersten Datensatz zurück.',
    ),
    examples: [
      {
        title: L('Ricaricare dopo un INSERT', 'Reload after an INSERT', 'Recargar tras un INSERT', 'Recharger après un INSERT', 'Nach INSERT neu laden'),
        code: `Sub EsempioDaoRequery()
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM Clienti WHERE Citta='Parma'", dbOpenDynaset)
    rs.MoveLast
    Debug.Print rs.RecordCount
    CurrentDb.Execute "INSERT INTO Clienti (RagioneSociale, Citta) VALUES ('Lamiere Nord','Parma')"
    rs.Requery
    rs.MoveLast
    Debug.Print rs.RecordCount
    rs.Close
End Sub`,
        result: L(
          'Immediata: 4  poi  5. Senza Requery il Recordset non vedrebbe Lamiere Nord.',
          'Immediate: 4  then  5. Without Requery the Recordset would not see Lamiere Nord.',
          'Inmediato: 4  luego  5. Sin Requery el Recordset no vería Lamiere Nord.',
          'Exécution : 4  puis  5. Sans Requery le Recordset ne verrait pas Lamiere Nord.',
          'Direktbereich: 4  dann  5. Ohne Requery sähe das Recordset Lamiere Nord nicht.',
        ),
      },
    ],
    related: ['ac-docmd-requery', 'ac-me-requery', 'ac-dao-open-recordset'],
  },
  {
    id: 'ac-dao-querydef-tabledef',
    name: 'QueryDef / TableDef',
    syntax: 'Database.QueryDefs(Name) As DAO.QueryDef | Database.TableDefs(Name) As DAO.TableDef',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'dao',
    description: L(
      'QueryDef è una query salvata (SQL, parametri). TableDef è la definizione di una tabella (campi, indici, connect).',
      'QueryDef is a saved query (SQL, parameters). TableDef is a table definition (fields, indexes, connect).',
      'QueryDef es una consulta guardada (SQL, parámetros). TableDef es la definición de una tabla (campos, índices, connect).',
      'QueryDef est une requête enregistrée (SQL, paramètres). TableDef est la définition d’une table (champs, index, connect).',
      'QueryDef ist eine gespeicherte Abfrage (SQL, Parameter). TableDef ist die Tabellendefinition (Felder, Indizes, Connect).',
    ),
    examples: [
      {
        title: L('SQL della query e campi della tabella', 'Query SQL and table fields', 'SQL de la consulta y campos de la tabla', 'SQL de la requête et champs de la table', 'Abfrage-SQL und Tabellenfelder'),
        code: `Sub EsempioDefs()
    Dim qdf As DAO.QueryDef
    Dim tdf As DAO.TableDef
    Set qdf = CurrentDb.QueryDefs("qryClientiMilano")
    Debug.Print qdf.SQL
    Set tdf = CurrentDb.TableDefs("Clienti")
    Debug.Print tdf.Fields.Count
    Debug.Print tdf.Connect
End Sub`,
        result: L(
          'Immediata: SELECT … FROM Clienti WHERE Citta="Milano";  poi  9  (campi)  poi  stringa vuota (tabella locale).',
          'Immediate: SELECT … FROM Clienti WHERE Citta="Milano";  then  9  (fields)  then  empty string (local table).',
          'Inmediato: SELECT … FROM Clienti WHERE Citta="Milano";  luego  9  (campos)  luego  cadena vacía (tabla local).',
          'Exécution : SELECT … FROM Clienti WHERE Citta="Milano";  puis  9  (champs)  puis  chaîne vide (table locale).',
          'Direktbereich: SELECT … FROM Clienti WHERE Citta="Milano";  dann  9  (Felder)  dann  leere Zeichenfolge (lokale Tabelle).',
        ),
      },
      {
        title: L('Query parametrica eseguita', 'Run a parameter query', 'Ejecutar una consulta con parámetros', 'Exécuter une requête paramétrée', 'Parameterabfrage ausführen'),
        code: `Sub EsempioQueryDefParam()
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Set qdf = CurrentDb.CreateQueryDef("")
    qdf.SQL = "PARAMETERS pCitta Text; SELECT * FROM Clienti WHERE Citta=[pCitta]"
    qdf.Parameters("pCitta") = "Torino"
    Set rs = qdf.OpenRecordset
    Debug.Print rs.RecordCount
    rs.Close
End Sub`,
        result: L(
          'Recordset temporaneo con i 11 clienti di Torino. CreateQueryDef("") non salva la query nel contenitore.',
          'Temporary recordset with the 11 Turin customers. CreateQueryDef("") does not save the query in the container.',
          'Recordset temporal con los 11 clientes de Turín. CreateQueryDef("") no guarda la consulta en el contenedor.',
          'Recordset temporaire avec les 11 clients de Turin. CreateQueryDef("") n’enregistre pas la requête dans le conteneur.',
          'Temporäres Recordset mit den 11 Turiner Kunden. CreateQueryDef("") speichert die Abfrage nicht im Container.',
        ),
      },
    ],
    related: ['ac-current-db', 'ac-dao-open-recordset', 'ac-docmd-open-query'],
  },
  {
    id: 'ac-ado-connection',
    name: 'ADODB.Connection / Connection.Open',
    syntax: 'Dim cn As New ADODB.Connection | cn.Open ConnectionString, [UserID], [Password], [Options]',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'ado',
    description: L(
      'Connessione ADO. Open apre il provider (ACE, SQL Server, ecc.) con una connection string.',
      'ADO connection. Open opens the provider (ACE, SQL Server, etc.) with a connection string.',
      'Conexión ADO. Open abre el proveedor (ACE, SQL Server, etc.) con una cadena de conexión.',
      'Connexion ADO. Open ouvre le fournisseur (ACE, SQL Server, etc.) avec une chaîne de connexion.',
      'ADO-Verbindung. Open öffnet den Provider (ACE, SQL Server usw.) mit einer Verbindungszeichenfolge.',
    ),
    params: [
      { name: 'ConnectionString', description: L('Provider, Data Source, credenziali.', 'Provider, Data Source, credentials.', 'Provider, Data Source, credenciales.', 'Provider, Data Source, identifiants.', 'Provider, Data Source, Anmeldeinformationen.') },
    ],
    examples: [
      {
        title: L('Aprire il database corrente via ACE', 'Open the current database via ACE', 'Abrir la base actual vía ACE', 'Ouvrir la base courante via ACE', 'Aktuelle Datenbank über ACE öffnen'),
        code: `Sub EsempioAdoOpen()
    Dim cn As ADODB.Connection
    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & CurrentProject.FullName
    Debug.Print cn.State
    Debug.Print cn.Provider
    cn.Close
End Sub`,
        result: L(
          'Immediata: 1  (adStateOpen)  poi  Microsoft.ACE.OLEDB.12.0. Richiede il riferimento Microsoft ActiveX Data Objects.',
          'Immediate: 1  (adStateOpen)  then  Microsoft.ACE.OLEDB.12.0. Requires the Microsoft ActiveX Data Objects reference.',
          'Inmediato: 1  (adStateOpen)  luego  Microsoft.ACE.OLEDB.12.0. Requiere la referencia Microsoft ActiveX Data Objects.',
          'Exécution : 1  (adStateOpen)  puis  Microsoft.ACE.OLEDB.12.0. Nécessite la référence Microsoft ActiveX Data Objects.',
          'Direktbereich: 1  (adStateOpen)  dann  Microsoft.ACE.OLEDB.12.0. Erfordert den Verweis Microsoft ActiveX Data Objects.',
        ),
      },
    ],
    notes: L(
      'In Access puoi anche usare CurrentProject.Connection (già aperta sul file corrente) invece di New + Open.',
      'In Access you can also use CurrentProject.Connection (already open on the current file) instead of New + Open.',
      'En Access también puedes usar CurrentProject.Connection (ya abierta sobre el archivo actual) en lugar de New + Open.',
      'Dans Access vous pouvez aussi utiliser CurrentProject.Connection (déjà ouverte sur le fichier courant) au lieu de New + Open.',
      'In Access können Sie auch CurrentProject.Connection verwenden (bereits auf der aktuellen Datei geöffnet) statt New + Open.',
    ),
    related: ['ac-ado-execute', 'ac-ado-recordset', 'ac-ado-command', 'ac-current-project'],
  },
  {
    id: 'ac-ado-command',
    name: 'ADODB.Command',
    syntax: 'Dim cmd As New ADODB.Command | cmd.CommandText = sql | cmd.Execute',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'ado',
    description: L(
      'Comando ADO parametrizzato (testo SQL o stored procedure) legato a una Connection.',
      'Parameterized ADO command (SQL text or stored procedure) bound to a Connection.',
      'Comando ADO parametrizado (texto SQL o procedimiento) ligado a una Connection.',
      'Commande ADO paramétrée (texte SQL ou procédure stockée) liée à une Connection.',
      'Parametrierter ADO-Befehl (SQL-Text oder gespeicherte Prozedur) an eine Connection gebunden.',
    ),
    examples: [
      {
        title: L('UPDATE con parametro', 'Parameterized UPDATE', 'UPDATE con parámetro', 'UPDATE paramétré', 'UPDATE mit Parameter'),
        code: `Sub EsempioAdoCommand()
    Dim cmd As ADODB.Command
    Dim n As Long
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = CurrentProject.Connection
        .CommandType = adCmdText
        .CommandText = "UPDATE Clienti SET Telefono=? WHERE Codice=?"
        .Parameters.Append .CreateParameter("pTel", adVarWChar, adParamInput, 30, "011-5551234")
        .Parameters.Append .CreateParameter("pId", adInteger, adParamInput, , 12)
        .Execute n
    End With
    Debug.Print n
End Sub`,
        result: L(
          'Il cliente 12 ha Telefono=011-5551234. Immediata: 1  (RecordsAffected). I ? rispettano l’ordine di Append.',
          'Customer 12 has Telefono=011-5551234. Immediate: 1  (RecordsAffected). The ? follow Append order.',
          'El cliente 12 tiene Telefono=011-5551234. Inmediato: 1  (RecordsAffected). Los ? siguen el orden de Append.',
          'Le client 12 a Telefono=011-5551234. Exécution : 1  (RecordsAffected). Les ? suivent l’ordre d’Append.',
          'Kunde 12 hat Telefono=011-5551234. Direktbereich: 1  (RecordsAffected). Die ? folgen der Append-Reihenfolge.',
        ),
      },
    ],
    related: ['ac-ado-connection', 'ac-ado-execute', 'ac-ado-recordset'],
  },
  {
    id: 'ac-ado-recordset',
    name: 'ADODB.Recordset',
    syntax: 'Dim rs As New ADODB.Recordset | rs.Open Source, ActiveConnection, [CursorType], [LockType]',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'ado',
    description: L(
      'Cursore ADO: Open su SQL o tabella, Move, Fields, Update; diverso dal Recordset DAO.',
      'ADO cursor: Open on SQL or a table, Move, Fields, Update; different from a DAO Recordset.',
      'Cursor ADO: Open sobre SQL o tabla, Move, Fields, Update; distinto del Recordset DAO.',
      'Curseur ADO : Open sur SQL ou table, Move, Fields, Update ; distinct du Recordset DAO.',
      'ADO-Cursor: Open auf SQL oder Tabelle, Move, Fields, Update; anders als ein DAO-Recordset.',
    ),
    examples: [
      {
        title: L('Leggere e aggiornare un prodotto', 'Read and update a product', 'Leer y actualizar un producto', 'Lire et mettre à jour un produit', 'Produkt lesen und aktualisieren'),
        code: `Sub EsempioAdoRecordset()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "SELECT Codice, PrezzoUnitario FROM Prodotti WHERE Codice=7", _
        CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    Debug.Print rs.Fields("PrezzoUnitario").Value
    rs!PrezzoUnitario = rs!PrezzoUnitario + 0.1
    rs.Update
    Debug.Print rs!PrezzoUnitario
    rs.Close
End Sub`,
        result: L(
          'Immediata: 2,5  poi  2,6. La riga prodotto 7 è salvata. CursorType/LockType vanno scelti per l’aggiornabilità.',
          'Immediate: 2.5  then  2.6. Product 7 is saved. CursorType/LockType must allow updates.',
          'Inmediato: 2,5  luego  2,6. La fila del producto 7 queda guardada. CursorType/LockType deben permitir actualizar.',
          'Exécution : 2,5  puis  2,6. La ligne du produit 7 est enregistrée. CursorType/LockType doivent autoriser la mise à jour.',
          'Direktbereich: 2,5  dann  2,6. Produkt 7 ist gespeichert. CursorType/LockType müssen Updates erlauben.',
        ),
      },
    ],
    related: ['ac-ado-connection', 'ac-dao-open-recordset', 'ac-ado-command'],
  },
  {
    id: 'ac-ado-execute',
    name: 'Connection.Execute',
    syntax: 'Connection.Execute CommandText, [RecordsAffected], [Options]',
    scope: ['access'],
    category: 'ac-data',
    subcategory: 'ado',
    description: L(
      'Esegue SQL di comando (o una procedura) sulla Connection e può restituire RecordsAffected.',
      'Runs action SQL (or a procedure) on the Connection and can return RecordsAffected.',
      'Ejecuta SQL de acción (o un procedimiento) en la Connection y puede devolver RecordsAffected.',
      'Exécute du SQL Action (ou une procédure) sur la Connection et peut renvoyer RecordsAffected.',
      'Führt Aktions-SQL (oder eine Prozedur) auf der Connection aus und kann RecordsAffected zurückgeben.',
    ),
    params: [
      { name: 'CommandText', description: L('SQL o nome procedura.', 'SQL or procedure name.', 'SQL o nombre de procedimiento.', 'SQL ou nom de procédure.', 'SQL oder Prozedurname.') },
      { name: 'RecordsAffected', optional: true, description: L('Variabile Long riempita con le righe toccate.', 'Long variable filled with rows touched.', 'Variable Long con las filas afectadas.', 'Variable Long remplie avec les lignes touchées.', 'Long-Variable mit den betroffenen Zeilen.') },
      { name: 'Options', optional: true, description: L('adCmdText, adCmdStoredProc, adExecuteNoRecords…', 'adCmdText, adCmdStoredProc, adExecuteNoRecords…', 'adCmdText, adCmdStoredProc, adExecuteNoRecords…', 'adCmdText, adCmdStoredProc, adExecuteNoRecords…', 'adCmdText, adCmdStoredProc, adExecuteNoRecords…') },
    ],
    examples: [
      {
        title: L('DELETE con conteggio', 'DELETE with row count', 'DELETE con recuento', 'DELETE avec comptage', 'DELETE mit Zeilenzahl'),
        code: `Sub EsempioAdoExecute()
    Dim n As Long
    CurrentProject.Connection.Execute _
        "DELETE FROM Ordini WHERE Stato='Bozza'", n, adCmdText + adExecuteNoRecords
    Debug.Print n
End Sub`,
        result: L(
          'Le bozze spariscono. Immediata: 3. adExecuteNoRecords evita un Recordset vuoto di ritorno.',
          'Drafts disappear. Immediate: 3. adExecuteNoRecords avoids a leftover empty Recordset.',
          'Los borradores desaparecen. Inmediato: 3. adExecuteNoRecords evita un Recordset vacío de retorno.',
          'Les brouillons disparaissent. Exécution : 3. adExecuteNoRecords évite un Recordset vide en retour.',
          'Entwürfe verschwinden. Direktbereich: 3. adExecuteNoRecords vermeidet ein leeres Rückgabe-Recordset.',
        ),
      },
    ],
    related: ['ac-ado-connection', 'ac-ado-command', 'ac-docmd-run-sql'],
  },
  {
    id: 'ac-forms-bang',
    name: 'Forms!Name',
    syntax: 'Forms!FormName!ControlName  |  Forms("FormName")("ControlName")',
    scope: ['access'],
    category: 'ac-ui',
    subcategory: 'forms',
    description: L(
      'Riferimento a una maschera aperta e ai suoi controlli dalla collezione Forms.',
      'Reference to an open form and its controls via the Forms collection.',
      'Referencia a un formulario abierto y a sus controles mediante la colección Forms.',
      'Référence à un formulaire ouvert et à ses contrôles via la collection Forms.',
      'Verweis auf ein geöffnetes Formular und seine Steuerelemente über die Forms-Auflistung.',
    ),
    examples: [
      {
        title: L('Leggere e scrivere un controllo aperto', 'Read and write an open control', 'Leer y escribir un control abierto', 'Lire et écrire un contrôle ouvert', 'Offenes Steuerelement lesen und schreiben'),
        code: `Sub EsempioFormsBang()
    DoCmd.OpenForm "frmClienti"
    Debug.Print Forms!frmClienti!RagioneSociale
    Forms!frmClienti!txtNote = "chiamare lunedì"
    Debug.Print Forms("frmClienti").Caption
End Sub`,
        result: L(
          'Immediata: Bianchi Srl  poi  Clienti  (Caption). Sul form, txtNote mostra "chiamare lunedì" e Me.Dirty diventa True.',
          'Immediate: Bianchi Srl  then  Clienti  (Caption). On the form, txtNote shows "chiamare lunedì" and Me.Dirty becomes True.',
          'Inmediato: Bianchi Srl  luego  Clienti  (Caption). En el formulario, txtNote muestra "chiamare lunedì" y Me.Dirty pasa a True.',
          'Exécution : Bianchi Srl  puis  Clienti  (Caption). Sur le formulaire, txtNote affiche "chiamare lunedì" et Me.Dirty passe à True.',
          'Direktbereich: Bianchi Srl  dann  Clienti  (Caption). Auf dem Formular zeigt txtNote "chiamare lunedì" und Me.Dirty wird True.',
        ),
      },
    ],
    notes: L(
      'Se la maschera è chiusa, Forms!frmClienti solleva errore 2450. Apri prima con OpenForm o testa SysCmd(acSysCmdGetObjectState).',
      'If the form is closed, Forms!frmClienti raises error 2450. Open it with OpenForm first or test SysCmd(acSysCmdGetObjectState).',
      'Si el formulario está cerrado, Forms!frmClienti lanza el error 2450. Ábrelo antes con OpenForm o prueba SysCmd(acSysCmdGetObjectState).',
      'Si le formulaire est fermé, Forms!frmClienti lève l’erreur 2450. Ouvrez-le d’abord avec OpenForm ou testez SysCmd(acSysCmdGetObjectState).',
      'Ist das Formular geschlossen, löst Forms!frmClienti Fehler 2450 aus. Zuerst mit OpenForm öffnen oder SysCmd(acSysCmdGetObjectState) prüfen.',
    ),
    related: ['ac-me', 'ac-docmd-open-form', 'ac-screen-active', 'ac-reports-bang'],
  },
  {
    id: 'ac-me',
    name: 'Me',
    syntax: 'Me  (modulo di classe della maschera o del report)',
    scope: ['access'],
    category: 'ac-ui',
    subcategory: 'forms',
    description: L(
      'Riferimento all’istanza della maschera o del report in cui gira il codice.',
      'Reference to the form or report instance where the code is running.',
      'Referencia a la instancia del formulario o informe donde corre el código.',
      'Référence à l’instance du formulaire ou de l’état où s’exécute le code.',
      'Verweis auf die Formular- oder Berichtsinstanz, in der der Code läuft.',
    ),
    examples: [
      {
        title: L('Modulo della maschera frmClienti', 'frmClienti form module', 'Módulo del formulario frmClienti', 'Module du formulaire frmClienti', 'Formularmodul von frmClienti'),
        code: `' Inserire nel modulo di frmClienti (non in un modulo standard).
Private Sub Form_Current()
    Me.Caption = "Cliente " & Nz(Me!Codice, "(nuovo)")
    Debug.Print Me.Name
    Debug.Print Me!RagioneSociale
End Sub`,
        result: L(
          'Cambiando record la didascalia diventa "Cliente 12". Immediata: frmClienti  /  Bianchi Srl. Me non esiste in un modulo standard.',
          'Changing record sets the caption to "Cliente 12". Immediate: frmClienti  /  Bianchi Srl. Me does not exist in a standard module.',
          'Al cambiar de registro el título pasa a "Cliente 12". Inmediato: frmClienti  /  Bianchi Srl. Me no existe en un módulo estándar.',
          'En changeant d’enregistrement le libellé devient "Cliente 12". Exécution : frmClienti  /  Bianchi Srl. Me n’existe pas dans un module standard.',
          'Beim Datensatzwechsel wird die Beschriftung "Cliente 12". Direktbereich: frmClienti  /  Bianchi Srl. Me gibt es in einem Standardmodul nicht.',
        ),
      },
    ],
    related: ['ac-forms-bang', 'ac-code-context-object', 'ac-me-requery', 'ac-open-args'],
  },
  {
    id: 'ac-me-requery',
    name: 'Me.Requery / ComboBox.Requery',
    syntax: 'Me.Requery | Me!cboNome.Requery',
    scope: ['access'],
    category: 'ac-ui',
    subcategory: 'forms',
    description: L(
      'Riesegue l’origine della maschera o della combo/listbox. La combo non si aggiorna da sola se cambia la sua RowSource.',
      'Re-runs the form’s or combo/list box’s data source. A combo does not refresh itself when its RowSource data changes.',
      'Vuelve a ejecutar el origen del formulario o del combo/lista. El combo no se actualiza solo si cambian los datos de RowSource.',
      'Ré-exécute la source du formulaire ou de la liste. Une liste ne se rafraîchit pas toute seule si sa RowSource change.',
      'Führt die Datenquelle des Formulars oder des Kombinations-/Listenfelds erneut aus. Ein Kombi aktualisiert sich nicht selbst, wenn sich die RowSource-Daten ändern.',
    ),
    examples: [
      {
        title: L('Ricaricare la maschera', 'Reload the form', 'Recargar el formulario', 'Recharger le formulaire', 'Formular neu laden'),
        code: `' Modulo di frmOrdini
Private Sub cmdAggiorna_Click()
    Me.Requery
    Debug.Print Me.Recordset.RecordCount
End Sub`,
        result: L(
          'frmOrdini ricarica gli ordini (es. da 40 a 41 se ne è stato aggiunto uno). Il record corrente torna al primo del set.',
          'frmOrdini reloads orders (e.g. from 40 to 41 if one was added). The current record jumps back to the first in the set.',
          'frmOrdini recarga los pedidos (p. ej. de 40 a 41 si se añadió uno). El registro actual vuelve al primero del conjunto.',
          'frmOrdini recharge les commandes (p. ex. de 40 à 41 si une a été ajoutée). L’enregistrement courant revient au premier du jeu.',
          'frmOrdini lädt Bestellungen neu (z. B. von 40 auf 41, wenn eine hinzukam). Der aktuelle Datensatz springt zum ersten der Menge.',
        ),
      },
      {
        title: L('Aggiornare la combo città', 'Requery the city combo', 'Reconsultar el combo de ciudad', 'Réactualiser la liste des villes', 'Städte-Kombi neu abfragen'),
        code: `' Modulo di frmClienti. cboCitta ha RowSource = SELECT DISTINCT Citta FROM Clienti
Private Sub Form_AfterInsert()
    Me!cboCitta.Requery
    Debug.Print Me!cboCitta.ListCount
End Sub`,
        result: L(
          'Dopo un inserimento a Rimini, cboCitta elenca anche Rimini. ListCount passa da 18 a 19.',
          'After an insert in Rimini, cboCitta also lists Rimini. ListCount goes from 18 to 19.',
          'Tras un alta en Rimini, cboCitta también lista Rimini. ListCount pasa de 18 a 19.',
          'Après une insertion à Rimini, cboCitta liste aussi Rimini. ListCount passe de 18 à 19.',
          'Nach einem Insert in Rimini listet cboCitta auch Rimini. ListCount steigt von 18 auf 19.',
        ),
      },
    ],
    related: ['ac-docmd-requery', 'ac-me-refresh-recalc', 'ac-dao-requery', 'ac-me'],
  },
  {
    id: 'ac-me-refresh-recalc',
    name: 'Me.Refresh / Me.Recalc',
    syntax: 'Me.Refresh | Me.Recalc',
    scope: ['access'],
    category: 'ac-ui',
    subcategory: 'forms',
    description: L(
      'Refresh rilegge i record già caricati (non aggiunge le righe nuove). Recalc ricalcola i controlli calcolati.',
      'Refresh rereads already-loaded records (does not add new rows). Recalc recalculates calculated controls.',
      'Refresh relee los registros ya cargados (no añade filas nuevas). Recalc recalcula los controles calculados.',
      'Refresh relit les enregistrements déjà chargés (n’ajoute pas les nouvelles lignes). Recalc recalcule les contrôles calculés.',
      'Refresh liest bereits geladene Datensätze erneut (fügt keine neuen Zeilen hinzu). Recalc berechnet berechnete Steuerelemente neu.',
    ),
    examples: [
      {
        title: L('Refresh vs Recalc su frmOrdini', 'Refresh vs Recalc on frmOrdini', 'Refresh frente a Recalc en frmOrdini', 'Refresh ou Recalc sur frmOrdini', 'Refresh vs. Recalc auf frmOrdini'),
        code: `' Modulo di frmOrdini. txtTotale ha ControlSource = =DSum("Importo","Ordini","IDCliente=" & [IDCliente])
Private Sub cmdRileggi_Click()
    Me.Refresh
    Me.Recalc
    Debug.Print Me!txtTotale
    Debug.Print Me.Dirty
End Sub`,
        result: L(
          'I campi bound mostrano i valori salvati da altri utenti. txtTotale si aggiorna (es. 18450). Dirty resta False se non hai modifiche locali. Le righe nuove non compaiono: serve Me.Requery.',
          'Bound fields show values saved by others. txtTotale updates (e.g. 18450). Dirty stays False if you have no local edits. New rows do not appear: use Me.Requery.',
          'Los campos enlazados muestran valores guardados por otros. txtTotale se actualiza (p. ej. 18450). Dirty sigue False si no hay cambios locales. Las filas nuevas no aparecen: usa Me.Requery.',
          'Les champs liés montrent les valeurs enregistrées par d’autres. txtTotale se met à jour (p. ex. 18450). Dirty reste False sans modifications locales. Les nouvelles lignes n’apparaissent pas : utilisez Me.Requery.',
          'Gebundene Felder zeigen von anderen gespeicherte Werte. txtTotale aktualisiert sich (z. B. 18450). Dirty bleibt False ohne lokale Änderungen. Neue Zeilen erscheinen nicht: Me.Requery verwenden.',
        ),
      },
    ],
    related: ['ac-me-requery', 'ac-me-undo-dirty', 'ac-dsum'],
  },
  {
    id: 'ac-me-undo-dirty',
    name: 'Me.Undo / Me.Dirty',
    syntax: 'Me.Undo | Me.Dirty As Boolean',
    scope: ['access'],
    category: 'ac-ui',
    subcategory: 'forms',
    description: L(
      'Dirty è True se il record corrente ha modifiche non salvate. Undo le scarta (come Esc).',
      'Dirty is True if the current record has unsaved edits. Undo discards them (like Esc).',
      'Dirty es True si el registro actual tiene cambios no guardados. Undo los descarta (como Esc).',
      'Dirty est True si l’enregistrement courant a des modifications non enregistrées. Undo les annule (comme Échap).',
      'Dirty ist True, wenn der aktuelle Datensatz ungespeicherte Änderungen hat. Undo verwirft sie (wie Esc).',
    ),
    examples: [
      {
        title: L('Annullare se la maschera è sporca', 'Undo if the form is dirty', 'Deshacer si el formulario está sucio', 'Annuler si le formulaire est sale', 'Rückgängig wenn das Formular dirty ist'),
        code: `' Modulo di frmClienti
Private Sub cmdAnnulla_Click()
    Debug.Print Me.Dirty
    If Me.Dirty Then
        Me.Undo
    End If
    Debug.Print Me.Dirty
    Debug.Print Me!Citta
End Sub`,
        result: L(
          'Se l’utente aveva digitato "Genova" al posto di Milano: prima Dirty=True, dopo Undo Dirty=False e Citta=Milano.',
          'If the user typed "Genova" instead of Milan: first Dirty=True, after Undo Dirty=False and Citta=Milano.',
          'Si el usuario escribió "Genova" en lugar de Milán: primero Dirty=True, tras Undo Dirty=False y Citta=Milano.',
          'Si l’utilisateur a saisi "Genova" à la place de Milan : d’abord Dirty=True, après Undo Dirty=False et Citta=Milano.',
          'Hat der Benutzer "Genova" statt Mailand eingegeben: zuerst Dirty=True, nach Undo Dirty=False und Citta=Milano.',
        ),
      },
    ],
    notes: L(
      'Puoi assegnare Me.Dirty = False in BeforeUpdate per forzare il salvataggio senza uscire dal record, ma Undo è il modo sicuro per abbandonare le modifiche.',
      'You can set Me.Dirty = False in BeforeUpdate to force a save without leaving the record, but Undo is the safe way to abandon edits.',
      'Puedes asignar Me.Dirty = False en BeforeUpdate para forzar el guardado sin salir del registro, pero Undo es la forma segura de abandonar los cambios.',
      'Vous pouvez affecter Me.Dirty = False dans BeforeUpdate pour forcer l’enregistrement sans quitter, mais Undo est la façon sûre d’abandonner les modifications.',
      'Me.Dirty = False in BeforeUpdate erzwingt Speichern ohne Datensatzwechsel; Undo ist der sichere Weg, Änderungen zu verwerfen.',
    ),
    related: ['ac-me', 'ac-dao-update', 'ac-docmd-close'],
  },
  {
    id: 'ac-screen-active',
    name: 'Screen.ActiveForm / ActiveControl',
    syntax: 'Screen.ActiveForm As Form | Screen.ActiveControl As Control',
    scope: ['access'],
    category: 'ac-ui',
    subcategory: 'forms',
    description: L(
      'Maschera e controllo che hanno lo stato attivo. Utili da un modulo standard.',
      'Form and control that have the focus. Useful from a standard module.',
      'Formulario y control que tienen el foco. Útiles desde un módulo estándar.',
      'Formulaire et contrôle qui ont le focus. Utiles depuis un module standard.',
      'Formular und Steuerelement mit dem Fokus. Nützlich aus einem Standardmodul.',
    ),
    examples: [
      {
        title: L('Nome della maschera e del controllo attivi', 'Active form and control names', 'Nombres del formulario y control activos', 'Noms du formulaire et du contrôle actifs', 'Namen von aktivem Formular und Steuerelement'),
        code: `Sub EsempioScreen()
    Debug.Print Screen.ActiveForm.Name
    Debug.Print Screen.ActiveControl.Name
    Debug.Print Screen.ActiveControl.Value
End Sub`,
        result: L(
          'Con il cursore su RagioneSociale in frmClienti: frmClienti  /  RagioneSociale  /  Bianchi Srl. Senza maschera aperta: errore 2474/2475.',
          'With the caret on RagioneSociale in frmClienti: frmClienti  /  RagioneSociale  /  Bianchi Srl. With no form open: error 2474/2475.',
          'Con el cursor en RagioneSociale de frmClienti: frmClienti  /  RagioneSociale  /  Bianchi Srl. Sin formulario abierto: error 2474/2475.',
          'Avec le curseur sur RagioneSociale dans frmClienti : frmClienti  /  RagioneSociale  /  Bianchi Srl. Sans formulaire ouvert : erreur 2474/2475.',
          'Mit Cursor auf RagioneSociale in frmClienti: frmClienti  /  RagioneSociale  /  Bianchi Srl. Ohne offenes Formular: Fehler 2474/2475.',
        ),
      },
    ],
    related: ['ac-me', 'ac-forms-bang', 'ac-code-context-object'],
  },
  {
    id: 'ac-open-args',
    name: 'OpenArgs',
    syntax: 'DoCmd.OpenForm FormName, OpenArgs:=args  →  Me.OpenArgs',
    scope: ['access'],
    category: 'ac-ui',
    subcategory: 'forms',
    description: L(
      'Stringa passata ad OpenForm/OpenReport e letta nella maschera o nel report aperto.',
      'String passed to OpenForm/OpenReport and read in the opened form or report.',
      'Cadena pasada a OpenForm/OpenReport y leída en el formulario o informe abierto.',
      'Chaîne passée à OpenForm/OpenReport et lue dans le formulaire ou l’état ouvert.',
      'Zeichenfolge, die an OpenForm/OpenReport übergeben und im geöffneten Formular oder Bericht gelesen wird.',
    ),
    examples: [
      {
        title: L('Aprire in sola lettura se OpenArgs = readonly', 'Open read-only if OpenArgs = readonly', 'Abrir en solo lectura si OpenArgs = readonly', 'Ouvrir en lecture seule si OpenArgs = readonly', 'Schreibgeschützt öffnen wenn OpenArgs = readonly'),
        code: `Sub ApriSoloLettura()
    DoCmd.OpenForm "frmClienti", OpenArgs:="readonly|12"
End Sub

' Modulo di frmClienti
Private Sub Form_Load()
    Dim pezzi() As String
    pezzi = Split(Nz(Me.OpenArgs, ""), "|")
    Debug.Print Me.OpenArgs
    If pezzi(0) = "readonly" Then
        Me.AllowEdits = False
    End If
    If UBound(pezzi) >= 1 Then
        Me.Filter = "Codice=" & pezzi(1)
        Me.FilterOn = True
    End If
End Sub`,
        result: L(
          'frmClienti si apre sul cliente 12, AllowEdits=False. Immediata: readonly|12. Senza OpenArgs la maschera resta modificabile e senza filtro.',
          'frmClienti opens on customer 12 with AllowEdits=False. Immediate: readonly|12. Without OpenArgs the form stays editable and unfiltered.',
          'frmClienti se abre en el cliente 12 con AllowEdits=False. Inmediato: readonly|12. Sin OpenArgs el formulario sigue editable y sin filtro.',
          'frmClienti s’ouvre sur le client 12 avec AllowEdits=False. Exécution : readonly|12. Sans OpenArgs le formulaire reste modifiable et sans filtre.',
          'frmClienti öffnet Kunde 12 mit AllowEdits=False. Direktbereich: readonly|12. Ohne OpenArgs bleibt das Formular editierbar und ungefiltert.',
        ),
      },
    ],
    related: ['ac-docmd-open-form', 'ac-me', 'ac-docmd-apply-filter', 'ac-open-report-preview'],
  },
  {
    id: 'ac-output-to',
    name: 'DoCmd.OutputTo',
    syntax: 'DoCmd.OutputTo ObjectType, [ObjectName], [OutputFormat], [OutputFile], [AutoStart], [TemplateFile], [Encoding], [OutputQuality]',
    scope: ['access'],
    category: 'ac-ui',
    subcategory: 'reports',
    description: L(
      'Esporta una maschera, un report, una query o una tabella in PDF, Excel, RTF, testo o HTML.',
      'Exports a form, report, query or table to PDF, Excel, RTF, text or HTML.',
      'Exporta un formulario, informe, consulta o tabla a PDF, Excel, RTF, texto o HTML.',
      'Exporte un formulaire, état, requête ou table vers PDF, Excel, RTF, texte ou HTML.',
      'Exportiert ein Formular, einen Bericht, eine Abfrage oder Tabelle nach PDF, Excel, RTF, Text oder HTML.',
    ),
    params: [
      { name: 'ObjectType', description: L('acOutputReport, acOutputForm, acOutputQuery, acOutputTable.', 'acOutputReport, acOutputForm, acOutputQuery, acOutputTable.', 'acOutputReport, acOutputForm, acOutputQuery, acOutputTable.', 'acOutputReport, acOutputForm, acOutputQuery, acOutputTable.', 'acOutputReport, acOutputForm, acOutputQuery, acOutputTable.') },
      { name: 'ObjectName', optional: true, description: L('Nome oggetto.', 'Object name.', 'Nombre del objeto.', 'Nom de l’objet.', 'Objektname.') },
      { name: 'OutputFormat', optional: true, description: L('acFormatPDF, acFormatXLSX, acFormatRTF, acFormatTXT, acFormatHTML.', 'acFormatPDF, acFormatXLSX, acFormatRTF, acFormatTXT, acFormatHTML.', 'acFormatPDF, acFormatXLSX, acFormatRTF, acFormatTXT, acFormatHTML.', 'acFormatPDF, acFormatXLSX, acFormatRTF, acFormatTXT, acFormatHTML.', 'acFormatPDF, acFormatXLSX, acFormatRTF, acFormatTXT, acFormatHTML.') },
      { name: 'OutputFile', optional: true, description: L('Percorso di destinazione.', 'Destination path.', 'Ruta de destino.', 'Chemin de destination.', 'Zielpfad.') },
      { name: 'AutoStart', optional: true, description: L('True apre il file dopo l’export.', 'True opens the file after export.', 'True abre el archivo tras exportar.', 'True ouvre le fichier après l’export.', 'True öffnet die Datei nach dem Export.') },
    ],
    examples: [
      {
        title: L('PDF del report ordini', 'PDF of the orders report', 'PDF del informe de pedidos', 'PDF de l’état des commandes', 'PDF des Bestellberichts'),
        code: `Sub EsportaPdfOrdini()
    DoCmd.OutputTo acOutputReport, "rptOrdini", acFormatPDF, "C:\\Dati\\Ordini.pdf", False
    Debug.Print Dir("C:\\Dati\\Ordini.pdf")
End Sub`,
        result: L(
          'Crea C:\\Dati\\Ordini.pdf (report a pagina intera). Immediata: Ordini.pdf. AutoStart:=True lo aprirebbe in Edge/Acrobat.',
          'Creates C:\\Dati\\Ordini.pdf (full-page report). Immediate: Ordini.pdf. AutoStart:=True would open it in Edge/Acrobat.',
          'Crea C:\\Dati\\Ordini.pdf (informe a página completa). Inmediato: Ordini.pdf. AutoStart:=True lo abriría en Edge/Acrobat.',
          'Crée C:\\Dati\\Ordini.pdf (état pleine page). Exécution : Ordini.pdf. AutoStart:=True l’ouvrirait dans Edge/Acrobat.',
          'Erstellt C:\\Dati\\Ordini.pdf (ganzseitiger Bericht). Direktbereich: Ordini.pdf. AutoStart:=True würde ihn in Edge/Acrobat öffnen.',
        ),
      },
    ],
    related: ['ac-transfer-spreadsheet', 'ac-transfer-text', 'ac-open-report-preview', 'ac-docmd-open-report'],
  },
  {
    id: 'ac-transfer-spreadsheet',
    name: 'DoCmd.TransferSpreadsheet',
    syntax: 'DoCmd.TransferSpreadsheet [TransferType], [SpreadsheetType], TableName, FileName, [HasFieldNames], [Range], [UseOA]',
    scope: ['access'],
    category: 'ac-ui',
    subcategory: 'reports',
    description: L(
      'Importa o esporta una tabella/query da o verso un foglio Excel.',
      'Imports or exports a table/query from or to an Excel workbook.',
      'Importa o exporta una tabla/consulta desde o hacia un libro de Excel.',
      'Importe ou exporte une table/requête depuis ou vers un classeur Excel.',
      'Importiert oder exportiert eine Tabelle/Abfrage aus oder in eine Excel-Mappe.',
    ),
    params: [
      { name: 'TransferType', optional: true, description: L('acExport, acImport, acLink.', 'acExport, acImport, acLink.', 'acExport, acImport, acLink.', 'acExport, acImport, acLink.', 'acExport, acImport, acLink.') },
      { name: 'SpreadsheetType', optional: true, description: L('acSpreadsheetTypeExcel12Xml (.xlsx).', 'acSpreadsheetTypeExcel12Xml (.xlsx).', 'acSpreadsheetTypeExcel12Xml (.xlsx).', 'acSpreadsheetTypeExcel12Xml (.xlsx).', 'acSpreadsheetTypeExcel12Xml (.xlsx).') },
      { name: 'TableName', description: L('Tabella o query Access.', 'Access table or query.', 'Tabla o consulta de Access.', 'Table ou requête Access.', 'Access-Tabelle oder -Abfrage.') },
      { name: 'FileName', description: L('Percorso .xlsx/.xls.', 'Path to .xlsx/.xls.', 'Ruta .xlsx/.xls.', 'Chemin .xlsx/.xls.', 'Pfad zur .xlsx/.xls.') },
      { name: 'HasFieldNames', optional: true, description: L('True = prima riga intestazioni.', 'True = first row is headers.', 'True = la primera fila son encabezados.', 'True = première ligne = en-têtes.', 'True = erste Zeile ist Überschriften.') },
    ],
    examples: [
      {
        title: L('Esportare Clienti in Excel', 'Export Clienti to Excel', 'Exportar Clienti a Excel', 'Exporter Clienti vers Excel', 'Clienti nach Excel exportieren'),
        code: `Sub EsportaClientiXlsx()
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, _
        "Clienti", "C:\\Dati\\Clienti.xlsx", True
    Debug.Print Dir("C:\\Dati\\Clienti.xlsx")
End Sub`,
        result: L(
          'File C:\\Dati\\Clienti.xlsx con foglio Clienti: riga 1 = Codice, RagioneSociale, Citta…; 128 righe dati. Immediata: Clienti.xlsx.',
          'File C:\\Dati\\Clienti.xlsx with sheet Clienti: row 1 = Codice, RagioneSociale, Citta…; 128 data rows. Immediate: Clienti.xlsx.',
          'Archivo C:\\Dati\\Clienti.xlsx con hoja Clienti: fila 1 = Codice, RagioneSociale, Citta…; 128 filas. Inmediato: Clienti.xlsx.',
          'Fichier C:\\Dati\\Clienti.xlsx avec feuille Clienti : ligne 1 = Codice, RagioneSociale, Citta… ; 128 lignes. Exécution : Clienti.xlsx.',
          'Datei C:\\Dati\\Clienti.xlsx mit Blatt Clienti: Zeile 1 = Codice, RagioneSociale, Citta…; 128 Datenzeilen. Direktbereich: Clienti.xlsx.',
        ),
      },
    ],
    related: ['ac-transfer-text', 'ac-output-to', 'ac-docmd-open-table'],
  },
  {
    id: 'ac-transfer-text',
    name: 'DoCmd.TransferText',
    syntax: 'DoCmd.TransferText [TransferType], [SpecificationName], TableName, FileName, [HasFieldNames], [HTMLTableName], [CodePage]',
    scope: ['access'],
    category: 'ac-ui',
    subcategory: 'reports',
    description: L(
      'Importa o esporta testo delimitato, a larghezza fissa o HTML.',
      'Imports or exports delimited, fixed-width or HTML text.',
      'Importa o exporta texto delimitado, de ancho fijo o HTML.',
      'Importe ou exporte du texte délimité, à largeur fixe ou HTML.',
      'Importiert oder exportiert getrennten, festbreiten oder HTML-Text.',
    ),
    params: [
      { name: 'TransferType', optional: true, description: L('acExportDelim, acImportDelim, acExportFixed, acImportFixed, acExportHTML.', 'acExportDelim, acImportDelim, acExportFixed, acImportFixed, acExportHTML.', 'acExportDelim, acImportDelim, acExportFixed, acImportFixed, acExportHTML.', 'acExportDelim, acImportDelim, acExportFixed, acImportFixed, acExportHTML.', 'acExportDelim, acImportDelim, acExportFixed, acImportFixed, acExportHTML.') },
      { name: 'SpecificationName', optional: true, description: L('Specifica di importazione salvata (può essere Null).', 'Saved import spec (may be Null).', 'Especificación de importación guardada (puede ser Null).', 'Spécification d’import enregistrée (peut être Null).', 'Gespeicherte Importspezifikation (darf Null sein).') },
      { name: 'TableName', description: L('Tabella o query.', 'Table or query.', 'Tabla o consulta.', 'Table ou requête.', 'Tabelle oder Abfrage.') },
      { name: 'FileName', description: L('Percorso .csv/.txt/.html.', 'Path to .csv/.txt/.html.', 'Ruta .csv/.txt/.html.', 'Chemin .csv/.txt/.html.', 'Pfad zur .csv/.txt/.html.') },
    ],
    examples: [
      {
        title: L('CSV dei clienti', 'Customer CSV', 'CSV de clientes', 'CSV des clients', 'Kunden-CSV'),
        code: `Sub EsportaClientiCsv()
    DoCmd.TransferText acExportDelim, , "Clienti", "C:\\Dati\\Clienti.csv", True
    Debug.Print Dir("C:\\Dati\\Clienti.csv")
End Sub`,
        result: L(
          'C:\\Dati\\Clienti.csv inizia con Codice,RagioneSociale,Citta… e 128 righe. Delimitatore = quello regionale (spesso ; in Italia).',
          'C:\\Dati\\Clienti.csv starts with Codice,RagioneSociale,Citta… and 128 rows. Delimiter follows regional settings (often ; in Italy).',
          'C:\\Dati\\Clienti.csv empieza por Codice,RagioneSociale,Citta… y 128 filas. El delimitador sigue la configuración regional (a menudo ; en Italia).',
          'C:\\Dati\\Clienti.csv commence par Codice,RagioneSociale,Citta… et 128 lignes. Le délimiteur suit les paramètres régionaux (souvent ; en Italie).',
          'C:\\Dati\\Clienti.csv beginnt mit Codice,RagioneSociale,Citta… und 128 Zeilen. Das Trennzeichen folgt den Regionals (in Italien oft ;).',
        ),
      },
    ],
    related: ['ac-transfer-spreadsheet', 'ac-output-to'],
  },
  {
    id: 'ac-reports-bang',
    name: 'Reports!Name',
    syntax: 'Reports!ReportName!ControlName',
    scope: ['access'],
    category: 'ac-ui',
    subcategory: 'reports',
    description: L(
      'Collezione dei report aperti: leggi o imposta controlli mentre il report è in anteprima o in Format.',
      'Collection of open reports: read or set controls while the report is in preview or Format.',
      'Colección de informes abiertos: lee o asigna controles mientras el informe está en vista previa o Format.',
      'Collection des états ouverts : lisez ou affectez des contrôles pendant l’aperçu ou Format.',
      'Auflistung geöffneter Berichte: Steuerelemente lesen oder setzen, während der Bericht in der Vorschau oder in Format ist.',
    ),
    examples: [
      {
        title: L('Titolo dinamico in anteprima', 'Dynamic title in preview', 'Título dinámico en vista previa', 'Titre dynamique en aperçu', 'Dynamischer Titel in der Vorschau'),
        code: `Sub EsempioReportsBang()
    DoCmd.OpenReport "rptOrdini", acViewPreview
    Reports!rptOrdini!lblTitolo.Caption = "Ordini 2026"
    Debug.Print Reports!rptOrdini.HasData
    Debug.Print Reports!rptOrdini!lblTitolo.Caption
End Sub`,
        result: L(
          'Anteprima di rptOrdini con etichetta "Ordini 2026". Immediata: True  /  Ordini 2026. Se il report è chiuso: errore 2451.',
          'Preview of rptOrdini with caption "Ordini 2026". Immediate: True  /  Ordini 2026. If the report is closed: error 2451.',
          'Vista previa de rptOrdini con título "Ordini 2026". Inmediato: True  /  Ordini 2026. Si el informe está cerrado: error 2451.',
          'Aperçu de rptOrdini avec libellé "Ordini 2026". Exécution : True  /  Ordini 2026. Si l’état est fermé : erreur 2451.',
          'Vorschau von rptOrdini mit Beschriftung "Ordini 2026". Direktbereich: True  /  Ordini 2026. Ist der Bericht geschlossen: Fehler 2451.',
        ),
      },
    ],
    related: ['ac-forms-bang', 'ac-open-report-preview', 'ac-docmd-open-report'],
  },
  {
    id: 'ac-open-report-preview',
    name: 'DoCmd.OpenReport acViewPreview',
    syntax: 'DoCmd.OpenReport ReportName, acViewPreview, [FilterName], [WhereCondition], [WindowMode], [OpenArgs]',
    scope: ['access'],
    category: 'ac-ui',
    subcategory: 'reports',
    description: L(
      'Apre il report in anteprima di stampa a video, senza inviarlo alla stampante.',
      'Opens the report in on-screen print preview without sending it to the printer.',
      'Abre el informe en vista previa de impresión en pantalla, sin enviarlo a la impresora.',
      'Ouvre l’état en aperçu avant impression à l’écran, sans l’envoyer à l’imprimante.',
      'Öffnet den Bericht in der Bildschirm-Druckvorschau, ohne ihn an den Drucker zu senden.',
    ),
    params: [
      { name: 'ReportName', description: L('Nome del report.', 'Report name.', 'Nombre del informe.', 'Nom de l’état.', 'Berichtsname.') },
      { name: 'WhereCondition', optional: true, description: L('Filtro record, es. "IDCliente=12".', 'Record filter, e.g. "IDCliente=12".', 'Filtro de registros, p. ej. "IDCliente=12".', 'Filtre d’enregistrements, p. ex. "IDCliente=12".', 'Datensatzfilter, z. B. "IDCliente=12".') },
      { name: 'OpenArgs', optional: true, description: L('Stringa letta da Report.OpenArgs.', 'String read from Report.OpenArgs.', 'Cadena leída en Report.OpenArgs.', 'Chaîne lue via Report.OpenArgs.', 'Zeichenfolge in Report.OpenArgs.') },
    ],
    examples: [
      {
        title: L('Anteprima filtrata con OpenArgs', 'Filtered preview with OpenArgs', 'Vista previa filtrada con OpenArgs', 'Aperçu filtré avec OpenArgs', 'Gefilterte Vorschau mit OpenArgs'),
        code: `Sub AnteprimaCliente()
    DoCmd.OpenReport "rptOrdini", acViewPreview, , "IDCliente=12", acWindowNormal, "cliente=12"
End Sub

' Modulo di rptOrdini
Private Sub Report_Open(Cancel As Integer)
    Debug.Print Me.OpenArgs
    Me!lblSottotitolo.Caption = "Cliente " & Mid(Nz(Me.OpenArgs, ""), 9)
End Sub`,
        result: L(
          'rptOrdini in anteprima: solo gli ordini del cliente 12 (es. 8 pagine). Immediata: cliente=12. lblSottotitolo = "Cliente 12". Nessun job di stampa.',
          'rptOrdini in preview: only customer 12’s orders (e.g. 8 pages). Immediate: cliente=12. lblSottotitolo = "Cliente 12". No print job.',
          'rptOrdini en vista previa: solo los pedidos del cliente 12 (p. ej. 8 páginas). Inmediato: cliente=12. lblSottotitolo = "Cliente 12". Sin trabajo de impresión.',
          'rptOrdini en aperçu : uniquement les commandes du client 12 (p. ex. 8 pages). Exécution : cliente=12. lblSottotitolo = "Cliente 12". Aucun travail d’impression.',
          'rptOrdini in der Vorschau: nur Bestellungen von Kunde 12 (z. B. 8 Seiten). Direktbereich: cliente=12. lblSottotitolo = "Cliente 12". Kein Druckauftrag.',
        ),
      },
    ],
    related: ['ac-docmd-open-report', 'ac-reports-bang', 'ac-output-to', 'ac-open-args'],
  },
]
