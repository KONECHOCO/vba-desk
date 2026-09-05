import { L, type Command } from '../types'

export const coreCommands: Command[] = [
  {
    id: 'dim',
    name: 'Dim',
    syntax: 'Dim nome [As tipo][, nome2 [As tipo2]] ...',
    scope: ['core', 'excel', 'access'],
    category: 'declarations',
    subcategory: 'variables',
    description: L(
      'Dichiara una o più variabili nella procedura o nel modulo. Senza As il tipo è Variant.',
      'Declares one or more variables in the procedure or module. Without As the type is Variant.',
      'Declara una o más variables en el procedimiento o módulo. Sin As el tipo es Variant.',
      'Déclare une ou plusieurs variables dans la procédure ou le module. Sans As le type est Variant.',
      'Deklariert eine oder mehrere Variablen in der Prozedur oder im Modul. Ohne As ist der Typ Variant.',
    ),
    params: [
      { name: 'nome', description: L('Identificatore della variabile.', 'Variable identifier.', 'Identificador de la variable.', 'Identifiant de la variable.', 'Bezeichner der Variable.') },
      { name: 'tipo', optional: true, description: L('Tipo As (Long, String, …). Default: Variant.', 'As type (Long, String, …). Default: Variant.', 'Tipo As (Long, String, …). Predeterminado: Variant.', 'Type As (Long, String, …). Par défaut : Variant.', 'As-Typ (Long, String, …). Standard: Variant.') },
    ],
    examples: [
      {
        title: L('Variabile Long', 'Long variable', 'Variable Long', 'Variable Long', 'Long-Variable'),
        code: `Sub DemoDim()
    Dim n As Long
    n = 42
    Debug.Print n
End Sub`,
        result: L('Finestra Immediata: 42', 'Immediate Window: 42', 'Ventana Inmediato: 42', 'Fenêtre Exécution : 42', 'Direktbereich: 42'),
      },
    ],
    notes: L(
      'Con Option Explicit ogni variabile va dichiarata. Dim a livello di modulo è privato del modulo.',
      'With Option Explicit every variable must be declared. Module-level Dim is private to the module.',
      'Con Option Explicit hay que declarar cada variable. Dim a nivel de módulo es privado del módulo.',
      'Avec Option Explicit chaque variable doit être déclarée. Dim au niveau module est privé au module.',
      'Mit Option Explicit muss jede Variable deklariert werden. Dim auf Modulebene ist modulprivat.',
    ),
    related: ['option-explicit', 'public', 'private', 'static', 'variant'],
  },
  {
    id: 'redim-preserve',
    name: 'ReDim Preserve',
    syntax: 'ReDim [Preserve] nome(indiceSuperiore) [As tipo]',
    scope: ['core', 'excel', 'access'],
    category: 'declarations',
    subcategory: 'variables',
    description: L(
      'Ridimensiona un array dinamico. Preserve mantiene i valori esistenti (solo l’ultima dimensione).',
      'Resizes a dynamic array. Preserve keeps existing values (last dimension only).',
      'Redimensiona una matriz dinámica. Preserve conserva los valores (solo la última dimensión).',
      'Redimensionne un tableau dynamique. Preserve conserve les valeurs (dernière dimension uniquement).',
      'Ändert die Größe eines dynamischen Arrays. Preserve behält vorhandene Werte (nur letzte Dimension).',
    ),
    params: [
      { name: 'Preserve', optional: true, description: L('Mantiene i dati già presenti.', 'Keeps data already stored.', 'Conserva los datos ya guardados.', 'Conserve les données déjà stockées.', 'Behält bereits gespeicherte Daten.') },
      { name: 'indiceSuperiore', description: L('Nuovo limite superiore (UBound).', 'New upper bound (UBound).', 'Nuevo límite superior (UBound).', 'Nouvelle borne supérieure (UBound).', 'Neue Obergrenze (UBound).') },
    ],
    examples: [
      {
        title: L('Allarga un array senza perdere i valori', 'Grow an array without losing values', 'Ampliar una matriz sin perder valores', 'Agrandir un tableau sans perdre les valeurs', 'Array vergrößern ohne Werteverlust'),
        code: `Sub DemoReDimPreserve()
    Dim a() As String
    ReDim a(1)
    a(0) = "uno": a(1) = "due"
    ReDim Preserve a(2)
    a(2) = "tre"
    Debug.Print a(0), a(1), a(2)
End Sub`,
        result: L('Finestra Immediata: uno    due    tre', 'Immediate Window: uno    due    tre', 'Ventana Inmediato: uno    due    tre', 'Fenêtre Exécution : uno    due    tre', 'Direktbereich: uno    due    tre'),
      },
    ],
    notes: L(
      'Senza Preserve i valori vengono azzerati. ReDim Preserve non può cambiare il tipo né le dimensioni inferiori.',
      'Without Preserve values are cleared. ReDim Preserve cannot change the type or lower dimensions.',
      'Sin Preserve los valores se borran. ReDim Preserve no puede cambiar el tipo ni las dimensiones inferiores.',
      'Sans Preserve les valeurs sont effacées. ReDim Preserve ne peut pas changer le type ni les dimensions inférieures.',
      'Ohne Preserve werden Werte gelöscht. ReDim Preserve kann Typ und untere Dimensionen nicht ändern.',
    ),
    related: ['dim', 'array-bounds', 'erase'],
  },
  {
    id: 'public',
    name: 'Public',
    syntax: 'Public nome [As tipo]',
    scope: ['core', 'excel', 'access'],
    category: 'declarations',
    subcategory: 'variables',
    description: L(
      'Dichiara una variabile, costante o procedura visibile da tutti i moduli del progetto.',
      'Declares a variable, constant or procedure visible from every module in the project.',
      'Declara una variable, constante o procedimiento visible desde todos los módulos del proyecto.',
      'Déclare une variable, constante ou procédure visible depuis tous les modules du projet.',
      'Deklariert eine Variable, Konstante oder Prozedur, die in allen Modulen des Projekts sichtbar ist.',
    ),
    examples: [
      {
        title: L('Contatore condiviso tra moduli', 'Counter shared across modules', 'Contador compartido entre módulos', 'Compteur partagé entre modules', 'Zähler über Module hinweg'),
        code: `' In un modulo standard:
Public gCount As Long

Sub DemoPublic()
    gCount = gCount + 1
    Debug.Print gCount
End Sub`,
        result: L('Prima esecuzione: 1; seconda: 2 (resta in memoria).', 'First run: 1; second: 2 (stays in memory).', 'Primera ejecución: 1; segunda: 2 (sigue en memoria).', '1re exécution : 1 ; 2e : 2 (reste en mémoire).', 'Erster Lauf: 1; zweiter: 2 (bleibt im Speicher).'),
      },
    ],
    notes: L(
      'Public a livello di modulo in un foglio o form è visibile; Option Private Module la limita al progetto VBA.',
      'Public at module level in a sheet or form is visible; Option Private Module limits it to the VBA project.',
      'Public a nivel de módulo en una hoja o formulario es visible; Option Private Module lo limita al proyecto VBA.',
      'Public au niveau module dans une feuille ou un form est visible ; Option Private Module le limite au projet VBA.',
      'Public auf Modulebene in Blatt oder Formular ist sichtbar; Option Private Module begrenzt es auf das VBA-Projekt.',
    ),
    related: ['private', 'dim', 'option-private-module'],
  },
  {
    id: 'private',
    name: 'Private',
    syntax: 'Private nome [As tipo]',
    scope: ['core', 'excel', 'access'],
    category: 'declarations',
    subcategory: 'variables',
    description: L(
      'Limita variabile o procedura al modulo corrente: gli altri moduli non la vedono.',
      'Limits a variable or procedure to the current module: other modules cannot see it.',
      'Limita una variable o procedimiento al módulo actual: los demás no lo ven.',
      'Limite une variable ou procédure au module courant : les autres ne la voient pas.',
      'Beschränkt Variable oder Prozedur auf das aktuelle Modul: andere Module sehen sie nicht.',
    ),
    examples: [
      {
        title: L('Helper nascosto agli altri moduli', 'Helper hidden from other modules', 'Helper oculto a otros módulos', 'Helper caché aux autres modules', 'Hilfsroutine für andere Module unsichtbar'),
        code: `Private Function Doppio(ByVal n As Long) As Long
    Doppio = n * 2
End Function

Sub DemoPrivate()
    Debug.Print Doppio(7)
End Sub`,
        result: L('Finestra Immediata: 14. Da un altro modulo Doppio non è chiamabile.', 'Immediate Window: 14. Another module cannot call Doppio.', 'Ventana Inmediato: 14. Otro módulo no puede llamar a Doppio.', 'Fenêtre Exécution : 14. Un autre module ne peut pas appeler Doppio.', 'Direktbereich: 14. Ein anderes Modul kann Doppio nicht aufrufen.'),
      },
    ],
    related: ['public', 'dim', 'option-private-module'],
  },
  {
    id: 'static',
    name: 'Static',
    syntax: 'Static nome [As tipo]',
    scope: ['core', 'excel', 'access'],
    category: 'declarations',
    subcategory: 'variables',
    description: L(
      'Variabile di procedura che conserva il valore tra una chiamata e la successiva.',
      'Procedure-level variable that keeps its value between calls.',
      'Variable de procedimiento que conserva el valor entre llamadas.',
      'Variable de procédure qui conserve sa valeur entre les appels.',
      'Prozedurvariable, die den Wert zwischen Aufrufen behält.',
    ),
    examples: [
      {
        title: L('Quante volte è stata eseguita', 'How many times it ran', 'Cuántas veces se ejecutó', 'Combien de fois elle a tourné', 'Wie oft sie lief'),
        code: `Sub DemoStatic()
    Static hits As Long
    hits = hits + 1
    Debug.Print hits
End Sub`,
        result: L('1ª chiamata: 1; 2ª: 2; 3ª: 3. Resetta con Fine / Reset.', '1st call: 1; 2nd: 2; 3rd: 3. Reset with End / Reset.', '1.ª llamada: 1; 2.ª: 2; 3.ª: 3. Se reinicia con End / Restablecer.', '1er appel : 1 ; 2e : 2 ; 3e : 3. Réinitialiser avec End / Reset.', '1. Aufruf: 1; 2.: 2; 3.: 3. Zurücksetzen mit End / Zurücksetzen.'),
      },
    ],
    notes: L(
      'Static Sub rende statiche tutte le variabili locali della procedura.',
      'Static Sub makes every local variable in the procedure static.',
      'Static Sub hace estáticas todas las variables locales del procedimiento.',
      'Static Sub rend statiques toutes les variables locales de la procédure.',
      'Static Sub macht alle lokalen Variablen der Prozedur statisch.',
    ),
    related: ['dim', 'private', 'end-stop'],
  },
  {
    id: 'const',
    name: 'Const',
    syntax: '[Public | Private] Const nome [As tipo] = espressione',
    scope: ['core', 'excel', 'access'],
    category: 'declarations',
    subcategory: 'variables',
    description: L(
      'Costante valutata in compilazione: il valore non può cambiare a runtime.',
      'Compile-time constant: the value cannot change at runtime.',
      'Constante evaluada en compilación: el valor no puede cambiar en ejecución.',
      'Constante évaluée à la compilation : la valeur ne peut pas changer à l’exécution.',
      'Kompilierzeitkonstante: der Wert kann zur Laufzeit nicht geändert werden.',
    ),
    params: [
      { name: 'espressione', description: L('Letterale o combinazione di altre costanti.', 'Literal or combination of other constants.', 'Literal o combinación de otras constantes.', 'Littéral ou combinaison d’autres constantes.', 'Literal oder Kombination anderer Konstanten.') },
    ],
    examples: [
      {
        title: L('IVA fissa', 'Fixed VAT rate', 'IVA fija', 'TVA fixe', 'Fester MwSt-Satz'),
        code: `Sub DemoConst()
    Const IVA As Double = 0.22
    Debug.Print 100 * (1 + IVA)
End Sub`,
        result: L('Finestra Immediata: 122', 'Immediate Window: 122', 'Ventana Inmediato: 122', 'Fenêtre Exécution : 122', 'Direktbereich: 122'),
      },
    ],
    notes: L(
      'Non si può assegnare una costante a una chiamata di funzione (Now, ecc.).',
      'You cannot assign a constant from a function call (Now, etc.).',
      'No se puede asignar una constante desde una llamada a función (Now, etc.).',
      'On ne peut pas affecter une constante depuis un appel de fonction (Now, etc.).',
      'Eine Konstante kann nicht aus einem Funktionsaufruf (Now usw.) zugewiesen werden.',
    ),
    related: ['dim', 'public', 'enum'],
  },
  {
    id: 'set-new-nothing',
    name: 'Set, New, Nothing',
    syntax: 'Set oggetto = New classe | Set oggetto = Nothing',
    scope: ['core', 'excel', 'access'],
    category: 'declarations',
    subcategory: 'variables',
    description: L(
      'Set assegna un riferimento a un oggetto. New crea l’istanza. Nothing rilascia il riferimento.',
      'Set assigns an object reference. New creates the instance. Nothing releases the reference.',
      'Set asigna una referencia a un objeto. New crea la instancia. Nothing libera la referencia.',
      'Set affecte une référence d’objet. New crée l’instance. Nothing libère la référence.',
      'Set weist eine Objektreferenz zu. New erzeugt die Instanz. Nothing gibt die Referenz frei.',
    ),
    params: [
      { name: 'oggetto', description: L('Variabile Object / classe concreta.', 'Object variable / concrete class.', 'Variable Object / clase concreta.', 'Variable Object / classe concrète.', 'Object-Variable / konkrete Klasse.') },
      { name: 'classe', description: L('Nome della classe da istanziare.', 'Class name to instantiate.', 'Nombre de la clase a instanciar.', 'Nom de la classe à instancier.', 'Klassenname zum Instanziieren.') },
    ],
    examples: [
      {
        title: L('Collection creata e poi rilasciata', 'Collection created then released', 'Collection creada y luego liberada', 'Collection créée puis libérée', 'Collection erzeugt und dann freigegeben'),
        code: `Sub DemoSetNewNothing()
    Dim c As Collection
    Set c = New Collection
    c.Add "alfa"
    Debug.Print c.Count
    Set c = Nothing
    Debug.Print c Is Nothing
End Sub`,
        result: L('Finestra Immediata: 1 poi True', 'Immediate Window: 1 then True', 'Ventana Inmediato: 1 luego True', 'Fenêtre Exécution : 1 puis True', 'Direktbereich: 1 dann True'),
      },
    ],
    notes: L(
      'Senza Set, VBA usa Let e fallisce sugli oggetti. New Collection è equivalente a CreateObject per alcune classi.',
      'Without Set, VBA uses Let and fails on objects. New Collection is equivalent to CreateObject for some classes.',
      'Sin Set, VBA usa Let y falla con objetos. New Collection equivale a CreateObject para algunas clases.',
      'Sans Set, VBA utilise Let et échoue sur les objets. New Collection équivaut à CreateObject pour certaines classes.',
      'Ohne Set verwendet VBA Let und scheitert bei Objekten. New Collection entspricht CreateObject für manche Klassen.',
    ),
    related: ['object', 'collection', 'createobject-getobject', 'operators-logic-compare'],
  },
  {
    id: 'user-type',
    name: 'Type ... End Type',
    syntax: '[Private | Public] Type NomeCampo\n    campo As tipo\nEnd Type',
    scope: ['core', 'excel', 'access'],
    category: 'declarations',
    subcategory: 'types-user',
    description: L(
      'Definisce una struttura (UDT) con campi nominati, dichiarata a livello di modulo.',
      'Defines a user-defined type (UDT) with named fields, declared at module level.',
      'Define una estructura (UDT) con campos con nombre, declarada a nivel de módulo.',
      'Définit une structure (UDT) avec des champs nommés, déclarée au niveau module.',
      'Definiert eine Struktur (UDT) mit benannten Feldern, auf Modulebene deklariert.',
    ),
    examples: [
      {
        title: L('Punto 2D', '2D point', 'Punto 2D', 'Point 2D', '2D-Punkt'),
        code: `Private Type Punto
    x As Long
    y As Long
End Type

Sub DemoType()
    Dim p As Punto
    p.x = 3: p.y = 4
    Debug.Print p.x, p.y
End Sub`,
        result: L('Finestra Immediata: 3    4', 'Immediate Window: 3    4', 'Ventana Inmediato: 3    4', 'Fenêtre Exécution : 3    4', 'Direktbereich: 3    4'),
      },
    ],
    notes: L(
      'Non si può dichiarare un Type dentro una Sub. Per oggetti usa Class Module, non Type.',
      'You cannot declare a Type inside a Sub. For objects use a Class Module, not Type.',
      'No se puede declarar un Type dentro de un Sub. Para objetos usa Class Module, no Type.',
      'On ne peut pas déclarer un Type dans un Sub. Pour les objets, utilisez un Class Module, pas Type.',
      'Ein Type kann nicht in einer Sub stehen. Für Objekte Class Module verwenden, nicht Type.',
    ),
    related: ['enum', 'dim'],
  },
  {
    id: 'enum',
    name: 'Enum',
    syntax: '[Public | Private] Enum Nome\n    Membro [= valore]\nEnd Enum',
    scope: ['core', 'excel', 'access'],
    category: 'declarations',
    subcategory: 'types-user',
    description: L(
      'Elenco di costanti Long con nome. I valori partono da 0 e incrementano di 1 se omessi.',
      'Named list of Long constants. Values start at 0 and increase by 1 when omitted.',
      'Lista de constantes Long con nombre. Los valores empiezan en 0 y aumentan de 1 si se omiten.',
      'Liste nommée de constantes Long. Les valeurs partent de 0 et s’incrémentent de 1 si omises.',
      'Benannte Liste von Long-Konstanten. Werte beginnen bei 0 und steigen um 1, wenn weggelassen.',
    ),
    examples: [
      {
        title: L('Stati di un ordine', 'Order statuses', 'Estados de un pedido', 'États d’une commande', 'Auftragsstatus'),
        code: `Private Enum StatoOrdine
    bozza = 0
    inviato = 1
    evaso = 2
End Enum

Sub DemoEnum()
    Dim s As StatoOrdine
    s = inviato
    Debug.Print s
End Sub`,
        result: L('Finestra Immediata: 1', 'Immediate Window: 1', 'Ventana Inmediato: 1', 'Fenêtre Exécution : 1', 'Direktbereich: 1'),
      },
    ],
    related: ['const', 'user-type', 'long'],
  },
  {
    id: 'option-explicit',
    name: 'Option Explicit',
    syntax: 'Option Explicit',
    scope: ['core', 'excel', 'access'],
    category: 'declarations',
    subcategory: 'options',
    description: L(
      'Obbliga a dichiarare ogni variabile. Va in cima al modulo, prima di qualsiasi procedura.',
      'Requires every variable to be declared. Must sit at the top of the module, before any procedure.',
      'Obliga a declarar cada variable. Debe ir al inicio del módulo, antes de cualquier procedimiento.',
      'Oblige à déclarer chaque variable. Doit figurer en tête du module, avant toute procédure.',
      'Erzwingt die Deklaration jeder Variable. Muss ganz oben im Modul stehen, vor jeder Prozedur.',
    ),
    examples: [
      {
        title: L('Refuso intercettato in compilazione', 'Typo caught at compile time', 'Errata detectada al compilar', 'Frappe détectée à la compilation', 'Tippfehler zur Kompilierzeit erkannt'),
        code: `Option Explicit

Sub DemoOptionExplicit()
    Dim totale As Long
    totale = 10
    ' totalee = 11  ' non compila: variabile non dichiarata
    Debug.Print totale
End Sub`,
        result: L('Finestra Immediata: 10. La riga commentata darebbe errore di compilazione.', 'Immediate Window: 10. The commented line would fail to compile.', 'Ventana Inmediato: 10. La línea comentada no compilaría.', 'Fenêtre Exécution : 10. La ligne commentée ne compilerait pas.', 'Direktbereich: 10. Die auskommentierte Zeile würde nicht kompilieren.'),
      },
    ],
    notes: L(
      'Strumenti → Opzioni → Dichiarazione di variabili obbligatoria la inserisce nei moduli nuovi.',
      'Tools → Options → Require Variable Declaration inserts it into new modules.',
      'Herramientas → Opciones → Declaración de variables obligatoria la inserta en módulos nuevos.',
      'Outils → Options → Déclaration des variables obligatoire l’insère dans les nouveaux modules.',
      'Extras → Optionen → Variablendeklaration erforderlich fügt sie in neue Module ein.',
    ),
    related: ['dim', 'variant'],
  },
  {
    id: 'option-compare',
    name: 'Option Compare',
    syntax: 'Option Compare {Binary | Text | Database}',
    scope: ['core', 'excel', 'access'],
    category: 'declarations',
    subcategory: 'options',
    description: L(
      'Stabilisce come =, <>, Like e StrComp confrontano le stringhe nel modulo (maiuscole e locale).',
      'Sets how =, <>, Like and StrComp compare strings in the module (case and locale).',
      'Define cómo =, <>, Like y StrComp comparan cadenas en el módulo (mayúsculas y configuración regional).',
      'Définit comment =, <>, Like et StrComp comparent les chaînes dans le module (casse et locale).',
      'Legt fest, wie =, <>, Like und StrComp Zeichenfolgen im Modul vergleichen (Groß/Klein und Locale).',
    ),
    examples: [
      {
        title: L('Binary distingue A e a', 'Binary treats A and a as different', 'Binary trata A y a como distintas', 'Binary traite A et a comme différents', 'Binary behandelt A und a verschieden'),
        code: `Option Compare Binary

Sub DemoOptionCompare()
    Debug.Print "A" = "a"
    Debug.Print "file" Like "FILE"
End Sub`,
        result: L('False poi False. Con Option Compare Text entrambi True.', 'False then False. With Option Compare Text both True.', 'False luego False. Con Option Compare Text ambos True.', 'False puis False. Avec Option Compare Text les deux True.', 'False dann False. Mit Option Compare Text beide True.'),
      },
    ],
    notes: L(
      'Database è tipico di Access e usa l’ordinamento del database. Default: Binary.',
      'Database is typical in Access and uses the database sort order. Default: Binary.',
      'Database es típico de Access y usa el orden de la base. Predeterminado: Binary.',
      'Database est typique d’Access et utilise l’ordre de la base. Par défaut : Binary.',
      'Database ist typisch für Access und nutzt die Datenbanksortierung. Standard: Binary.',
    ),
    related: ['operators-logic-compare', 'strcomp'],
  },
  {
    id: 'option-base',
    name: 'Option Base',
    syntax: 'Option Base {0 | 1}',
    scope: ['core', 'excel', 'access'],
    category: 'declarations',
    subcategory: 'options',
    description: L(
      'Imposta il limite inferiore predefinito degli array dichiarati con Dim nel modulo (0 o 1).',
      'Sets the default lower bound of arrays declared with Dim in the module (0 or 1).',
      'Establece el límite inferior predeterminado de las matrices declaradas con Dim en el módulo (0 o 1).',
      'Définit la borne inférieure par défaut des tableaux déclarés avec Dim dans le module (0 ou 1).',
      'Setzt die Standard-Untergrenze von mit Dim deklarierten Arrays im Modul (0 oder 1).',
    ),
    examples: [
      {
        title: L('Array 1-based', '1-based array', 'Matriz 1-based', 'Tableau 1-based', '1-basiertes Array'),
        code: `Option Base 1

Sub DemoOptionBase()
    Dim a(3) As Long
    a(1) = 10: a(2) = 20: a(3) = 30
    Debug.Print LBound(a), UBound(a)
End Sub`,
        result: L('Finestra Immediata: 1    3. Senza Option Base sarebbe 0    3.', 'Immediate Window: 1    3. Without Option Base it would be 0    3.', 'Ventana Inmediato: 1    3. Sin Option Base sería 0    3.', 'Fenêtre Exécution : 1    3. Sans Option Base ce serait 0    3.', 'Direktbereich: 1    3. Ohne Option Base wäre es 0    3.'),
      },
    ],
    notes: L(
      'Non cambia Array(), Split() né i Range di Excel, che restano 0-based o 1-based per conto proprio.',
      'Does not change Array(), Split(), or Excel Range arrays, which keep their own bounds.',
      'No cambia Array(), Split() ni las matrices de Range de Excel, que mantienen sus propios límites.',
      'Ne change pas Array(), Split() ni les tableaux Range d’Excel, qui gardent leurs propres bornes.',
      'Ändert weder Array() noch Split() noch Excel-Range-Arrays; die behalten eigene Grenzen.',
    ),
    related: ['array-bounds', 'dim', 'split'],
  },
  {
    id: 'option-private-module',
    name: 'Option Private Module',
    syntax: 'Option Private Module',
    scope: ['core', 'excel', 'access'],
    category: 'declarations',
    subcategory: 'options',
    description: L(
      'Nasconde il modulo all’esterno del progetto VBA (macro, elenco oggetti, altri progetti).',
      'Hides the module from outside the VBA project (macros, object browser, other projects).',
      'Oculta el módulo fuera del proyecto VBA (macros, examinador de objetos, otros proyectos).',
      'Masque le module à l’extérieur du projet VBA (macros, explorateur d’objets, autres projets).',
      'Verbirgt das Modul außerhalb des VBA-Projekts (Makros, Objektkatalog, andere Projekte).',
    ),
    examples: [
      {
        title: L('Modulo di sole routine interne', 'Module of internal-only routines', 'Módulo solo de rutinas internas', 'Module de routines internes uniquement', 'Modul nur für interne Routinen'),
        code: `Option Private Module

Public Sub DemoOptionPrivateModule()
    Debug.Print "visibile solo dentro questo progetto VBA"
End Sub`,
        result: L('La Sub esiste, ma non compare nell’elenco macro di Excel/Access.', 'The Sub exists, but does not appear in the Excel/Access macro list.', 'El Sub existe, pero no aparece en la lista de macros de Excel/Access.', 'Le Sub existe, mais n’apparaît pas dans la liste des macros Excel/Access.', 'Die Sub existiert, erscheint aber nicht in der Excel/Access-Makroliste.'),
      },
    ],
    related: ['public', 'private'],
  },
  {
    id: 'integer',
    name: 'Integer',
    syntax: 'Dim nome As Integer  \' -32 768 … 32 767',
    scope: ['core', 'excel', 'access'],
    category: 'types',
    subcategory: 'numeric',
    description: L(
      'Intero a 16 bit. In VBA moderno Long è più sicuro: Integer viene comunque allineato a 32 bit in memoria.',
      '16-bit integer. In modern VBA Long is safer: Integer is still 32-bit aligned in memory.',
      'Entero de 16 bits. En VBA moderno Long es más seguro: Integer sigue alineado a 32 bits en memoria.',
      'Entier 16 bits. En VBA moderne Long est plus sûr : Integer reste aligné sur 32 bits en mémoire.',
      '16-Bit-Ganzzahl. In modernem VBA ist Long sicherer: Integer bleibt im Speicher 32-Bit-ausgerichtet.',
    ),
    examples: [
      {
        title: L('Overflow oltre 32767', 'Overflow past 32767', 'Desbordamiento más allá de 32767', 'Dépassement au-delà de 32767', 'Überlauf über 32767'),
        code: `Sub DemoInteger()
    Dim n As Integer
    n = 32767
    Debug.Print n
    ' n = n + 1  ' errore 6: overflow
End Sub`,
        result: L('Finestra Immediata: 32767. n + 1 solleva overflow.', 'Immediate Window: 32767. n + 1 raises overflow.', 'Ventana Inmediato: 32767. n + 1 provoca desbordamiento.', 'Fenêtre Exécution : 32767. n + 1 lève un dépassement.', 'Direktbereich: 32767. n + 1 löst Überlauf aus.'),
      },
    ],
    related: ['long', 'c-numeric', 'byte'],
  },
  {
    id: 'long',
    name: 'Long',
    syntax: 'Dim nome As Long  \' -2 147 483 648 … 2 147 483 647',
    scope: ['core', 'excel', 'access'],
    category: 'types',
    subcategory: 'numeric',
    description: L(
      'Intero a 32 bit: scelta predefinita per contatori, indici e ID in VBA.',
      '32-bit integer: the default choice for counters, indexes and IDs in VBA.',
      'Entero de 32 bits: la opción predeterminada para contadores, índices e ID en VBA.',
      'Entier 32 bits : le choix par défaut pour compteurs, index et ID en VBA.',
      '32-Bit-Ganzzahl: Standardwahl für Zähler, Indizes und IDs in VBA.',
    ),
    examples: [
      {
        title: L('Contatore ampio', 'Wide counter', 'Contador amplio', 'Compteur large', 'Großer Zähler'),
        code: `Sub DemoLong()
    Dim n As Long
    n = 100000
    Debug.Print n * 2
End Sub`,
        result: L('Finestra Immediata: 200000', 'Immediate Window: 200000', 'Ventana Inmediato: 200000', 'Fenêtre Exécution : 200000', 'Direktbereich: 200000'),
      },
    ],
    related: ['integer', 'c-numeric', 'dim'],
  },
  {
    id: 'double-single',
    name: 'Double / Single',
    syntax: 'Dim d As Double | Dim s As Single',
    scope: ['core', 'excel', 'access'],
    category: 'types',
    subcategory: 'numeric',
    description: L(
      'Virgola mobile: Double (64 bit, ~15 cifre) è lo standard; Single (32 bit, ~7 cifre) occupa meno memoria.',
      'Floating point: Double (64-bit, ~15 digits) is the standard; Single (32-bit, ~7 digits) uses less memory.',
      'Punto flotante: Double (64 bits, ~15 dígitos) es el estándar; Single (32 bits, ~7 dígitos) usa menos memoria.',
      'Virgule flottante : Double (64 bits, ~15 chiffres) est le standard ; Single (32 bits, ~7 chiffres) prend moins de mémoire.',
      'Gleitkomma: Double (64 Bit, ~15 Stellen) ist Standard; Single (32 Bit, ~7 Stellen) braucht weniger Speicher.',
    ),
    examples: [
      {
        title: L('Divisione reale', 'Real division', 'División real', 'Division réelle', 'Reelle Division'),
        code: `Sub DemoDoubleSingle()
    Dim d As Double, s As Single
    d = 10 / 3
    s = 10 / 3
    Debug.Print d
    Debug.Print s
End Sub`,
        result: L('Double ≈ 3.33333333333333; Single ≈ 3.333333 (meno cifre).', 'Double ≈ 3.33333333333333; Single ≈ 3.333333 (fewer digits).', 'Double ≈ 3.33333333333333; Single ≈ 3.333333 (menos dígitos).', 'Double ≈ 3.33333333333333 ; Single ≈ 3.333333 (moins de chiffres).', 'Double ≈ 3.33333333333333; Single ≈ 3.333333 (weniger Stellen).'),
      },
    ],
    notes: L(
      'Non usare Double per soldi: preferisci Currency. 0.1 + 0.2 può non essere esatto.',
      'Do not use Double for money: prefer Currency. 0.1 + 0.2 may not be exact.',
      'No uses Double para dinero: prefiere Currency. 0.1 + 0.2 puede no ser exacto.',
      'N’utilisez pas Double pour l’argent : préférez Currency. 0.1 + 0.2 peut ne pas être exact.',
      'Double nicht für Geld verwenden: Currency bevorzugen. 0.1 + 0.2 kann ungenau sein.',
    ),
    related: ['currency', 'c-numeric', 'operators-arithmetic'],
  },
  {
    id: 'currency',
    name: 'Currency',
    syntax: 'Dim nome As Currency  \' 4 decimali fissi, ±922 337 203 685 477.5807',
    scope: ['core', 'excel', 'access'],
    category: 'types',
    subcategory: 'numeric',
    description: L(
      'Decimale a scala fissa (4 cifre dopo la virgola): adatto a importi in denaro senza errori di binario.',
      'Fixed-scale decimal (4 digits after the point): suited to money amounts without binary rounding error.',
      'Decimal de escala fija (4 dígitos tras el punto): adecuado para dinero sin error binario.',
      'Décimal à échelle fixe (4 chiffres après la virgule) : adapté aux montants sans erreur binaire.',
      'Dezimal mit fester Skala (4 Nachkommastellen): geeignet für Geldbeträge ohne Binärfehler.',
    ),
    examples: [
      {
        title: L('Due importi sommati', 'Two amounts added', 'Dos importes sumados', 'Deux montants additionnés', 'Zwei Beträge addiert'),
        code: `Sub DemoCurrency()
    Dim a As Currency, b As Currency
    a = 10.1055
    b = 0.0002
    Debug.Print a + b
End Sub`,
        result: L('Finestra Immediata: 10.1057', 'Immediate Window: 10.1057', 'Ventana Inmediato: 10.1057', 'Fenêtre Exécution : 10.1057', 'Direktbereich: 10.1057'),
      },
    ],
    related: ['double-single', 'c-numeric', 'format-helpers'],
  },
  {
    id: 'byte',
    name: 'Byte',
    syntax: 'Dim nome As Byte  \' 0 … 255',
    scope: ['core', 'excel', 'access'],
    category: 'types',
    subcategory: 'numeric',
    description: L(
      'Intero senza segno a 8 bit. Utile per RGB, buffer e flag 0–255.',
      'Unsigned 8-bit integer. Useful for RGB, buffers and 0–255 flags.',
      'Entero sin signo de 8 bits. Útil para RGB, búferes y flags 0–255.',
      'Entier non signé 8 bits. Utile pour RGB, tampons et drapeaux 0–255.',
      'Vorzeichenlose 8-Bit-Ganzzahl. Nützlich für RGB, Puffer und Flags 0–255.',
    ),
    examples: [
      {
        title: L('Canale di colore', 'Color channel', 'Canal de color', 'Canal de couleur', 'Farbkanal'),
        code: `Sub DemoByte()
    Dim r As Byte
    r = 200
    Debug.Print r
    ' r = -1  ' overflow
End Sub`,
        result: L('Finestra Immediata: 200. Un negativo provoca overflow.', 'Immediate Window: 200. A negative value overflows.', 'Ventana Inmediato: 200. Un negativo provoca desbordamiento.', 'Fenêtre Exécution : 200. Une valeur négative déborde.', 'Direktbereich: 200. Ein negativer Wert löst Überlauf aus.'),
      },
    ],
    related: ['integer', 'rgb', 'c-numeric'],
  },
  {
    id: 'string',
    name: 'String',
    syntax: 'Dim s As String | Dim fissa As String * n',
    scope: ['core', 'excel', 'access'],
    category: 'types',
    subcategory: 'text-date',
    description: L(
      'Testo Unicode a lunghezza variabile, oppure String * n a lunghezza fissa (spazi di padding).',
      'Variable-length Unicode text, or fixed-length String * n (padded with spaces).',
      'Texto Unicode de longitud variable, o String * n de longitud fija (relleno con espacios).',
      'Texte Unicode de longueur variable, ou String * n de longueur fixe (complété d’espaces).',
      'Unicode-Text variabler Länge oder festes String * n (mit Leerzeichen aufgefüllt).',
    ),
    examples: [
      {
        title: L('Concatenazione e lunghezza fissa', 'Concatenation and fixed length', 'Concatenación y longitud fija', 'Concaténation et longueur fixe', 'Verkettung und feste Länge'),
        code: `Sub DemoString()
    Dim s As String, fissa As String * 5
    s = "VBA"
    fissa = "ok"
    Debug.Print s & "!"
    Debug.Print Len(fissa), "'" & fissa & "'"
End Sub`,
        result: L('VBA! poi 5 e \'ok   \' (due spazi di padding).', 'VBA! then 5 and \'ok   \' (two padding spaces).', 'VBA! luego 5 y \'ok   \' (dos espacios de relleno).', 'VBA ! puis 5 et \'ok   \' (deux espaces de remplissage).', 'VBA! dann 5 und \'ok   \' (zwei Füllleerzeichen).'),
      },
    ],
    related: ['len', 'left-right-mid', 'operators-arithmetic'],
  },
  {
    id: 'boolean',
    name: 'Boolean',
    syntax: 'Dim nome As Boolean  \' True / False',
    scope: ['core', 'excel', 'access'],
    category: 'types',
    subcategory: 'text-date',
    description: L(
      'Valore logico True o False. In calcoli numerici True è -1 e False è 0.',
      'Logical value True or False. In numeric math True is -1 and False is 0.',
      'Valor lógico True o False. En cálculos numéricos True es -1 y False es 0.',
      'Valeur logique True ou False. En calcul numérique True vaut -1 et False 0.',
      'Logischer Wert True oder False. In Zahlenrechnungen ist True -1 und False 0.',
    ),
    examples: [
      {
        title: L('True come numero', 'True as a number', 'True como número', 'True en tant que nombre', 'True als Zahl'),
        code: `Sub DemoBoolean()
    Dim ok As Boolean
    ok = (5 > 2)
    Debug.Print ok
    Debug.Print CInt(ok)
End Sub`,
        result: L('Finestra Immediata: True poi -1', 'Immediate Window: True then -1', 'Ventana Inmediato: True luego -1', 'Fenêtre Exécution : True puis -1', 'Direktbereich: True dann -1'),
      },
    ],
    related: ['operators-logic-compare', 'c-other', 'if-then-else'],
  },
  {
    id: 'date-type',
    name: 'Date',
    syntax: 'Dim nome As Date',
    scope: ['core', 'excel', 'access'],
    category: 'types',
    subcategory: 'text-date',
    description: L(
      'Data/ora internamente come Double (parte intera = giorno da 30/12/1899, frazione = ora).',
      'Date/time stored as a Double (integer part = days since 1899-12-30, fraction = time of day).',
      'Fecha/hora almacenada como Double (parte entera = días desde 1899-12-30, fracción = hora).',
      'Date/heure stockée en Double (partie entière = jours depuis 1899-12-30, fraction = heure).',
      'Datum/Uhrzeit intern als Double (Ganzzahlteil = Tage seit 1899-12-30, Bruch = Uhrzeit).',
    ),
    examples: [
      {
        title: L('Letterale data', 'Date literal', 'Literal de fecha', 'Littéral de date', 'Datumsliteral'),
        code: `Sub DemoDateType()
    Dim d As Date
    d = #3/15/2024 14:30:00#
    Debug.Print Year(d), Hour(d)
End Sub`,
        result: L('Finestra Immediata: 2024    14', 'Immediate Window: 2024    14', 'Ventana Inmediato: 2024    14', 'Fenêtre Exécution : 2024    14', 'Direktbereich: 2024    14'),
      },
    ],
    notes: L(
      'I letterali #m/g/aaaa# seguono spesso il formato USA indipendentemente dal locale.',
      'Literals #m/d/yyyy# often follow US order regardless of locale.',
      'Los literales #m/d/yyyy# suelen seguir el orden de EE. UU. independientemente de la configuración regional.',
      'Les littéraux #m/d/yyyy# suivent souvent l’ordre US quel que soit le paramètre régional.',
      'Literale #m/d/yyyy# folgen oft der US-Reihenfolge unabhängig vom Gebietsschema.',
    ),
    related: ['now', 'date-parts', 'c-other'],
  },
  {
    id: 'variant',
    name: 'Variant',
    syntax: 'Dim nome As Variant  \' oppure Dim nome',
    scope: ['core', 'excel', 'access'],
    category: 'types',
    subcategory: 'text-date',
    description: L(
      'Contenitore che può essere Empty, Null, numero, stringa, data, array o oggetto. Flessibile ma più lento.',
      'Container that can hold Empty, Null, a number, string, date, array or object. Flexible but slower.',
      'Contenedor que puede ser Empty, Null, número, cadena, fecha, matriz u objeto. Flexible pero más lento.',
      'Conteneur qui peut être Empty, Null, nombre, chaîne, date, tableau ou objet. Souple mais plus lent.',
      'Behälter für Empty, Null, Zahl, Zeichenfolge, Datum, Array oder Objekt. Flexibel, aber langsamer.',
    ),
    examples: [
      {
        title: L('Empty poi numero poi testo', 'Empty then number then text', 'Empty luego número luego texto', 'Empty puis nombre puis texte', 'Empty dann Zahl dann Text'),
        code: `Sub DemoVariant()
    Dim v As Variant
    Debug.Print IsEmpty(v), TypeName(v)
    v = 12
    Debug.Print TypeName(v), v
    v = "ciao"
    Debug.Print TypeName(v), v
End Sub`,
        result: L('True Empty / Integer 12 / String ciao', 'True Empty / Integer 12 / String ciao', 'True Empty / Integer 12 / String ciao', 'True Empty / Integer 12 / String ciao', 'True Empty / Integer 12 / String ciao'),
      },
    ],
    related: ['is-functions', 'typename-vartype', 'dim'],
  },
  {
    id: 'object',
    name: 'Object',
    syntax: 'Dim nome As Object',
    scope: ['core', 'excel', 'access'],
    category: 'types',
    subcategory: 'text-date',
    description: L(
      'Riferimento tardivo (late binding) a qualsiasi oggetto COM. Serve Set per assegnare.',
      'Late-bound reference to any COM object. Assignment requires Set.',
      'Referencia de enlace tardío a cualquier objeto COM. La asignación requiere Set.',
      'Référence en liaison tardive vers n’importe quel objet COM. L’affectation exige Set.',
      'Spät gebunden Referenz auf ein beliebiges COM-Objekt. Zuweisung erfordert Set.',
    ),
    examples: [
      {
        title: L('Late binding su Dictionary', 'Late binding a Dictionary', 'Late binding a Dictionary', 'Liaison tardive sur Dictionary', 'Late Binding auf Dictionary'),
        code: `Sub DemoObject()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "k", 9
    Debug.Print d("k"), TypeName(d)
    Set d = Nothing
End Sub`,
        result: L('Finestra Immediata: 9    Dictionary', 'Immediate Window: 9    Dictionary', 'Ventana Inmediato: 9    Dictionary', 'Fenêtre Exécution : 9    Dictionary', 'Direktbereich: 9    Dictionary'),
      },
    ],
    related: ['set-new-nothing', 'createobject-getobject', 'collection'],
  },
  {
    id: 'collection',
    name: 'Collection',
    syntax: 'Dim c As New Collection  \' Add / Item / Count / Remove',
    scope: ['core', 'excel', 'access'],
    category: 'types',
    subcategory: 'text-date',
    description: L(
      'Elenco ordinato 1-based di elementi Variant, con chiave stringa opzionale.',
      'Ordered 1-based list of Variant items, with an optional string key.',
      'Lista ordenada 1-based de elementos Variant, con clave de cadena opcional.',
      'Liste ordonnée 1-based d’éléments Variant, avec une clé chaîne facultative.',
      'Geordnete 1-basierte Liste von Variant-Elementen, optional mit Zeichenfolgeschlüssel.',
    ),
    params: [
      { name: 'item', description: L('Valore da aggiungere (qualsiasi tipo).', 'Value to add (any type).', 'Valor a agregar (cualquier tipo).', 'Valeur à ajouter (tout type).', 'Hinzuzufügender Wert (jeder Typ).') },
      { name: 'key', optional: true, description: L('Chiave univoca stringa.', 'Unique string key.', 'Clave de cadena única.', 'Clé chaîne unique.', 'Eindeutiger Zeichenfolgeschlüssel.') },
    ],
    examples: [
      {
        title: L('Add con chiave e For Each', 'Add with key and For Each', 'Add con clave y For Each', 'Add avec clé et For Each', 'Add mit Schlüssel und For Each'),
        code: `Sub DemoCollection()
    Dim c As New Collection, v As Variant
    c.Add "Roma", "RM"
    c.Add "Milano", "MI"
    Debug.Print c.Count, c("RM")
    For Each v In c
        Debug.Print v
    Next v
End Sub`,
        result: L('2    Roma poi Roma e Milano (una riga ciascuno).', '2    Roma then Roma and Milano (one line each).', '2    Roma luego Roma y Milano (una línea cada uno).', '2    Roma puis Roma et Milano (une ligne chacun).', '2    Roma dann Roma und Milano (je eine Zeile).'),
      },
    ],
    notes: L(
      'Item è 1-based. Una chiave duplicata solleva errore 457. Non è un Dictionary (niente Exists nativo).',
      'Item is 1-based. A duplicate key raises error 457. It is not a Dictionary (no native Exists).',
      'Item es 1-based. Una clave duplicada provoca el error 457. No es un Dictionary (no hay Exists nativo).',
      'Item est 1-based. Une clé en double lève l’erreur 457. Ce n’est pas un Dictionary (pas d’Exists natif).',
      'Item ist 1-basiert. Ein doppelter Schlüssel löst Fehler 457 aus. Kein Dictionary (kein natives Exists).',
    ),
    related: ['set-new-nothing', 'for-each', 'object'],
  },
  {
    id: 'operators-arithmetic',
    name: 'Operatori aritmetici',
    syntax: 'a + b | a - b | a * b | a / b | a \\ b | a Mod b | a ^ b | a & b',
    scope: ['core', 'excel', 'access'],
    category: 'operators',
    subcategory: 'arithmetic',
    description: L(
      'Somma, sottrazione, prodotto, divisione reale (/), intera (\\), resto (Mod), potenza (^) e concatenazione (&).',
      'Addition, subtraction, product, real division (/), integer division (\\), remainder (Mod), power (^) and concatenation (&).',
      'Suma, resta, producto, división real (/), entera (\\), resto (Mod), potencia (^) y concatenación (&).',
      'Addition, soustraction, produit, division réelle (/), entière (\\), reste (Mod), puissance (^) et concaténation (&).',
      'Addition, Subtraktion, Produkt, reelle Division (/), Ganzzahldivision (\\), Rest (Mod), Potenz (^) und Verkettung (&).',
    ),
    examples: [
      {
        title: L('Tutti gli operatori su 5 e 2', 'Every operator on 5 and 2', 'Todos los operadores sobre 5 y 2', 'Tous les opérateurs sur 5 et 2', 'Alle Operatoren mit 5 und 2'),
        code: `Sub DemoArithmetic()
    Debug.Print 5 + 2, 5 - 2, 5 * 2
    Debug.Print 5 / 2, 5 \\ 2, 5 Mod 2
    Debug.Print 2 ^ 3, "A" & "B"
End Sub`,
        result: L('7 3 10 / 2.5 2 1 / 8 AB', '7 3 10 / 2.5 2 1 / 8 AB', '7 3 10 / 2.5 2 1 / 8 AB', '7 3 10 / 2.5 2 1 / 8 AB', '7 3 10 / 2.5 2 1 / 8 AB'),
      },
    ],
    notes: L(
      '& converte gli operandi in String. + su due stringhe le concatena, ma su numeri le somma: preferisci & per il testo.',
      '& coerces operands to String. + on two strings concatenates, but on numbers it adds: prefer & for text.',
      '& convierte los operandos en String. + en dos cadenas concatena, pero en números suma: prefiere & para texto.',
      '& convertit les opérandes en String. + sur deux chaînes concatène, mais sur des nombres additionne : préférez & pour le texte.',
      '& wandelt Operanden in String. + verkettet zwei Zeichenfolgen, addiert aber Zahlen: für Text & bevorzugen.',
    ),
    related: ['operators-logic-compare', 'string', 'int-fix'],
  },
  {
    id: 'operators-logic-compare',
    name: 'Operatori logici e di confronto',
    syntax: 'a = b | a <> b | a < b | a > b | And | Or | Not | Xor | Like | Is',
    scope: ['core', 'excel', 'access'],
    category: 'operators',
    subcategory: 'logic-compare',
    description: L(
      'Confronti (= <> < >), logica bit a bit (And Or Not Xor), pattern Like e identità di oggetti Is.',
      'Comparisons (= <> < >), bitwise logic (And Or Not Xor), Like patterns and object identity Is.',
      'Comparaciones (= <> < >), lógica bit a bit (And Or Not Xor), patrones Like e identidad de objetos Is.',
      'Comparaisons (= <> < >), logique bit à bit (And Or Not Xor), motifs Like et identité d’objet Is.',
      'Vergleiche (= <> < >), bitweise Logik (And Or Not Xor), Like-Muster und Objektidentität Is.',
    ),
    examples: [
      {
        title: L('Like, logica e Is Nothing', 'Like, logic and Is Nothing', 'Like, lógica e Is Nothing', 'Like, logique et Is Nothing', 'Like, Logik und Is Nothing'),
        code: `Sub DemoLogicCompare()
    Debug.Print 3 > 2, 3 <> 3
    Debug.Print True And False, True Or False, Not False, True Xor True
    Debug.Print "file.txt" Like "*.txt"
    Debug.Print New Collection Is Nothing
End Sub`,
        result: L('True False / False True True False / True / False', 'True False / False True True False / True / False', 'True False / False True True False / True / False', 'True False / False True True False / True / False', 'True False / False True True False / True / False'),
      },
    ],
    notes: L(
      'And/Or non sono cortocircuito: entrambe le parti vengono valutate. Like: ? un carattere, * sequenza, # cifra, [A-C] elenco.',
      'And/Or are not short-circuit: both sides are evaluated. Like: ? one char, * sequence, # digit, [A-C] list.',
      'And/Or no son de cortocircuito: se evalúan ambos lados. Like: ? un carácter, * secuencia, # dígito, [A-C] lista.',
      'And/Or ne sont pas en court-circuit : les deux côtés sont évalués. Like : ? un caractère, * une suite, # un chiffre, [A-C] liste.',
      'And/Or sind nicht kurzschließend: beide Seiten werden ausgewertet. Like: ? ein Zeichen, * Folge, # Ziffer, [A-C] Liste.',
    ),
    related: ['if-then-else', 'boolean', 'set-new-nothing', 'option-compare'],
  },
  {
    id: 'if-then-else',
    name: 'If Then Else',
    syntax: 'If cond Then\n    ...\nElseIf cond2 Then\n    ...\nElse\n    ...\nEnd If',
    scope: ['core', 'excel', 'access'],
    category: 'control',
    subcategory: 'branch',
    description: L(
      'Sceglie un blocco in base a una condizione. ElseIf aggiunge rami; Else è il resto.',
      'Chooses a block based on a condition. ElseIf adds branches; Else is the remainder.',
      'Elige un bloque según una condición. ElseIf añade ramas; Else es el resto.',
      'Choisit un bloc selon une condition. ElseIf ajoute des branches ; Else est le reste.',
      'Wählt einen Block anhand einer Bedingung. ElseIf fügt Zweige hinzu; Else ist der Rest.',
    ),
    params: [
      { name: 'cond', description: L('Espressione che VBA tratta come True se diversa da 0/False/Null.', 'Expression VBA treats as True when not 0/False/Null.', 'Expresión que VBA trata como True si no es 0/False/Null.', 'Expression que VBA traite comme True si elle n’est pas 0/False/Null.', 'Ausdruck, den VBA als True behandelt, wenn nicht 0/False/Null.') },
    ],
    examples: [
      {
        title: L('Voto in fasce', 'Grade bands', 'Nota por tramos', 'Note par tranches', 'Note nach Stufen'),
        code: `Sub DemoIf()
    Dim voto As Long: voto = 24
    If voto >= 27 Then
        Debug.Print "ottimo"
    ElseIf voto >= 18 Then
        Debug.Print "ok"
    Else
        Debug.Print "insufficiente"
    End If
End Sub`,
        result: L('Finestra Immediata: ok', 'Immediate Window: ok', 'Ventana Inmediato: ok', 'Fenêtre Exécution : ok', 'Direktbereich: ok'),
      },
    ],
    related: ['select-case', 'iif', 'operators-logic-compare'],
  },
  {
    id: 'select-case',
    name: 'Select Case',
    syntax: 'Select Case expr\n    Case v1, v2\n    Case n To m\n    Case Is > n\n    Case Else\nEnd Select',
    scope: ['core', 'excel', 'access'],
    category: 'control',
    subcategory: 'branch',
    description: L(
      'Confronta un’espressione con elenchi, intervalli (To) e confronti (Is). Esegue il primo Case vero.',
      'Matches an expression against lists, ranges (To) and comparisons (Is). Runs the first true Case.',
      'Compara una expresión con listas, intervalos (To) y comparaciones (Is). Ejecuta el primer Case verdadero.',
      'Compare une expression à des listes, plages (To) et comparaisons (Is). Exécute le premier Case vrai.',
      'Vergleicht einen Ausdruck mit Listen, Bereichen (To) und Vergleichen (Is). Führt den ersten wahren Case aus.',
    ),
    examples: [
      {
        title: L('Giorno lavorativo', 'Working day', 'Día laborable', 'Jour ouvré', 'Werktag'),
        code: `Sub DemoSelectCase()
    Dim g As Long: g = 6
    Select Case g
        Case 1 To 5
            Debug.Print "feriale"
        Case 6, 7
            Debug.Print "weekend"
        Case Else
            Debug.Print "invalido"
    End Select
End Sub`,
        result: L('Finestra Immediata: weekend', 'Immediate Window: weekend', 'Ventana Inmediato: weekend', 'Fenêtre Exécution : weekend', 'Direktbereich: weekend'),
      },
    ],
    related: ['if-then-else', 'choose-switch'],
  },
  {
    id: 'iif',
    name: 'IIf',
    syntax: 'IIf(condizione, seVero, seFalso)',
    scope: ['core', 'excel', 'access'],
    category: 'control',
    subcategory: 'branch',
    description: L(
      'Restituisce seVero o seFalso. Attenzione: valuta sempre entrambi gli argomenti (effetti collaterali e divisioni per zero).',
      'Returns seVero or seFalso. Warning: always evaluates both arguments (side effects and division by zero).',
      'Devuelve seVero o seFalso. Atención: evalúa siempre ambos argumentos (efectos secundarios y división por cero).',
      'Renvoie seVero ou seFalso. Attention : évalue toujours les deux arguments (effets de bord et division par zéro).',
      'Gibt seVero oder seFalso zurück. Achtung: wertet immer beide Argumente aus (Nebeneffekte und Division durch 0).',
    ),
    params: [
      { name: 'condizione', description: L('Test logico.', 'Logical test.', 'Prueba lógica.', 'Test logique.', 'Logischer Test.') },
      { name: 'seVero', description: L('Valore se la condizione è True.', 'Value if the condition is True.', 'Valor si la condición es True.', 'Valeur si la condition est True.', 'Wert wenn die Bedingung True ist.') },
      { name: 'seFalso', description: L('Valore se la condizione è False.', 'Value if the condition is False.', 'Valor si la condición es False.', 'Valeur si la condition est False.', 'Wert wenn die Bedingung False ist.') },
    ],
    examples: [
      {
        title: L('Etichetta da un flag', 'Label from a flag', 'Etiqueta desde un flag', 'Libellé depuis un drapeau', 'Beschriftung aus einem Flag'),
        code: `Sub DemoIIf()
    Dim attivo As Boolean: attivo = True
    Debug.Print IIf(attivo, "on", "off")
End Sub`,
        result: L('Finestra Immediata: on', 'Immediate Window: on', 'Ventana Inmediato: on', 'Fenêtre Exécution : on', 'Direktbereich: on'),
      },
    ],
    related: ['if-then-else', 'choose-switch'],
  },
  {
    id: 'with',
    name: 'With',
    syntax: 'With oggetto\n    .proprietà = valore\n    .metodo\nEnd With',
    scope: ['core', 'excel', 'access'],
    category: 'control',
    subcategory: 'branch',
    description: L(
      'Fissa un oggetto di riferimento: i membri che iniziano con il punto si applicano a quell’oggetto.',
      'Pins a reference object: members that start with a dot apply to that object.',
      'Fija un objeto de referencia: los miembros que empiezan por punto se aplican a ese objeto.',
      'Fixe un objet de référence : les membres commençant par un point s’appliquent à cet objet.',
      'Fixiert ein Bezugsobjekt: Mitglieder mit führendem Punkt gelten für dieses Objekt.',
    ),
    examples: [
      {
        title: L('Due proprietà su una Collection', 'Two properties on a Collection', 'Dos propiedades de una Collection', 'Deux propriétés d’une Collection', 'Zwei Eigenschaften einer Collection'),
        code: `Sub DemoWith()
    Dim c As New Collection
    With c
        .Add "a"
        .Add "b"
        Debug.Print .Count
    End With
End Sub`,
        result: L('Finestra Immediata: 2', 'Immediate Window: 2', 'Ventana Inmediato: 2', 'Fenêtre Exécution : 2', 'Direktbereich: 2'),
      },
    ],
    related: ['object', 'set-new-nothing'],
  },
  {
    id: 'for-next',
    name: 'For Next',
    syntax: 'For i = inizio To fine [Step passo]\n    ...\nNext i',
    scope: ['core', 'excel', 'access'],
    category: 'control',
    subcategory: 'loops',
    description: L(
      'Ripete un blocco con un contatore numerico. Step può essere negativo. Il limite viene valutato una sola volta.',
      'Repeats a block with a numeric counter. Step may be negative. The end bound is evaluated once.',
      'Repite un bloque con un contador numérico. Step puede ser negativo. El límite se evalúa una sola vez.',
      'Répète un bloc avec un compteur numérique. Step peut être négatif. La borne finale n’est évaluée qu’une fois.',
      'Wiederholt einen Block mit numerischem Zähler. Step darf negativ sein. Die Endgrenze wird nur einmal ausgewertet.',
    ),
    examples: [
      {
        title: L('Somma 1..5 e countdown', 'Sum 1..5 and countdown', 'Suma 1..5 y cuenta atrás', 'Somme 1..5 et compte à rebours', 'Summe 1..5 und Countdown'),
        code: `Sub DemoForNext()
    Dim i As Long, tot As Long
    For i = 1 To 5
        tot = tot + i
    Next i
    Debug.Print tot
    For i = 3 To 1 Step -1
        Debug.Print i
    Next i
End Sub`,
        result: L('15 poi 3, 2, 1', '15 then 3, 2, 1', '15 luego 3, 2, 1', '15 puis 3, 2, 1', '15 dann 3, 2, 1'),
      },
    ],
    related: ['for-each', 'exit', 'do-loop'],
  },
  {
    id: 'for-each',
    name: 'For Each',
    syntax: 'For Each elemento In collezione\n    ...\nNext elemento',
    scope: ['core', 'excel', 'access'],
    category: 'control',
    subcategory: 'loops',
    description: L(
      'Scorre ogni elemento di una Collection, array o insieme COM (es. Worksheets).',
      'Walks every item in a Collection, array or COM set (e.g. Worksheets).',
      'Recorre cada elemento de una Collection, matriz o conjunto COM (p. ej. Worksheets).',
      'Parcourt chaque élément d’une Collection, d’un tableau ou d’un ensemble COM (ex. Worksheets).',
      'Durchläuft jedes Element einer Collection, eines Arrays oder einer COM-Menge (z. B. Worksheets).',
    ),
    examples: [
      {
        title: L('Elementi di un array', 'Items of an array', 'Elementos de una matriz', 'Éléments d’un tableau', 'Elemente eines Arrays'),
        code: `Sub DemoForEach()
    Dim v As Variant, n As Long
    For Each v In Array("a", "b", "c")
        n = n + 1
        Debug.Print n, v
    Next v
End Sub`,
        result: L('1 a / 2 b / 3 c', '1 a / 2 b / 3 c', '1 a / 2 b / 3 c', '1 a / 2 b / 3 c', '1 a / 2 b / 3 c'),
      },
    ],
    related: ['for-next', 'collection', 'array-bounds', 'exit'],
  },
  {
    id: 'do-loop',
    name: 'Do While / Do Until',
    syntax: 'Do While cond ... Loop | Do Until cond ... Loop | Do ... Loop While cond | Do ... Loop Until cond',
    scope: ['core', 'excel', 'access'],
    category: 'control',
    subcategory: 'loops',
    description: L(
      'Ciclo condizionale. While continua se True; Until si ferma quando True. La condizione in coda esegue il corpo almeno una volta.',
      'Conditional loop. While continues while True; Until stops when True. A trailing condition runs the body at least once.',
      'Bucle condicional. While sigue si True; Until para cuando True. La condición al final ejecuta el cuerpo al menos una vez.',
      'Boucle conditionnelle. While continue tant que True ; Until s’arrête quand True. La condition en fin exécute le corps au moins une fois.',
      'Bedingte Schleife. While läuft solange True; Until stoppt wenn True. Bedingung am Ende führt den Rumpf mindestens einmal aus.',
    ),
    examples: [
      {
        title: L('While cresce, Until cala', 'While grows, Until shrinks', 'While crece, Until mengua', 'While grandit, Until diminue', 'While wächst, Until schrumpft'),
        code: `Sub DemoDoLoop()
    Dim n As Long
    n = 1
    Do While n < 4
        Debug.Print "W" & n
        n = n + 1
    Loop
    n = 3
    Do Until n = 0
        Debug.Print "U" & n
        n = n - 1
    Loop
End Sub`,
        result: L('W1 W2 W3 poi U3 U2 U1', 'W1 W2 W3 then U3 U2 U1', 'W1 W2 W3 luego U3 U2 U1', 'W1 W2 W3 puis U3 U2 U1', 'W1 W2 W3 dann U3 U2 U1'),
      },
    ],
    related: ['while-wend', 'exit', 'for-next'],
  },
  {
    id: 'while-wend',
    name: 'While Wend',
    syntax: 'While condizione\n    ...\nWend',
    scope: ['core', 'excel', 'access'],
    category: 'control',
    subcategory: 'loops',
    description: L(
      'Ciclo storico equivalente a Do While ... Loop. Preferisci Do Loop: supporta Exit Do e Until.',
      'Legacy loop equivalent to Do While ... Loop. Prefer Do Loop: it supports Exit Do and Until.',
      'Bucle histórico equivalente a Do While ... Loop. Prefiere Do Loop: admite Exit Do y Until.',
      'Boucle historique équivalente à Do While ... Loop. Préférez Do Loop : elle gère Exit Do et Until.',
      'Historische Schleife, gleichwertig zu Do While ... Loop. Do Loop bevorzugen: unterstützt Exit Do und Until.',
    ),
    examples: [
      {
        title: L('Conteggio fino a 3', 'Count up to 3', 'Contar hasta 3', 'Compter jusqu’à 3', 'Zählen bis 3'),
        code: `Sub DemoWhileWend()
    Dim n As Long
    n = 1
    While n <= 3
        Debug.Print n
        n = n + 1
    Wend
End Sub`,
        result: L('Finestra Immediata: 1, 2, 3', 'Immediate Window: 1, 2, 3', 'Ventana Inmediato: 1, 2, 3', 'Fenêtre Exécution : 1, 2, 3', 'Direktbereich: 1, 2, 3'),
      },
    ],
    related: ['do-loop', 'exit'],
  },
  {
    id: 'exit',
    name: 'Exit',
    syntax: 'Exit For | Exit Do | Exit Sub | Exit Function | Exit Property',
    scope: ['core', 'excel', 'access'],
    category: 'control',
    subcategory: 'jumps',
    description: L(
      'Esce subito dal ciclo o dalla procedura corrente, senza eseguire il resto del blocco.',
      'Leaves the current loop or procedure immediately, skipping the rest of the block.',
      'Sale inmediatamente del bucle o procedimiento actual, sin ejecutar el resto del bloque.',
      'Quitte immédiatement la boucle ou la procédure courante, sans exécuter le reste du bloc.',
      'Verlässt die aktuelle Schleife oder Prozedur sofort, ohne den Rest des Blocks auszuführen.',
    ),
    examples: [
      {
        title: L('Stop al primo pari e uscita anticipata', 'Stop at first even and early return', 'Parar en el primer par y salida temprana', 'Stop au premier pair et retour anticipé', 'Stop beim ersten Geraden und früher Rücksprung'),
        code: `Function PrimoPari(ParamArray nums() As Variant) As Variant
    Dim v As Variant
    For Each v In nums
        If v Mod 2 = 0 Then
            PrimoPari = v
            Exit Function
        End If
    Next v
    PrimoPari = Null
End Function

Sub DemoExit()
    Dim i As Long
    For i = 1 To 10
        If i = 3 Then Exit For
        Debug.Print i
    Next i
    Debug.Print PrimoPari(1, 3, 8, 9)
End Sub`,
        result: L('1, 2 poi 8', '1, 2 then 8', '1, 2 luego 8', '1, 2 puis 8', '1, 2 dann 8'),
      },
    ],
    related: ['for-next', 'do-loop', 'sub', 'function', 'goto'],
  },
  {
    id: 'goto',
    name: 'GoTo',
    syntax: 'GoTo etichetta',
    scope: ['core', 'excel', 'access'],
    category: 'control',
    subcategory: 'jumps',
    description: L(
      'Salta a un’etichetta nella stessa procedura. Usalo soprattutto con On Error, non per strutturare la logica.',
      'Jumps to a label in the same procedure. Use it mainly with On Error, not to structure logic.',
      'Salta a una etiqueta en el mismo procedimiento. Úsalo sobre todo con On Error, no para estructurar la lógica.',
      'Saute à une étiquette dans la même procédure. Réservez-le surtout à On Error, pas à structurer la logique.',
      'Springt zu einer Marke in derselben Prozedur. Vor allem mit On Error nutzen, nicht zur Strukturierung.',
    ),
    examples: [
      {
        title: L('Salto a pulizia', 'Jump to cleanup', 'Salto a limpieza', 'Saut vers le nettoyage', 'Sprung zur Bereinigung'),
        code: `Sub DemoGoTo()
    Dim n As Long: n = 1
    GoTo Pulizia
    n = 99
Pulizia:
    Debug.Print n
End Sub`,
        result: L('Finestra Immediata: 1 (la riga n = 99 non viene eseguita).', 'Immediate Window: 1 (the n = 99 line never runs).', 'Ventana Inmediato: 1 (la línea n = 99 no se ejecuta).', 'Fenêtre Exécution : 1 (la ligne n = 99 n’est pas exécutée).', 'Direktbereich: 1 (die Zeile n = 99 wird nicht ausgeführt).'),
      },
    ],
    related: ['on-error', 'exit', 'end-stop'],
  },
  {
    id: 'end-stop',
    name: 'End / Stop',
    syntax: 'End | Stop',
    scope: ['core', 'excel', 'access'],
    category: 'control',
    subcategory: 'jumps',
    description: L(
      'End termina l’intero progetto VBA (reset variabili Public/Static). Stop entra in modalità interruzione nel debugger.',
      'End terminates the whole VBA project (resets Public/Static variables). Stop enters break mode in the debugger.',
      'End termina todo el proyecto VBA (restablece variables Public/Static). Stop entra en modo interrupción del depurador.',
      'End termine tout le projet VBA (réinitialise les variables Public/Static). Stop passe en mode Arrêt du débogueur.',
      'End beendet das gesamte VBA-Projekt (setzt Public/Static zurück). Stop wechselt in den Haltemodus des Debuggers.',
    ),
    examples: [
      {
        title: L('Stop per ispezionare, End per azzerare', 'Stop to inspect, End to wipe state', 'Stop para inspeccionar, End para borrar estado', 'Stop pour inspecter, End pour tout réinitialiser', 'Stop zum Prüfen, End zum Zurücksetzen'),
        code: `Sub DemoEndStop()
    Dim x As Long: x = 5
    Debug.Print x
    Stop          ' il debugger si ferma qui
    ' End         ' decommentare per chiudere l'esecuzione VBA
End Sub`,
        result: L('Stampa 5 e apre il debugger su Stop. End chiuderebbe tutto il runtime.', 'Prints 5 and opens the debugger on Stop. End would shut down the whole runtime.', 'Imprime 5 y abre el depurador en Stop. End cerraría todo el runtime.', 'Affiche 5 et ouvre le débogueur sur Stop. End arrêterait tout le runtime.', 'Gibt 5 aus und öffnet den Debugger bei Stop. End würde die gesamte Runtime beenden.'),
      },
    ],
    notes: L(
      'Evita End in produzione: chiude file aperti e perde lo stato. Preferisci Exit Sub dopo aver rilasciato gli oggetti.',
      'Avoid End in production: it closes open files and drops state. Prefer Exit Sub after releasing objects.',
      'Evita End en producción: cierra archivos abiertos y pierde el estado. Prefiere Exit Sub tras liberar objetos.',
      'Évitez End en production : il ferme les fichiers ouverts et perd l’état. Préférez Exit Sub après avoir libéré les objets.',
      'End in Produktion vermeiden: schließt offene Dateien und verwirft den Zustand. Lieber Exit Sub nach Freigabe der Objekte.',
    ),
    related: ['exit', 'static', 'debug'],
  },
  {
    id: 'sub',
    name: 'Sub',
    syntax: '[Public | Private] Sub Nome([argomenti])\n    ...\nEnd Sub',
    scope: ['core', 'excel', 'access'],
    category: 'procedures',
    subcategory: 'define',
    description: L(
      'Procedura che esegue azioni e non restituisce un valore (a differenza di Function).',
      'Procedure that performs actions and does not return a value (unlike Function).',
      'Procedimiento que ejecuta acciones y no devuelve un valor (a diferencia de Function).',
      'Procédure qui exécute des actions et ne renvoie pas de valeur (contrairement à Function).',
      'Prozedur, die Aktionen ausführt und keinen Wert zurückgibt (im Gegensatz zu Function).',
    ),
    examples: [
      {
        title: L('Sub con un argomento', 'Sub with one argument', 'Sub con un argumento', 'Sub avec un argument', 'Sub mit einem Argument'),
        code: `Sub Saluta(ByVal chi As String)
    Debug.Print "ciao " & chi
End Sub

Sub DemoSub()
    Saluta "Ada"
End Sub`,
        result: L('Finestra Immediata: ciao Ada', 'Immediate Window: ciao Ada', 'Ventana Inmediato: ciao Ada', 'Fenêtre Exécution : ciao Ada', 'Direktbereich: ciao Ada'),
      },
    ],
    related: ['function', 'call', 'byval-byref', 'exit'],
  },
  {
    id: 'function',
    name: 'Function',
    syntax: '[Public | Private] Function Nome([argomenti]) [As tipo]\n    Nome = valore\nEnd Function',
    scope: ['core', 'excel', 'access'],
    category: 'procedures',
    subcategory: 'define',
    description: L(
      'Procedura che restituisce un valore assegnando il risultato al nome della Function.',
      'Procedure that returns a value by assigning the result to the Function name.',
      'Procedimiento que devuelve un valor asignando el resultado al nombre de la Function.',
      'Procédure qui renvoie une valeur en l’affectant au nom de la Function.',
      'Prozedur, die einen Wert zurückgibt, indem sie ihn dem Function-Namen zuweist.',
    ),
    examples: [
      {
        title: L('Area di un rettangolo', 'Rectangle area', 'Área de un rectángulo', 'Aire d’un rectangle', 'Rechteckfläche'),
        code: `Function Area(ByVal w As Double, ByVal h As Double) As Double
    Area = w * h
End Function

Sub DemoFunction()
    Debug.Print Area(3, 4)
End Sub`,
        result: L('Finestra Immediata: 12', 'Immediate Window: 12', 'Ventana Inmediato: 12', 'Fenêtre Exécution : 12', 'Direktbereich: 12'),
      },
    ],
    related: ['sub', 'property', 'exit', 'byval-byref'],
  },
  {
    id: 'property',
    name: 'Property Get / Let',
    syntax: 'Property Get Nome() As tipo\nProperty Let Nome(ByVal valore As tipo)',
    scope: ['core', 'excel', 'access'],
    category: 'procedures',
    subcategory: 'define',
    description: L(
      'Espone un campo come proprietà: Get legge, Let assegna. Funziona anche in un modulo standard.',
      'Exposes a field as a property: Get reads, Let assigns. Also works in a standard module.',
      'Expone un campo como propiedad: Get lee, Let asigna. También funciona en un módulo estándar.',
      'Expose un champ comme propriété : Get lit, Let affecte. Fonctionne aussi dans un module standard.',
      'Gibt ein Feld als Eigenschaft frei: Get liest, Let weist zu. Funktioniert auch in einem Standardmodul.',
    ),
    examples: [
      {
        title: L('Wrapper su una variabile di modulo', 'Wrapper around a module variable', 'Contenedor de una variable de módulo', 'Enveloppe d’une variable de module', 'Hülle um eine Modulvariable'),
        code: `Private mNome As String

Public Property Get Nome() As String
    Nome = mNome
End Property

Public Property Let Nome(ByVal value As String)
    mNome = Trim$(value)
End Property

Sub DemoProperty()
    Nome = "  Ada  "
    Debug.Print "'" & Nome & "'"
End Sub`,
        result: L('Finestra Immediata: \'Ada\'', 'Immediate Window: \'Ada\'', 'Ventana Inmediato: \'Ada\'', 'Fenêtre Exécution : \'Ada\'', 'Direktbereich: \'Ada\''),
      },
    ],
    notes: L(
      'Property Set si usa per assegnare oggetti (Set x = ...). Let è per i tipi valore e le stringhe.',
      'Property Set is used to assign objects (Set x = ...). Let is for value types and strings.',
      'Property Set se usa para asignar objetos (Set x = ...). Let es para tipos valor y cadenas.',
      'Property Set sert à affecter des objets (Set x = ...). Let est pour les types valeur et les chaînes.',
      'Property Set weist Objekte zu (Set x = ...). Let ist für Werttypen und Zeichenfolgen.',
    ),
    related: ['function', 'sub', 'set-new-nothing', 'byval-byref'],
  },
  {
    id: 'call',
    name: 'Call',
    syntax: 'Call Nome(arg1, arg2)  \' oppure Nome arg1, arg2',
    scope: ['core', 'excel', 'access'],
    category: 'procedures',
    subcategory: 'define',
    description: L(
      'Invoca una Sub o Function ignorando il valore di ritorno. Con Call le parentesi sono obbligatorie.',
      'Invokes a Sub or Function and discards the return value. Parentheses are required with Call.',
      'Invoca un Sub o Function e ignora el valor devuelto. Con Call los paréntesis son obligatorios.',
      'Invoque un Sub ou une Function et ignore la valeur de retour. Avec Call les parenthèses sont obligatoires.',
      'Ruft eine Sub oder Function auf und verwirft den Rückgabewert. Mit Call sind Klammern Pflicht.',
    ),
    examples: [
      {
        title: L('Due stili di chiamata', 'Two calling styles', 'Dos estilos de llamada', 'Deux styles d’appel', 'Zwei Aufrufstile'),
        code: `Sub Mostra(ByVal a As Long, ByVal b As Long)
    Debug.Print a + b
End Sub

Sub DemoCall()
    Call Mostra(2, 3)
    Mostra 4, 5
End Sub`,
        result: L('Finestra Immediata: 5 poi 9', 'Immediate Window: 5 then 9', 'Ventana Inmediato: 5 luego 9', 'Fenêtre Exécution : 5 puis 9', 'Direktbereich: 5 dann 9'),
      },
    ],
    related: ['sub', 'function', 'byval-byref'],
  },
  {
    id: 'byval-byref',
    name: 'ByVal / ByRef',
    syntax: 'Sub Nome(ByVal x As tipo, ByRef y As tipo)',
    scope: ['core', 'excel', 'access'],
    category: 'procedures',
    subcategory: 'args',
    description: L(
      'ByVal passa una copia; ByRef passa il riferimento (default) e la procedura può modificare la variabile del chiamante.',
      'ByVal passes a copy; ByRef passes the reference (default) so the procedure can change the caller’s variable.',
      'ByVal pasa una copia; ByRef pasa la referencia (predeterminado) y el procedimiento puede cambiar la variable del llamador.',
      'ByVal passe une copie ; ByRef passe la référence (défaut) et la procédure peut modifier la variable de l’appelant.',
      'ByVal übergibt eine Kopie; ByRef übergibt die Referenz (Standard) und die Prozedur kann die Variable des Aufrufers ändern.',
    ),
    examples: [
      {
        title: L('Solo ByRef modifica il chiamante', 'Only ByRef changes the caller', 'Solo ByRef cambia al llamador', 'Seul ByRef modifie l’appelant', 'Nur ByRef ändert den Aufrufer'),
        code: `Sub Prova(ByVal copia As Long, ByRef rif As Long)
    copia = 99
    rif = 99
End Sub

Sub DemoByValByRef()
    Dim a As Long, b As Long
    a = 1: b = 1
    Prova a, b
    Debug.Print a, b
End Sub`,
        result: L('Finestra Immediata: 1    99', 'Immediate Window: 1    99', 'Ventana Inmediato: 1    99', 'Fenêtre Exécution : 1    99', 'Direktbereich: 1    99'),
      },
    ],
    notes: L(
      'Gli oggetti sono sempre riferimenti: ByVal impedisce di sostituire il puntatore, non di mutare l’oggetto.',
      'Objects are always references: ByVal stops you replacing the pointer, not mutating the object.',
      'Los objetos siempre son referencias: ByVal impide sustituir el puntero, no mutar el objeto.',
      'Les objets sont toujours des références : ByVal empêche de remplacer le pointeur, pas de muter l’objet.',
      'Objekte sind immer Referenzen: ByVal verhindert das Ersetzen des Zeigers, nicht die Mutation des Objekts.',
    ),
    related: ['optional', 'paramarray', 'sub', 'function'],
  },
  {
    id: 'optional',
    name: 'Optional',
    syntax: 'Function Nome(Optional ByVal x As tipo = default) As tipo',
    scope: ['core', 'excel', 'access'],
    category: 'procedures',
    subcategory: 'args',
    description: L(
      'Argomento omissibile. Dopo un Optional tutti i successivi devono esserlo. Il default si scrive con =.',
      'Omittable argument. After an Optional, every following argument must be Optional too. Default is written with =.',
      'Argumento omissible. Tras un Optional, todos los siguientes deben serlo. El valor predeterminado se escribe con =.',
      'Argument omissible. Après un Optional, tous les suivants doivent l’être. La valeur par défaut s’écrit avec =.',
      'Weglassbares Argument. Nach Optional müssen alle folgenden ebenfalls Optional sein. Standardwert mit =.',
    ),
    examples: [
      {
        title: L('Saluto con default', 'Greeting with a default', 'Saludo con valor predeterminado', 'Salut avec valeur par défaut', 'Gruß mit Standardwert'),
        code: `Function Ciao(Optional ByVal chi As String = "mondo") As String
    Ciao = "ciao " & chi
End Function

Sub DemoOptional()
    Debug.Print Ciao()
    Debug.Print Ciao("Ada")
End Sub`,
        result: L('ciao mondo poi ciao Ada', 'ciao mondo then ciao Ada', 'ciao mondo luego ciao Ada', 'ciao mondo puis ciao Ada', 'ciao mondo dann ciao Ada'),
      },
    ],
    related: ['is-functions', 'byval-byref', 'paramarray'],
  },
  {
    id: 'paramarray',
    name: 'ParamArray',
    syntax: 'Function Nome(ParamArray args() As Variant) As tipo',
    scope: ['core', 'excel', 'access'],
    category: 'procedures',
    subcategory: 'args',
    description: L(
      'Accetta un numero variabile di argomenti in un array Variant. Deve essere l’ultimo parametro e non può essere ByVal/Optional.',
      'Accepts a variable number of arguments in a Variant array. Must be the last parameter and cannot be ByVal/Optional.',
      'Acepta un número variable de argumentos en una matriz Variant. Debe ser el último parámetro y no puede ser ByVal/Optional.',
      'Accepte un nombre variable d’arguments dans un tableau Variant. Doit être le dernier paramètre et ne peut pas être ByVal/Optional.',
      'Nimmt eine variable Anzahl Argumente in einem Variant-Array. Muss letzter Parameter sein und darf nicht ByVal/Optional sein.',
    ),
    examples: [
      {
        title: L('Somma di N numeri', 'Sum of N numbers', 'Suma de N números', 'Somme de N nombres', 'Summe von N Zahlen'),
        code: `Function Somma(ParamArray nums() As Variant) As Double
    Dim v As Variant, t As Double
    For Each v In nums
        t = t + v
    Next v
    Somma = t
End Function

Sub DemoParamArray()
    Debug.Print Somma(1, 2, 3, 4)
End Sub`,
        result: L('Finestra Immediata: 10', 'Immediate Window: 10', 'Ventana Inmediato: 10', 'Fenêtre Exécution : 10', 'Direktbereich: 10'),
      },
    ],
    related: ['optional', 'array-bounds', 'for-each'],
  },
  {
    id: 'on-error',
    name: 'On Error',
    syntax: 'On Error GoTo etichetta | On Error Resume Next | On Error GoTo 0',
    scope: ['core', 'excel', 'access'],
    category: 'errors',
    subcategory: 'trapping',
    description: L(
      'GoTo instrada l’errore a un’etichetta; Resume Next lo ignora e prosegue; GoTo 0 ripristina il gestore predefinito.',
      'GoTo routes the error to a label; Resume Next skips it and continues; GoTo 0 restores the default handler.',
      'GoTo envía el error a una etiqueta; Resume Next lo ignora y sigue; GoTo 0 restaura el controlador predeterminado.',
      'GoTo envoie l’erreur vers une étiquette ; Resume Next l’ignore et continue ; GoTo 0 rétablit le gestionnaire par défaut.',
      'GoTo leitet den Fehler an eine Marke; Resume Next überspringt ihn und macht weiter; GoTo 0 stellt den Standardhandler wieder her.',
    ),
    examples: [
      {
        title: L('Divisione protetta', 'Guarded division', 'División protegida', 'Division protégée', 'Abgesicherte Division'),
        code: `Sub DemoOnError()
    Dim a As Long, b As Long
    a = 10: b = 0
    On Error Resume Next
    Debug.Print a / b
    Debug.Print Err.Number
    On Error GoTo 0
End Sub`,
        result: L('Nessun crash; Err.Number = 11 (divisione per zero).', 'No crash; Err.Number = 11 (division by zero).', 'Sin bloqueo; Err.Number = 11 (división por cero).', 'Pas de plantage ; Err.Number = 11 (division par zéro).', 'Kein Absturz; Err.Number = 11 (Division durch Null).'),
      },
    ],
    related: ['err', 'resume', 'goto'],
  },
  {
    id: 'err',
    name: 'Err',
    syntax: 'Err.Number | Err.Description | Err.Raise number, [source], [desc] | Err.Clear',
    scope: ['core', 'excel', 'access'],
    category: 'errors',
    subcategory: 'trapping',
    description: L(
      'Oggetto globale dell’ultimo errore. Number e Description lo leggono; Raise lo genera; Clear lo azzera.',
      'Global object for the last error. Number and Description read it; Raise throws it; Clear resets it.',
      'Objeto global del último error. Number y Description lo leen; Raise lo genera; Clear lo borra.',
      'Objet global de la dernière erreur. Number et Description la lisent ; Raise la lève ; Clear la remet à zéro.',
      'Globales Objekt des letzten Fehlers. Number und Description lesen ihn; Raise löst ihn aus; Clear setzt ihn zurück.',
    ),
    params: [
      { name: 'number', description: L('Codice errore (vbObjectError + n per errori custom).', 'Error code (vbObjectError + n for custom errors).', 'Código de error (vbObjectError + n para errores propios).', 'Code d’erreur (vbObjectError + n pour les erreurs perso).', 'Fehlercode (vbObjectError + n für eigene Fehler).') },
    ],
    examples: [
      {
        title: L('Errore applicativo e reset', 'Application error and reset', 'Error de aplicación y reinicio', 'Erreur applicative et remise à zéro', 'Anwendungsfehler und Zurücksetzen'),
        code: `Sub DemoErr()
    On Error Resume Next
    Err.Raise 513, "DemoErr", "importo negativo"
    Debug.Print Err.Number, Err.Description
    Err.Clear
    Debug.Print Err.Number
    On Error GoTo 0
End Sub`,
        result: L('513    importo negativo poi 0', '513    importo negativo then 0', '513    importo negativo luego 0', '513    importo negativo puis 0', '513    importo negativo dann 0'),
      },
    ],
    related: ['on-error', 'resume'],
  },
  {
    id: 'resume',
    name: 'Resume / Resume Next',
    syntax: 'Resume | Resume Next | Resume etichetta',
    scope: ['core', 'excel', 'access'],
    category: 'errors',
    subcategory: 'trapping',
    description: L(
      'Dal gestore: Resume ritenta l’istruzione che ha fallito; Resume Next passa alla successiva; Resume etichetta salta lì.',
      'From the handler: Resume retries the failed statement; Resume Next goes to the next one; Resume label jumps there.',
      'Desde el controlador: Resume reintenta la instrucción fallida; Resume Next pasa a la siguiente; Resume etiqueta salta allí.',
      'Depuis le gestionnaire : Resume réessaie l’instruction en échec ; Resume Next passe à la suivante ; Resume étiquette y saute.',
      'Aus dem Handler: Resume wiederholt die fehlgeschlagene Anweisung; Resume Next geht zur nächsten; Resume Marke springt dorthin.',
    ),
    examples: [
      {
        title: L('Riprova dopo aver corretto i dati', 'Retry after fixing the data', 'Reintentar tras corregir los datos', 'Réessayer après correction des données', 'Erneut versuchen nach Datenkorrektur'),
        code: `Sub DemoResume()
    Dim a As Double, b As Double
    a = 10: b = 0
    On Error GoTo Handler
    Debug.Print a / b
    Exit Sub
Handler:
    b = 2
    Resume
End Sub`,
        result: L('Finestra Immediata: 5 (dopo Resume b vale 2).', 'Immediate Window: 5 (after Resume, b is 2).', 'Ventana Inmediato: 5 (tras Resume, b vale 2).', 'Fenêtre Exécution : 5 (après Resume, b vaut 2).', 'Direktbereich: 5 (nach Resume ist b 2).'),
      },
    ],
    related: ['on-error', 'err', 'goto'],
  },
  {
    id: 'msgbox',
    name: 'MsgBox',
    syntax: 'MsgBox(prompt, [buttons], [title]) As VbMsgBoxResult',
    scope: ['core', 'excel', 'access'],
    category: 'interaction',
    subcategory: 'dialogs',
    description: L(
      'Mostra una finestra di dialogo e restituisce il pulsante premuto (vbOK, vbYes, vbNo, …).',
      'Shows a dialog and returns the button pressed (vbOK, vbYes, vbNo, …).',
      'Muestra un cuadro de diálogo y devuelve el botón pulsado (vbOK, vbYes, vbNo, …).',
      'Affiche une boîte de dialogue et renvoie le bouton cliqué (vbOK, vbYes, vbNo, …).',
      'Zeigt einen Dialog und gibt die gedrückte Schaltfläche zurück (vbOK, vbYes, vbNo, …).',
    ),
    params: [
      { name: 'prompt', description: L('Testo del messaggio.', 'Message text.', 'Texto del mensaje.', 'Texte du message.', 'Meldungstext.') },
      { name: 'buttons', optional: true, description: L('Combinazione vbYesNo, vbQuestion, vbDefaultButton2, …', 'Combination of vbYesNo, vbQuestion, vbDefaultButton2, …', 'Combinación vbYesNo, vbQuestion, vbDefaultButton2, …', 'Combinaison vbYesNo, vbQuestion, vbDefaultButton2, …', 'Kombination vbYesNo, vbQuestion, vbDefaultButton2, …') },
      { name: 'title', optional: true, description: L('Titolo della finestra.', 'Dialog title.', 'Título de la ventana.', 'Titre de la fenêtre.', 'Fenstertitel.') },
    ],
    examples: [
      {
        title: L('Conferma Sì/No', 'Yes/No confirm', 'Confirmación Sí/No', 'Confirmation Oui/Non', 'Ja/Nein-Bestätigung'),
        code: `Sub DemoMsgBox()
    Dim r As VbMsgBoxResult
    r = MsgBox("Salvare?", vbYesNo + vbQuestion, "Conferma")
    If r = vbYes Then Debug.Print "yes" Else Debug.Print "no"
End Sub`,
        result: L('Compare il dialogo; Immediata stampa yes o no in base al pulsante.', 'Dialog appears; Immediate prints yes or no from the button.', 'Aparece el diálogo; Inmediato imprime yes o no según el botón.', 'La boîte s’affiche ; Exécution imprime yes ou no selon le bouton.', 'Dialog erscheint; Direktbereich druckt yes oder no je nach Schaltfläche.'),
      },
    ],
    related: ['inputbox', 'beep'],
  },
  {
    id: 'inputbox',
    name: 'InputBox',
    syntax: 'InputBox(prompt, [title], [default], [xpos], [ypos]) As String',
    scope: ['core', 'excel', 'access'],
    category: 'interaction',
    subcategory: 'dialogs',
    description: L(
      'Chiede una riga di testo. Annulla restituisce stringa vuota (non si distingue da OK su campo vuoto).',
      'Asks for a line of text. Cancel returns an empty string (not distinguishable from OK on an empty field).',
      'Pide una línea de texto. Cancelar devuelve cadena vacía (no se distingue de Aceptar con campo vacío).',
      'Demande une ligne de texte. Annuler renvoie une chaîne vide (indistinct d’OK sur un champ vide).',
      'Fragt eine Textzeile ab. Abbrechen liefert eine leere Zeichenfolge (nicht von OK bei leerem Feld unterscheidbar).',
    ),
    params: [
      { name: 'prompt', description: L('Testo della richiesta.', 'Prompt text.', 'Texto de la solicitud.', 'Texte de la demande.', 'Aufforderungstext.') },
      { name: 'default', optional: true, description: L('Valore iniziale nel campo.', 'Initial value in the field.', 'Valor inicial en el campo.', 'Valeur initiale dans le champ.', 'Anfangswert im Feld.') },
    ],
    examples: [
      {
        title: L('Nome con default', 'Name with default', 'Nombre con valor predeterminado', 'Nom avec valeur par défaut', 'Name mit Standardwert'),
        code: `Sub DemoInputBox()
    Dim nome As String
    nome = InputBox("Come ti chiami?", "Profilo", "Ada")
    If Len(nome) = 0 Then
        Debug.Print "annullato"
    Else
        Debug.Print "ciao " & nome
    End If
End Sub`,
        result: L('Se confermi Ada: ciao Ada. Se Annulla: annullato.', 'If you confirm Ada: ciao Ada. If Cancel: annullato.', 'Si confirmas Ada: ciao Ada. Si Cancelar: annullato.', 'Si vous confirmez Ada : ciao Ada. Si Annuler : annullato.', 'Bei Bestätigung Ada: ciao Ada. Bei Abbrechen: annullato.'),
      },
    ],
    related: ['msgbox', 'c-other', 'len'],
  },
  {
    id: 'beep',
    name: 'Beep',
    syntax: 'Beep',
    scope: ['core', 'excel', 'access'],
    category: 'interaction',
    subcategory: 'dialogs',
    description: L(
      'Emette il suono di sistema predefinito. Non accetta argomenti e non restituisce un valore.',
      'Plays the default system sound. Takes no arguments and returns no value.',
      'Emite el sonido predeterminado del sistema. No acepta argumentos ni devuelve valor.',
      'Émet le son système par défaut. Ne prend aucun argument et ne renvoie rien.',
      'Gibt den Standard-Systemton aus. Nimmt keine Argumente und liefert keinen Wert.',
    ),
    examples: [
      {
        title: L('Avviso sonoro su errore', 'Sound alert on error', 'Aviso sonoro ante error', 'Alerte sonore en cas d’erreur', 'Tonwarnung bei Fehler'),
        code: `Sub DemoBeep()
    If 2 + 2 <> 5 Then
        Beep
        Debug.Print "beep ok"
    End If
End Sub`,
        result: L('Si sente il beep di sistema e Immediata stampa beep ok.', 'The system beep plays and Immediate prints beep ok.', 'Se oye el beep del sistema e Inmediato imprime beep ok.', 'Le bip système retentit et Exécution imprime beep ok.', 'Der Systemton erklingt und der Direktbereich druckt beep ok.'),
      },
    ],
    related: ['msgbox', 'err'],
  },
  {
    id: 'doevents',
    name: 'DoEvents',
    syntax: 'DoEvents',
    scope: ['core', 'excel', 'access'],
    category: 'interaction',
    subcategory: 'runtime',
    description: L(
      'Cede il controllo a Windows: ridisegna le finestre e accetta clic/tasti durante un ciclo lungo.',
      'Yields control to Windows: repaints windows and accepts clicks/keys during a long loop.',
      'Cede el control a Windows: redibuja ventanas y acepta clics/teclas durante un bucle largo.',
      'Rend la main à Windows : redessine les fenêtres et accepte clics/touches pendant une longue boucle.',
      'Gibt die Kontrolle an Windows: zeichnet Fenster neu und nimmt Klicks/Tasten in einer langen Schleife an.',
    ),
    examples: [
      {
        title: L('Ciclo che resta reattivo', 'Loop that stays responsive', 'Bucle que sigue respondiendo', 'Boucle qui reste réactive', 'Schleife bleibt bedienbar'),
        code: `Sub DemoDoEvents()
    Dim i As Long
    For i = 1 To 3
        Debug.Print i
        DoEvents
    Next i
End Sub`,
        result: L('Stampa 1, 2, 3; tra un giro e l’altro l’UI può aggiornarsi.', 'Prints 1, 2, 3; the UI can refresh between iterations.', 'Imprime 1, 2, 3; la IU puede actualizarse entre iteraciones.', 'Affiche 1, 2, 3 ; l’UI peut se rafraîchir entre les itérations.', 'Gibt 1, 2, 3 aus; die UI kann zwischen den Durchläufen aktualisieren.'),
      },
    ],
    notes: L(
      'DoEvents può far rieseguire eventi (reentrancy). Non usarlo come timer di precisione.',
      'DoEvents can re-enter events (reentrancy). Do not use it as a precision timer.',
      'DoEvents puede reentrar en eventos (reentrancy). No lo uses como temporizador preciso.',
      'DoEvents peut réentrer dans des événements (réentrance). Ne l’utilisez pas comme minuteur précis.',
      'DoEvents kann Ereignisse erneut betreten (Reentrancy). Nicht als Präzisionstimer nutzen.',
    ),
    related: ['for-next', 'timer', 'debug'],
  },
  {
    id: 'shell',
    name: 'Shell',
    syntax: 'Shell(pathname, [windowstyle]) As Double',
    scope: ['core', 'excel', 'access'],
    category: 'interaction',
    subcategory: 'runtime',
    description: L(
      'Avvia un programma esterno e restituisce l’ID di processo, oppure 0 se fallisce (con On Error).',
      'Starts an external program and returns the process ID, or 0 on failure (with On Error).',
      'Inicia un programa externo y devuelve el ID de proceso, o 0 si falla (con On Error).',
      'Lance un programme externe et renvoie l’ID de processus, ou 0 en cas d’échec (avec On Error).',
      'Startet ein externes Programm und gibt die Prozess-ID zurück, oder 0 bei Fehler (mit On Error).',
    ),
    params: [
      { name: 'pathname', description: L('Comando o percorso eseguibile.', 'Command or executable path.', 'Comando o ruta del ejecutable.', 'Commande ou chemin de l’exécutable.', 'Befehl oder Pfad der ausführbaren Datei.') },
      { name: 'windowstyle', optional: true, description: L('vbNormalFocus, vbHide, vbMinimizedNoFocus, …', 'vbNormalFocus, vbHide, vbMinimizedNoFocus, …', 'vbNormalFocus, vbHide, vbMinimizedNoFocus, …', 'vbNormalFocus, vbHide, vbMinimizedNoFocus, …', 'vbNormalFocus, vbHide, vbMinimizedNoFocus, …') },
    ],
    examples: [
      {
        title: L('Apri il Blocco note', 'Open Notepad', 'Abrir el Bloc de notas', 'Ouvrir le Bloc-notes', 'Editor öffnen'),
        code: `Sub DemoShell()
    Dim pid As Double
    pid = Shell("notepad.exe", vbNormalFocus)
    Debug.Print pid
End Sub`,
        result: L('Si apre Notepad; Immediata stampa un PID > 0 (es. 12345).', 'Notepad opens; Immediate prints a PID > 0 (e.g. 12345).', 'Se abre el Bloc de notas; Inmediato imprime un PID > 0 (p. ej. 12345).', 'Le Bloc-notes s’ouvre ; Exécution imprime un PID > 0 (ex. 12345).', 'Editor öffnet sich; Direktbereich druckt eine PID > 0 (z. B. 12345).'),
      },
    ],
    related: ['environ', 'doevents'],
  },
  {
    id: 'environ',
    name: 'Environ',
    syntax: 'Environ(envstring | number) As String',
    scope: ['core', 'excel', 'access'],
    category: 'interaction',
    subcategory: 'runtime',
    description: L(
      'Legge una variabile d’ambiente (USERNAME, TEMP, PATH, …) o la n-esima voce se passi un numero.',
      'Reads an environment variable (USERNAME, TEMP, PATH, …) or the n-th entry if you pass a number.',
      'Lee una variable de entorno (USERNAME, TEMP, PATH, …) o la n-ésima entrada si pasas un número.',
      'Lit une variable d’environnement (USERNAME, TEMP, PATH, …) ou la n-ième entrée si vous passez un numéro.',
      'Liest eine Umgebungsvariable (USERNAME, TEMP, PATH, …) oder den n-ten Eintrag bei einer Nummer.',
    ),
    examples: [
      {
        title: L('Cartella temporanea', 'Temp folder', 'Carpeta temporal', 'Dossier temporaire', 'Temp-Ordner'),
        code: `Sub DemoEnviron()
    Debug.Print Environ("TEMP")
    Debug.Print Len(Environ("USERNAME")) > 0
End Sub`,
        result: L('Percorso tipo C:\\Users\\...\\AppData\\Local\\Temp e True.', 'A path like C:\\Users\\...\\AppData\\Local\\Temp and True.', 'Una ruta tipo C:\\Users\\...\\AppData\\Local\\Temp y True.', 'Un chemin du type C:\\Users\\...\\AppData\\Local\\Temp et True.', 'Ein Pfad wie C:\\Users\\...\\AppData\\Local\\Temp und True.'),
      },
    ],
    related: ['dir', 'open-close', 'shell'],
  },
  {
    id: 'createobject-getobject',
    name: 'CreateObject / GetObject',
    syntax: 'CreateObject(class, [server]) | GetObject([pathname], [class])',
    scope: ['core', 'excel', 'access'],
    category: 'interaction',
    subcategory: 'runtime',
    description: L(
      'CreateObject istanzia un server COM (late binding). GetObject si aggancia a un’istanza già aperta o a un file.',
      'CreateObject instantiates a COM server (late binding). GetObject attaches to an already running instance or a file.',
      'CreateObject instancia un servidor COM (late binding). GetObject se engancha a una instancia ya abierta o a un archivo.',
      'CreateObject instancie un serveur COM (liaison tardive). GetObject s’attache à une instance déjà ouverte ou à un fichier.',
      'CreateObject erzeugt einen COM-Server (Late Binding). GetObject hängt sich an eine laufende Instanz oder eine Datei.',
    ),
    params: [
      { name: 'class', description: L('ProgID, es. Scripting.FileSystemObject.', 'ProgID, e.g. Scripting.FileSystemObject.', 'ProgID, p. ej. Scripting.FileSystemObject.', 'ProgID, ex. Scripting.FileSystemObject.', 'ProgID, z. B. Scripting.FileSystemObject.') },
    ],
    examples: [
      {
        title: L('FileSystemObject e cartella TEMP', 'FileSystemObject and TEMP folder', 'FileSystemObject y carpeta TEMP', 'FileSystemObject et dossier TEMP', 'FileSystemObject und TEMP-Ordner'),
        code: `Sub DemoCreateGetObject()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Debug.Print fso.FolderExists(Environ("TEMP"))
    Set fso = Nothing
End Sub`,
        result: L('Finestra Immediata: True (TEMP esiste).', 'Immediate Window: True (TEMP exists).', 'Ventana Inmediato: True (TEMP existe).', 'Fenêtre Exécution : True (TEMP existe).', 'Direktbereich: True (TEMP existiert).'),
      },
    ],
    notes: L(
      'GetObject(, "Excel.Application") richiede che Excel sia già in esecuzione, altrimenti errore 429.',
      'GetObject(, "Excel.Application") requires Excel already running, otherwise error 429.',
      'GetObject(, "Excel.Application") exige que Excel ya esté en ejecución; si no, error 429.',
      'GetObject(, "Excel.Application") exige qu’Excel soit déjà lancé, sinon erreur 429.',
      'GetObject(, "Excel.Application") setzt eine laufende Excel-Instanz voraus, sonst Fehler 429.',
    ),
    related: ['object', 'set-new-nothing', 'environ'],
  },
  {
    id: 'debug',
    name: 'Debug.Print / Debug.Assert',
    syntax: 'Debug.Print [expr[, expr2]...] | Debug.Assert condizione',
    scope: ['core', 'excel', 'access'],
    category: 'interaction',
    subcategory: 'runtime',
    description: L(
      'Print scrive nella Finestra Immediata. Assert interrompe se la condizione è False (solo in IDE, viene compilato via in EXE).',
      'Print writes to the Immediate Window. Assert breaks when the condition is False (IDE only; compiled out of EXE).',
      'Print escribe en la ventana Inmediato. Assert interrumpe si la condición es False (solo IDE; se omite en el EXE).',
      'Print écrit dans la fenêtre Exécution. Assert s’arrête si la condition est False (IDE seulement ; retiré de l’EXE).',
      'Print schreibt in den Direktbereich. Assert hält an, wenn die Bedingung False ist (nur IDE; im EXE entfernt).',
    ),
    examples: [
      {
        title: L('Traccia e asserzione', 'Trace and assertion', 'Traza y aserción', 'Trace et assertion', 'Trace und Assertion'),
        code: `Sub DemoDebug()
    Dim n As Long: n = 4
    Debug.Print "n=", n
    Debug.Assert n > 0
    Debug.Print "ok"
End Sub`,
        result: L('Immediata: n= 4 poi ok. Se n fosse 0, il debugger si fermerebbe su Assert.', 'Immediate: n= 4 then ok. If n were 0, the debugger would stop on Assert.', 'Inmediato: n= 4 luego ok. Si n fuera 0, el depurador se detendría en Assert.', 'Exécution : n= 4 puis ok. Si n valait 0, le débogueur s’arrêterait sur Assert.', 'Direktbereich: n= 4 dann ok. Wäre n 0, würde der Debugger bei Assert anhalten.'),
      },
    ],
    related: ['end-stop', 'msgbox'],
  },
  {
    id: 'left-right-mid',
    name: 'Left / Right / Mid',
    syntax: 'Left(string, length) | Right(string, length) | Mid(string, start, [length])',
    scope: ['core', 'excel', 'access'],
    category: 'fn-string',
    subcategory: 'extract',
    description: L(
      'Estrae caratteri da sinistra, da destra o da una posizione 1-based (Mid). Senza length, Mid arriva in fondo.',
      'Extracts characters from the left, right, or a 1-based position (Mid). Without length, Mid goes to the end.',
      'Extrae caracteres desde la izquierda, la derecha o una posición 1-based (Mid). Sin length, Mid llega al final.',
      'Extrait des caractères à gauche, à droite ou depuis une position 1-based (Mid). Sans length, Mid va jusqu’à la fin.',
      'Extrahiert Zeichen von links, rechts oder ab einer 1-basierten Position (Mid). Ohne length geht Mid bis zum Ende.',
    ),
    params: [
      { name: 'string', description: L('Testo di partenza.', 'Source text.', 'Texto de origen.', 'Texte source.', 'Quelltext.') },
      { name: 'start', description: L('Posizione iniziale 1-based (solo Mid).', '1-based start position (Mid only).', 'Posición inicial 1-based (solo Mid).', 'Position de départ 1-based (Mid uniquement).', '1-basierte Startposition (nur Mid).') },
      { name: 'length', optional: true, description: L('Numero di caratteri.', 'Number of characters.', 'Número de caracteres.', 'Nombre de caractères.', 'Anzahl der Zeichen.') },
    ],
    examples: [
      {
        title: L('Pezzi di "Visual"', 'Pieces of "Visual"', 'Trozos de "Visual"', 'Morceaux de « Visual »', 'Teile von „Visual“'),
        code: `Sub DemoLeftRightMid()
    Debug.Print Left("Visual", 3)
    Debug.Print Right("Visual", 3)
    Debug.Print Mid("Visual", 2, 3)
End Sub`,
        result: L('Vis / ual / isu', 'Vis / ual / isu', 'Vis / ual / isu', 'Vis / ual / isu', 'Vis / ual / isu'),
      },
    ],
    related: ['len', 'instr', 'string'],
  },
  {
    id: 'len',
    name: 'Len',
    syntax: 'Len(string) As Long',
    scope: ['core', 'excel', 'access'],
    category: 'fn-string',
    subcategory: 'extract',
    description: L(
      'Numero di caratteri di una stringa. Su un UDT restituisce i byte; su Object non è valido.',
      'Character count of a string. On a UDT it returns bytes; it is not valid on Object.',
      'Número de caracteres de una cadena. En un UDT devuelve bytes; no es válido en Object.',
      'Nombre de caractères d’une chaîne. Sur un UDT il renvoie des octets ; invalide sur Object.',
      'Zeichenanzahl einer Zeichenfolge. Bei einem UDT Bytes; bei Object ungültig.',
    ),
    examples: [
      {
        title: L('Lunghezza e stringa vuota', 'Length and empty string', 'Longitud y cadena vacía', 'Longueur et chaîne vide', 'Länge und leere Zeichenfolge'),
        code: `Sub DemoLen()
    Debug.Print Len("VBA")
    Debug.Print Len("")
End Sub`,
        result: L('Finestra Immediata: 3 poi 0', 'Immediate Window: 3 then 0', 'Ventana Inmediato: 3 luego 0', 'Fenêtre Exécution : 3 puis 0', 'Direktbereich: 3 dann 0'),
      },
    ],
    related: ['left-right-mid', 'string', 'trim'],
  },
  {
    id: 'instr',
    name: 'InStr / InStrRev',
    syntax: 'InStr([start], string1, string2, [compare]) | InStrRev(string1, string2, [start], [compare])',
    scope: ['core', 'excel', 'access'],
    category: 'fn-string',
    subcategory: 'extract',
    description: L(
      'Posizione 1-based della sottostringa. InStr cerca da sinistra, InStrRev da destra. 0 = non trovata.',
      '1-based position of a substring. InStr searches left-to-right, InStrRev right-to-left. 0 = not found.',
      'Posición 1-based de una subcadena. InStr busca de izquierda a derecha, InStrRev de derecha a izquierda. 0 = no hallada.',
      'Position 1-based d’une sous-chaîne. InStr cherche de gauche à droite, InStrRev de droite à gauche. 0 = introuvable.',
      '1-basierte Position einer Teilzeichenfolge. InStr sucht von links, InStrRev von rechts. 0 = nicht gefunden.',
    ),
    params: [
      { name: 'start', optional: true, description: L('Indice da cui partire (default 1 / Len per Rev).', 'Index to start from (default 1 / Len for Rev).', 'Índice de inicio (predeterminado 1 / Len para Rev).', 'Index de départ (défaut 1 / Len pour Rev).', 'Startindex (Standard 1 / Len für Rev).') },
      { name: 'compare', optional: true, description: L('vbBinaryCompare o vbTextCompare.', 'vbBinaryCompare or vbTextCompare.', 'vbBinaryCompare o vbTextCompare.', 'vbBinaryCompare ou vbTextCompare.', 'vbBinaryCompare oder vbTextCompare.') },
    ],
    examples: [
      {
        title: L('Prima e ultima "l"', 'First and last "l"', 'Primera y última "l"', 'Premier et dernier « l »', 'Erstes und letztes „l“'),
        code: `Sub DemoInStr()
    Debug.Print InStr("hello", "l")
    Debug.Print InStrRev("hello", "l")
    Debug.Print InStr("hello", "x")
End Sub`,
        result: L('Finestra Immediata: 3, 4, 0', 'Immediate Window: 3, 4, 0', 'Ventana Inmediato: 3, 4, 0', 'Fenêtre Exécution : 3, 4, 0', 'Direktbereich: 3, 4, 0'),
      },
    ],
    related: ['left-right-mid', 'replace', 'strcomp'],
  },
  {
    id: 'split',
    name: 'Split',
    syntax: 'Split(expression, [delimiter], [limit], [compare]) As String()',
    scope: ['core', 'excel', 'access'],
    category: 'fn-string',
    subcategory: 'extract',
    description: L(
      'Spezza una stringa in un array 0-based. delimiter di default è lo spazio. limit = -1 prende tutti i pezzi.',
      'Breaks a string into a 0-based array. Default delimiter is space. limit = -1 takes every piece.',
      'Parte una cadena en una matriz 0-based. El delimitador predeterminado es el espacio. limit = -1 toma todas las piezas.',
      'Découpe une chaîne en tableau 0-based. Le délimiteur par défaut est l’espace. limit = -1 prend tous les morceaux.',
      'Zerlegt eine Zeichenfolge in ein 0-basiertes Array. Standardtrenner ist Leerzeichen. limit = -1 nimmt alle Teile.',
    ),
    examples: [
      {
        title: L('CSV a tre campi', 'Three-field CSV', 'CSV de tres campos', 'CSV à trois champs', 'CSV mit drei Feldern'),
        code: `Sub DemoSplit()
    Dim p() As String
    p = Split("Ada,Lovelace,1815", ",")
    Debug.Print LBound(p), UBound(p)
    Debug.Print p(0), p(2)
End Sub`,
        result: L('0 2 poi Ada    1815', '0 2 then Ada    1815', '0 2 luego Ada    1815', '0 2 puis Ada    1815', '0 2 dann Ada    1815'),
      },
    ],
    related: ['join', 'array-bounds', 'instr'],
  },
  {
    id: 'trim',
    name: 'Trim / LTrim / RTrim',
    syntax: 'Trim(string) | LTrim(string) | RTrim(string)',
    scope: ['core', 'excel', 'access'],
    category: 'fn-string',
    subcategory: 'transform',
    description: L(
      'Toglie gli spazi: Trim da entrambi i lati, LTrim a sinistra, RTrim a destra. Non toglie tabulazioni.',
      'Strips spaces: Trim on both sides, LTrim on the left, RTrim on the right. Does not strip tabs.',
      'Quita espacios: Trim a ambos lados, LTrim a la izquierda, RTrim a la derecha. No quita tabulaciones.',
      'Retire les espaces : Trim des deux côtés, LTrim à gauche, RTrim à droite. Ne retire pas les tabulations.',
      'Entfernt Leerzeichen: Trim beidseitig, LTrim links, RTrim rechts. Entfernt keine Tabulatoren.',
    ),
    examples: [
      {
        title: L('Spazi residui', 'Leftover spaces', 'Espacios restantes', 'Espaces restants', 'Verbleibende Leerzeichen'),
        code: `Sub DemoTrim()
    Dim s As String: s = "  Ada  "
    Debug.Print "'" & Trim(s) & "'"
    Debug.Print "'" & LTrim(s) & "'"
    Debug.Print "'" & RTrim(s) & "'"
End Sub`,
        result: L('\'Ada\' / \'Ada  \' / \'  Ada\'', '\'Ada\' / \'Ada  \' / \'  Ada\'', '\'Ada\' / \'Ada  \' / \'  Ada\'', '\'Ada\' / \'Ada  \' / \'  Ada\'', '\'Ada\' / \'Ada  \' / \'  Ada\''),
      },
    ],
    related: ['replace', 'len', 'space-string'],
  },
  {
    id: 'ucase-lcase',
    name: 'UCase / LCase',
    syntax: 'UCase(string) | LCase(string)',
    scope: ['core', 'excel', 'access'],
    category: 'fn-string',
    subcategory: 'transform',
    description: L(
      'Converte l’intero testo in maiuscolo o minuscolo secondo le regole del locale di sistema.',
      'Converts the whole text to upper or lower case using the system locale rules.',
      'Convierte todo el texto a mayúsculas o minúsculas según las reglas de la configuración regional.',
      'Convertit tout le texte en majuscules ou minuscules selon les règles de la locale système.',
      'Wandelt den gesamten Text nach Systemlocale in Groß- oder Kleinbuchstaben um.',
    ),
    examples: [
      {
        title: L('Maiuscolo e minuscolo', 'Upper and lower case', 'Mayúsculas y minúsculas', 'Majuscules et minuscules', 'Groß- und Kleinschreibung'),
        code: `Sub DemoUCaseLCase()
    Debug.Print UCase("Vba")
    Debug.Print LCase("Vba")
End Sub`,
        result: L('VBA poi vba', 'VBA then vba', 'VBA luego vba', 'VBA puis vba', 'VBA dann vba'),
      },
    ],
    related: ['strconv', 'strcomp', 'replace'],
  },
  {
    id: 'replace',
    name: 'Replace',
    syntax: 'Replace(expression, find, replace, [start], [count], [compare])',
    scope: ['core', 'excel', 'access'],
    category: 'fn-string',
    subcategory: 'transform',
    description: L(
      'Sostituisce le occorrenze di find con replace. count = -1 (default) le cambia tutte.',
      'Replaces occurrences of find with replace. count = -1 (default) changes every match.',
      'Sustituye las ocurrencias de find por replace. count = -1 (predeterminado) las cambia todas.',
      'Remplace les occurrences de find par replace. count = -1 (défaut) les change toutes.',
      'Ersetzt Vorkommen von find durch replace. count = -1 (Standard) ändert alle Treffer.',
    ),
    examples: [
      {
        title: L('Due sostituzioni', 'Two replacements', 'Dos sustituciones', 'Deux remplacements', 'Zwei Ersetzungen'),
        code: `Sub DemoReplace()
    Debug.Print Replace("a-a-a", "a", "b")
    Debug.Print Replace("a-a-a", "a", "b", 1, 1)
End Sub`,
        result: L('b-b-b poi b-a-a', 'b-b-b then b-a-a', 'b-b-b luego b-a-a', 'b-b-b puis b-a-a', 'b-b-b dann b-a-a'),
      },
    ],
    related: ['instr', 'join', 'strconv'],
  },
  {
    id: 'join',
    name: 'Join',
    syntax: 'Join(sourcearray, [delimiter]) As String',
    scope: ['core', 'excel', 'access'],
    category: 'fn-string',
    subcategory: 'transform',
    description: L(
      'Unisce un array di stringhe (o Variant) con un delimitatore. Inverso di Split. delimiter default = spazio.',
      'Joins a string (or Variant) array with a delimiter. Inverse of Split. Default delimiter is space.',
      'Une una matriz de cadenas (o Variant) con un delimitador. Inverso de Split. Delimitador predeterminado: espacio.',
      'Assemble un tableau de chaînes (ou Variant) avec un délimiteur. Inverse de Split. Délimiteur par défaut : espace.',
      'Fügt ein String- (oder Variant-)Array mit einem Trenner zusammen. Umkehr von Split. Standardtrenner: Leerzeichen.',
    ),
    examples: [
      {
        title: L('Array → CSV', 'Array → CSV', 'Matriz → CSV', 'Tableau → CSV', 'Array → CSV'),
        code: `Sub DemoJoin()
    Debug.Print Join(Array("Ada", "Lovelace", "1815"), ",")
End Sub`,
        result: L('Finestra Immediata: Ada,Lovelace,1815', 'Immediate Window: Ada,Lovelace,1815', 'Ventana Inmediato: Ada,Lovelace,1815', 'Fenêtre Exécution : Ada,Lovelace,1815', 'Direktbereich: Ada,Lovelace,1815'),
      },
    ],
    related: ['split', 'array-bounds', 'replace'],
  },
  {
    id: 'strconv',
    name: 'StrConv',
    syntax: 'StrConv(string, conversion, [LCID])',
    scope: ['core', 'excel', 'access'],
    category: 'fn-string',
    subcategory: 'transform',
    description: L(
      'Conversioni di caso e codice: vbUpperCase, vbLowerCase, vbProperCase, vbUnicode, vbFromUnicode.',
      'Case and code conversions: vbUpperCase, vbLowerCase, vbProperCase, vbUnicode, vbFromUnicode.',
      'Conversiones de caja y código: vbUpperCase, vbLowerCase, vbProperCase, vbUnicode, vbFromUnicode.',
      'Conversions de casse et de code : vbUpperCase, vbLowerCase, vbProperCase, vbUnicode, vbFromUnicode.',
      'Schreibweisen- und Codewandlungen: vbUpperCase, vbLowerCase, vbProperCase, vbUnicode, vbFromUnicode.',
    ),
    examples: [
      {
        title: L('Title case', 'Title case', 'Tipo título', 'Casse de titre', 'Titelschreibweise'),
        code: `Sub DemoStrConv()
    Debug.Print StrConv("ada lovelace", vbProperCase)
End Sub`,
        result: L('Finestra Immediata: Ada Lovelace', 'Immediate Window: Ada Lovelace', 'Ventana Inmediato: Ada Lovelace', 'Fenêtre Exécution : Ada Lovelace', 'Direktbereich: Ada Lovelace'),
      },
    ],
    related: ['ucase-lcase', 'replace'],
  },
  {
    id: 'space-string',
    name: 'Space / String',
    syntax: 'Space(number) | String(number, character)',
    scope: ['core', 'excel', 'access'],
    category: 'fn-string',
    subcategory: 'transform',
    description: L(
      'Space produce n spazi. String ripete un carattere (o il primo di una stringa) n volte.',
      'Space produces n spaces. String repeats a character (or the first of a string) n times.',
      'Space produce n espacios. String repite un carácter (o el primero de una cadena) n veces.',
      'Space produit n espaces. String répète un caractère (ou le premier d’une chaîne) n fois.',
      'Space erzeugt n Leerzeichen. String wiederholt ein Zeichen (oder das erste einer Zeichenfolge) n-mal.',
    ),
    examples: [
      {
        title: L('Padding e barra', 'Padding and a bar', 'Relleno y una barra', 'Remplissage et une barre', 'Auffüllen und ein Balken'),
        code: `Sub DemoSpaceString()
    Debug.Print "'" & Space(3) & "'"
    Debug.Print String(5, "*")
    Debug.Print String(3, "AB")
End Sub`,
        result: L('\'   \' / ***** / AAA', '\'   \' / ***** / AAA', '\'   \' / ***** / AAA', '\'   \' / ***** / AAA', '\'   \' / ***** / AAA'),
      },
    ],
    related: ['trim', 'len', 'string'],
  },
  {
    id: 'format',
    name: 'Format',
    syntax: 'Format(expression, [format], [firstdayofweek], [firstweekofyear])',
    scope: ['core', 'excel', 'access'],
    category: 'fn-string',
    subcategory: 'search-format',
    description: L(
      'Formatta numeri, date e stringhe con un modello (0, #, yyyy, hh:nn, @). Il risultato dipende dal locale.',
      'Formats numbers, dates and strings with a pattern (0, #, yyyy, hh:nn, @). The result depends on the locale.',
      'Da formato a números, fechas y cadenas con un patrón (0, #, yyyy, hh:nn, @). El resultado depende de la configuración regional.',
      'Met en forme nombres, dates et chaînes avec un motif (0, #, yyyy, hh:nn, @). Le résultat dépend de la locale.',
      'Formatiert Zahlen, Daten und Zeichenfolgen mit einem Muster (0, #, yyyy, hh:nn, @). Das Ergebnis hängt vom Locale ab.',
    ),
    examples: [
      {
        title: L('Data e numero con zeri', 'Date and zero-padded number', 'Fecha y número con ceros', 'Date et nombre à zéros', 'Datum und nullaufgefüllte Zahl'),
        code: `Sub DemoFormat()
    Debug.Print Format(#3/15/2024#, "yyyy-mm-dd")
    Debug.Print Format(7, "000")
End Sub`,
        result: L('2024-03-15 poi 007', '2024-03-15 then 007', '2024-03-15 luego 007', '2024-03-15 puis 007', '2024-03-15 dann 007'),
      },
    ],
    related: ['format-helpers', 'now', 'c-other'],
  },
  {
    id: 'format-helpers',
    name: 'FormatNumber / Currency / Percent / DateTime',
    syntax: 'FormatNumber(n, [digits]) | FormatCurrency(n) | FormatPercent(n) | FormatDateTime(d, [style])',
    scope: ['core', 'excel', 'access'],
    category: 'fn-string',
    subcategory: 'search-format',
    description: L(
      'Helper di formato legati al locale: decimali, valuta, percentuale e stili data/ora (vbShortDate, vbLongTime, …).',
      'Locale-aware format helpers: decimals, currency, percent and date/time styles (vbShortDate, vbLongTime, …).',
      'Ayudas de formato según la configuración regional: decimales, moneda, porcentaje y estilos de fecha/hora.',
      'Aides de format liées à la locale : décimales, devise, pourcentage et styles date/heure.',
      'Locale-abhängige Formathelfer: Dezimalen, Währung, Prozent und Datums-/Zeitstile.',
    ),
    examples: [
      {
        title: L('Percentuale e data breve', 'Percent and short date', 'Porcentaje y fecha corta', 'Pourcentage et date courte', 'Prozent und kurzes Datum'),
        code: `Sub DemoFormatHelpers()
    Debug.Print FormatPercent(0.256, 1)
    Debug.Print FormatNumber(1234.5, 1)
    Debug.Print FormatDateTime(#3/15/2024#, vbShortDate)
End Sub`,
        result: L('25.6% (o 25,6% nel locale), 1,234.5 (o 1.234,5), data breve locale.', '25.6% (or 25,6% in the locale), 1,234.5 (or 1.234,5), locale short date.', '25.6% (o 25,6% según la configuración), 1,234.5 (o 1.234,5), fecha corta local.', '25.6 % (ou 25,6 % selon la locale), 1 234.5 (ou 1.234,5), date courte locale.', '25.6% (oder 25,6% je Locale), 1,234.5 (oder 1.234,5), kurzes Datum laut Locale.'),
      },
    ],
    related: ['format', 'currency', 'now'],
  },
  {
    id: 'strcomp',
    name: 'StrComp',
    syntax: 'StrComp(string1, string2, [compare]) As Integer',
    scope: ['core', 'excel', 'access'],
    category: 'fn-string',
    subcategory: 'search-format',
    description: L(
      'Confronta due stringhe: -1 se la prima è minore, 0 se uguali, 1 se maggiore. Null se un argomento è Null.',
      'Compares two strings: -1 if the first is smaller, 0 if equal, 1 if greater. Null if an argument is Null.',
      'Compara dos cadenas: -1 si la primera es menor, 0 si iguales, 1 si mayor. Null si un argumento es Null.',
      'Compare deux chaînes : -1 si la première est plus petite, 0 si égales, 1 si plus grande. Null si un argument est Null.',
      'Vergleicht zwei Zeichenfolgen: -1 wenn die erste kleiner, 0 wenn gleich, 1 wenn größer. Null wenn ein Argument Null ist.',
    ),
    examples: [
      {
        title: L('Binary vs Text', 'Binary vs Text', 'Binary frente a Text', 'Binary contre Text', 'Binary gegen Text'),
        code: `Sub DemoStrComp()
    Debug.Print StrComp("A", "a", vbBinaryCompare)
    Debug.Print StrComp("A", "a", vbTextCompare)
End Sub`,
        result: L('-1 (A < a in ASCII) poi 0 (stesso testo).', '-1 (A < a in ASCII) then 0 (same text).', '-1 (A < a en ASCII) luego 0 (mismo texto).', '-1 (A < a en ASCII) puis 0 (même texte).', '-1 (A < a in ASCII) dann 0 (gleicher Text).'),
      },
    ],
    related: ['option-compare', 'ucase-lcase', 'instr'],
  },
  {
    id: 'asc-chr',
    name: 'Asc / Chr',
    syntax: 'Asc(string) As Integer | Chr(charcode) As String',
    scope: ['core', 'excel', 'access'],
    category: 'fn-string',
    subcategory: 'search-format',
    description: L(
      'Asc restituisce il codice del primo carattere; Chr costruisce il carattere da un codice (es. 65 = A, 10 = lf).',
      'Asc returns the code of the first character; Chr builds the character from a code (e.g. 65 = A, 10 = lf).',
      'Asc devuelve el código del primer carácter; Chr construye el carácter desde un código (p. ej. 65 = A, 10 = lf).',
      'Asc renvoie le code du premier caractère ; Chr construit le caractère depuis un code (ex. 65 = A, 10 = lf).',
      'Asc liefert den Code des ersten Zeichens; Chr erzeugt das Zeichen aus einem Code (z. B. 65 = A, 10 = lf).',
    ),
    examples: [
      {
        title: L('A e ritorno a capo', 'A and a line feed', 'A y un salto de línea', 'A et un saut de ligne', 'A und ein Zeilenvorschub'),
        code: `Sub DemoAscChr()
    Debug.Print Asc("A")
    Debug.Print Chr(65)
    Debug.Print "r1" & Chr(10) & "r2"
End Sub`,
        result: L('65 / A / due righe r1 e r2', '65 / A / two lines r1 and r2', '65 / A / dos líneas r1 y r2', '65 / A / deux lignes r1 et r2', '65 / A / zwei Zeilen r1 und r2'),
      },
    ],
    related: ['val-str-hex', 'string'],
  },
  {
    id: 'val-str-hex',
    name: 'Val / Str / Hex',
    syntax: 'Val(string) | Str(number) | Hex(number)',
    scope: ['core', 'excel', 'access'],
    category: 'fn-string',
    subcategory: 'search-format',
    description: L(
      'Val legge un numero in formato USA (punto decimale) e si ferma al primo non valido. Str aggiunge uno spazio ai positivi. Hex è esadecimale.',
      'Val parses a US-format number (dot decimal) and stops at the first invalid char. Str adds a leading space on positives. Hex is hexadecimal.',
      'Val lee un número en formato US (punto decimal) y se detiene en el primer no válido. Str añade un espacio a los positivos. Hex es hexadecimal.',
      'Val lit un nombre au format US (point décimal) et s’arrête au premier caractère invalide. Str ajoute une espace aux positifs. Hex est hexadécimal.',
      'Val liest eine Zahl im US-Format (Punkt) und stoppt am ersten ungültigen Zeichen. Str setzt bei Positiven ein Leerzeichen davor. Hex ist hexadezimal.',
    ),
    examples: [
      {
        title: L('Parse, spazio e FF', 'Parse, space and FF', 'Análisis, espacio y FF', 'Analyse, espace et FF', 'Parse, Leerzeichen und FF'),
        code: `Sub DemoValStrHex()
    Debug.Print Val("12.5abc")
    Debug.Print "'" & Str(12) & "'"
    Debug.Print Hex(255)
End Sub`,
        result: L('12.5 / \' 12\' / FF', '12.5 / \' 12\' / FF', '12.5 / \' 12\' / FF', '12.5 / \' 12\' / FF', '12.5 / \' 12\' / FF'),
      },
    ],
    related: ['c-numeric', 'c-other', 'asc-chr'],
  },
  {
    id: 'now',
    name: 'Now / Date / Time',
    syntax: 'Now | Date | Time',
    scope: ['core', 'excel', 'access'],
    category: 'fn-date',
    subcategory: 'now',
    description: L(
      'Now è data+ora di sistema. Date e Time come funzioni restituiscono solo la parte data o ora; come istruzioni le impostano (serve permesso).',
      'Now is the system date+time. Date and Time as functions return just the date or time part; as statements they set the clock (needs permission).',
      'Now es fecha+hora del sistema. Date y Time como funciones devuelven solo la parte fecha u hora; como instrucciones las establecen (hace falta permiso).',
      'Now est la date+heure système. Date et Time en fonctions renvoient seulement la partie date ou heure ; en instructions elles règlent l’horloge (permission requise).',
      'Now ist Systemdatum+uhrzeit. Date und Time als Funktionen liefern nur Datum- oder Zeitteil; als Anweisungen setzen sie die Uhr (Berechtigung nötig).',
    ),
    examples: [
      {
        title: L('Parti dell’orologio', 'Clock parts', 'Partes del reloj', 'Parties de l’horloge', 'Uhrteile'),
        code: `Sub DemoNow()
    Debug.Print Year(Now)
    Debug.Print Year(Date) = Year(Now)
    Debug.Print Hour(Time) >= 0
End Sub`,
        result: L('Anno corrente, True, True.', 'Current year, True, True.', 'Año actual, True, True.', 'Année courante, True, True.', 'Aktuelles Jahr, True, True.'),
      },
    ],
    related: ['timer', 'date-parts', 'date-type'],
  },
  {
    id: 'timer',
    name: 'Timer',
    syntax: 'Timer As Single',
    scope: ['core', 'excel', 'access'],
    category: 'fn-date',
    subcategory: 'now',
    description: L(
      'Secondi trascorsi dalla mezzanotte (con frazione). Utile per misurare durate brevi nello stesso giorno.',
      'Seconds since midnight (with a fraction). Useful for timing short spans on the same day.',
      'Segundos transcurridos desde medianoche (con fracción). Útil para medir duraciones cortas el mismo día.',
      'Secondes écoulées depuis minuit (avec fraction). Utile pour mesurer de courtes durées le même jour.',
      'Sekunden seit Mitternacht (mit Bruchteil). Nützlich für kurze Zeitmessungen am selben Tag.',
    ),
    examples: [
      {
        title: L('Durata di un ciclo', 'Duration of a loop', 'Duración de un bucle', 'Durée d’une boucle', 'Dauer einer Schleife'),
        code: `Sub DemoTimer()
    Dim t0 As Single, i As Long, n As Long
    t0 = Timer
    For i = 1 To 10000
        n = n + 1
    Next i
    Debug.Print n, Timer - t0 >= 0
End Sub`,
        result: L('10000    True (durata ≥ 0 secondi).', '10000    True (elapsed ≥ 0 seconds).', '10000    True (duración ≥ 0 segundos).', '10000    True (durée ≥ 0 secondes).', '10000    True (Dauer ≥ 0 Sekunden).'),
      },
    ],
    related: ['now', 'doevents'],
  },
  {
    id: 'dateadd',
    name: 'DateAdd',
    syntax: 'DateAdd(interval, number, date)',
    scope: ['core', 'excel', 'access'],
    category: 'fn-date',
    subcategory: 'parts',
    description: L(
      'Aggiunge (o toglie se number è negativo) un intervallo: "yyyy", "q", "m", "d", "ww", "h", "n", "s".',
      'Adds (or subtracts if number is negative) an interval: "yyyy", "q", "m", "d", "ww", "h", "n", "s".',
      'Suma (o resta si number es negativo) un intervalo: "yyyy", "q", "m", "d", "ww", "h", "n", "s".',
      'Ajoute (ou retranche si number est négatif) un intervalle : "yyyy", "q", "m", "d", "ww", "h", "n", "s".',
      'Addiert (oder subtrahiert bei negativem number) ein Intervall: "yyyy", "q", "m", "d", "ww", "h", "n", "s".',
    ),
    params: [
      { name: 'interval', description: L('Unità: yyyy anno, m mese, d giorno, n minuto, …', 'Unit: yyyy year, m month, d day, n minute, …', 'Unidad: yyyy año, m mes, d día, n minuto, …', 'Unité : yyyy année, m mois, d jour, n minute, …', 'Einheit: yyyy Jahr, m Monat, d Tag, n Minute, …') },
      { name: 'number', description: L('Quante unità da aggiungere.', 'How many units to add.', 'Cuántas unidades sumar.', 'Combien d’unités à ajouter.', 'Wie viele Einheiten addiert werden.') },
    ],
    examples: [
      {
        title: L('Un mese e un giorno', 'One month and one day', 'Un mes y un día', 'Un mois et un jour', 'Ein Monat und ein Tag'),
        code: `Sub DemoDateAdd()
    Debug.Print Format(DateAdd("m", 1, #3/15/2024#), "yyyy-mm-dd")
    Debug.Print Format(DateAdd("d", -1, #3/15/2024#), "yyyy-mm-dd")
End Sub`,
        result: L('2024-04-15 poi 2024-03-14', '2024-04-15 then 2024-03-14', '2024-04-15 luego 2024-03-14', '2024-04-15 puis 2024-03-14', '2024-04-15 dann 2024-03-14'),
      },
    ],
    related: ['datediff', 'datepart', 'dateserial-timeserial'],
  },
  {
    id: 'datediff',
    name: 'DateDiff',
    syntax: 'DateDiff(interval, date1, date2, [firstdayofweek], [firstweekofyear])',
    scope: ['core', 'excel', 'access'],
    category: 'fn-date',
    subcategory: 'parts',
    description: L(
      'Differenza intera tra due date nell’unità scelta. Il segno è date2 - date1.',
      'Whole-number difference between two dates in the chosen unit. The sign is date2 - date1.',
      'Diferencia entera entre dos fechas en la unidad elegida. El signo es date2 - date1.',
      'Écart entier entre deux dates dans l’unité choisie. Le signe est date2 - date1.',
      'Ganzzahlige Differenz zweier Daten in der gewählten Einheit. Das Vorzeichen ist date2 - date1.',
    ),
    examples: [
      {
        title: L('Giorni tra due date', 'Days between two dates', 'Días entre dos fechas', 'Jours entre deux dates', 'Tage zwischen zwei Daten'),
        code: `Sub DemoDateDiff()
    Debug.Print DateDiff("d", #3/15/2024#, #3/20/2024#)
    Debug.Print DateDiff("m", #1/31/2024#, #3/1/2024#)
End Sub`,
        result: L('5 poi 2 (marzo - gennaio).', '5 then 2 (March minus January).', '5 luego 2 (marzo menos enero).', '5 puis 2 (mars moins janvier).', '5 dann 2 (März minus Januar).'),
      },
    ],
    related: ['dateadd', 'datepart', 'now'],
  },
  {
    id: 'datepart',
    name: 'DatePart',
    syntax: 'DatePart(interval, date, [firstdayofweek], [firstweekofyear])',
    scope: ['core', 'excel', 'access'],
    category: 'fn-date',
    subcategory: 'parts',
    description: L(
      'Estrae una parte (yyyy, q, m, ww, w, d, h, n, s, y) come Long. w = weekday, y = giorno dell’anno.',
      'Extracts a part (yyyy, q, m, ww, w, d, h, n, s, y) as Long. w = weekday, y = day of year.',
      'Extrae una parte (yyyy, q, m, ww, w, d, h, n, s, y) como Long. w = weekday, y = día del año.',
      'Extrait une partie (yyyy, q, m, ww, w, d, h, n, s, y) en Long. w = jour de semaine, y = jour de l’année.',
      'Extrahiert einen Teil (yyyy, q, m, ww, w, d, h, n, s, y) als Long. w = Wochentag, y = Tag des Jahres.',
    ),
    examples: [
      {
        title: L('Trimestre e settimana', 'Quarter and week', 'Trimestre y semana', 'Trimestre et semaine', 'Quartal und Woche'),
        code: `Sub DemoDatePart()
    Debug.Print DatePart("q", #3/15/2024#)
    Debug.Print DatePart("yyyy", #3/15/2024#)
    Debug.Print DatePart("m", #3/15/2024#)
End Sub`,
        result: L('1, 2024, 3', '1, 2024, 3', '1, 2024, 3', '1, 2024, 3', '1, 2024, 3'),
      },
    ],
    related: ['date-parts', 'weekday-names', 'dateadd'],
  },
  {
    id: 'dateserial-timeserial',
    name: 'DateSerial / TimeSerial',
    syntax: 'DateSerial(year, month, day) | TimeSerial(hour, minute, second)',
    scope: ['core', 'excel', 'access'],
    category: 'fn-date',
    subcategory: 'parts',
    description: L(
      'Costruisce una Date da componenti numerici. Valori fuori range vengono riportati (es. mese 13 → gennaio dell’anno dopo).',
      'Builds a Date from numeric parts. Out-of-range values roll over (e.g. month 13 → January of the next year).',
      'Construye una Date a partir de partes numéricas. Los valores fuera de rango se ajustan (p. ej. mes 13 → enero del año siguiente).',
      'Construit une Date à partir de composantes numériques. Les valeurs hors plage sont reportées (ex. mois 13 → janvier de l’année suivante).',
      'Baut ein Date aus Zahlenteilen. Werte außerhalb des Bereichs rollen über (z. B. Monat 13 → Januar des Folgejahres).',
    ),
    examples: [
      {
        title: L('Data e overflow del mese', 'Date and month overflow', 'Fecha y desbordamiento de mes', 'Date et dépassement de mois', 'Datum und Monatsüberlauf'),
        code: `Sub DemoSerial()
    Debug.Print Format(DateSerial(2024, 3, 15), "yyyy-mm-dd")
    Debug.Print Format(DateSerial(2024, 13, 1), "yyyy-mm-dd")
    Debug.Print Format(TimeSerial(14, 30, 0), "hh:nn:ss")
End Sub`,
        result: L('2024-03-15, 2025-01-01, 14:30:00', '2024-03-15, 2025-01-01, 14:30:00', '2024-03-15, 2025-01-01, 14:30:00', '2024-03-15, 2025-01-01, 14:30:00', '2024-03-15, 2025-01-01, 14:30:00'),
      },
    ],
    related: ['date-parts', 'now', 'c-other'],
  },
  {
    id: 'date-parts',
    name: 'Year / Month / Day / Hour / Minute / Second',
    syntax: 'Year(d) | Month(d) | Day(d) | Hour(d) | Minute(d) | Second(d)',
    scope: ['core', 'excel', 'access'],
    category: 'fn-date',
    subcategory: 'parts',
    description: L(
      'Estrae i componenti numerici di una Date. Year è a 4 cifre; Hour è 0–23.',
      'Extracts numeric components of a Date. Year is 4 digits; Hour is 0–23.',
      'Extrae los componentes numéricos de una Date. Year tiene 4 dígitos; Hour es 0–23.',
      'Extrait les composantes numériques d’une Date. Year a 4 chiffres ; Hour va de 0 à 23.',
      'Extrahiert die Zahlenanteile eines Date. Year hat 4 Stellen; Hour ist 0–23.',
    ),
    examples: [
      {
        title: L('Scomposizione di un istante', 'Breakdown of an instant', 'Desglose de un instante', 'Décomposition d’un instant', 'Zerlegung eines Zeitpunkts'),
        code: `Sub DemoDateParts()
    Dim d As Date
    d = #3/15/2024 14:30:45#
    Debug.Print Year(d), Month(d), Day(d)
    Debug.Print Hour(d), Minute(d), Second(d)
End Sub`,
        result: L('2024 3 15 / 14 30 45', '2024 3 15 / 14 30 45', '2024 3 15 / 14 30 45', '2024 3 15 / 14 30 45', '2024 3 15 / 14 30 45'),
      },
    ],
    related: ['datepart', 'dateserial-timeserial', 'now'],
  },
  {
    id: 'weekday-names',
    name: 'Weekday / WeekdayName / MonthName / DateValue / TimeValue',
    syntax: 'Weekday(d, [firstday]) | WeekdayName(n) | MonthName(n) | DateValue(s) | TimeValue(s)',
    scope: ['core', 'excel', 'access'],
    category: 'fn-date',
    subcategory: 'parts',
    description: L(
      'Weekday è 1–7 (default domenica = 1). WeekdayName e MonthName dipendono dal locale. DateValue/TimeValue parsano una stringa.',
      'Weekday is 1–7 (default Sunday = 1). WeekdayName and MonthName depend on the locale. DateValue/TimeValue parse a string.',
      'Weekday es 1–7 (domingo = 1 por defecto). WeekdayName y MonthName dependen de la configuración regional. DateValue/TimeValue analizan una cadena.',
      'Weekday vaut 1–7 (dimanche = 1 par défaut). WeekdayName et MonthName dépendent de la locale. DateValue/TimeValue analysent une chaîne.',
      'Weekday ist 1–7 (Standard Sonntag = 1). WeekdayName und MonthName hängen vom Locale ab. DateValue/TimeValue parsen eine Zeichenfolge.',
    ),
    examples: [
      {
        title: L('Venerdì 15 marzo 2024', 'Friday 15 March 2024', 'Viernes 15 de marzo de 2024', 'Vendredi 15 mars 2024', 'Freitag 15. März 2024'),
        code: `Sub DemoWeekdayNames()
    Debug.Print Weekday(#3/15/2024#, vbSunday)
    Debug.Print DateValue("2024-03-15") = #3/15/2024#
    Debug.Print Hour(TimeValue("14:30"))
End Sub`,
        result: L('6 (venerdì), True, 14. WeekdayName/MonthName seguono la lingua di Windows.', '6 (Friday), True, 14. WeekdayName/MonthName follow the Windows language.', '6 (viernes), True, 14. WeekdayName/MonthName siguen el idioma de Windows.', '6 (vendredi), True, 14. WeekdayName/MonthName suivent la langue de Windows.', '6 (Freitag), True, 14. WeekdayName/MonthName folgen der Windows-Sprache.'),
      },
    ],
    related: ['date-parts', 'c-other', 'format'],
  },
  {
    id: 'abs-sgn',
    name: 'Abs / Sgn',
    syntax: 'Abs(number) | Sgn(number)',
    scope: ['core', 'excel', 'access'],
    category: 'fn-math',
    subcategory: 'round',
    description: L(
      'Abs è il valore assoluto. Sgn restituisce -1, 0 o 1 a seconda del segno.',
      'Abs is the absolute value. Sgn returns -1, 0 or 1 according to the sign.',
      'Abs es el valor absoluto. Sgn devuelve -1, 0 o 1 según el signo.',
      'Abs est la valeur absolue. Sgn renvoie -1, 0 ou 1 selon le signe.',
      'Abs ist der Absolutwert. Sgn liefert -1, 0 oder 1 je nach Vorzeichen.',
    ),
    examples: [
      {
        title: L('Segno e modulo', 'Sign and magnitude', 'Signo y módulo', 'Signe et module', 'Vorzeichen und Betrag'),
        code: `Sub DemoAbsSgn()
    Debug.Print Abs(-7.5), Sgn(-7.5), Sgn(0), Sgn(3)
End Sub`,
        result: L('7.5    -1    0    1', '7.5    -1    0    1', '7.5    -1    0    1', '7.5    -1    0    1', '7.5    -1    0    1'),
      },
    ],
    related: ['int-fix', 'round'],
  },
  {
    id: 'int-fix',
    name: 'Int / Fix',
    syntax: 'Int(number) | Fix(number)',
    scope: ['core', 'excel', 'access'],
    category: 'fn-math',
    subcategory: 'round',
    description: L(
      'Entrambe tolgono la parte decimale. Int va verso -∞ (Int(-1.5) = -2); Fix verso 0 (Fix(-1.5) = -1).',
      'Both drop the fractional part. Int goes toward −∞ (Int(-1.5) = -2); Fix toward 0 (Fix(-1.5) = -1).',
      'Ambas quitan la parte decimal. Int va hacia −∞ (Int(-1.5) = -2); Fix hacia 0 (Fix(-1.5) = -1).',
      'Les deux ôtent la partie décimale. Int va vers −∞ (Int(-1.5) = -2) ; Fix vers 0 (Fix(-1.5) = -1).',
      'Beide entfernen den Bruchteil. Int geht gegen −∞ (Int(-1.5) = -2); Fix gegen 0 (Fix(-1.5) = -1).',
    ),
    examples: [
      {
        title: L('Positivi uguali, negativi diversi', 'Same on positives, different on negatives', 'Iguales en positivos, distintos en negativos', 'Identiques sur les positifs, différents sur les négatifs', 'Gleich bei Positiven, verschieden bei Negativen'),
        code: `Sub DemoIntFix()
    Debug.Print Int(1.9), Fix(1.9)
    Debug.Print Int(-1.5), Fix(-1.5)
End Sub`,
        result: L('1 1 / -2 -1', '1 1 / -2 -1', '1 1 / -2 -1', '1 1 / -2 -1', '1 1 / -2 -1'),
      },
    ],
    related: ['round', 'abs-sgn', 'c-numeric'],
  },
  {
    id: 'round',
    name: 'Round',
    syntax: 'Round(expression, [numdecimalplaces])',
    scope: ['core', 'excel', 'access'],
    category: 'fn-math',
    subcategory: 'round',
    description: L(
      'Arrotonda a N decimali con banker’s rounding: .5 va verso il pari (Round(2.5) = 2, Round(3.5) = 4).',
      'Rounds to N decimals with banker’s rounding: .5 goes to even (Round(2.5) = 2, Round(3.5) = 4).',
      'Redondea a N decimales con banker’s rounding: .5 va al par (Round(2.5) = 2, Round(3.5) = 4).',
      'Arrondit à N décimales avec l’arrondi du banquier : .5 va vers le pair (Round(2.5) = 2, Round(3.5) = 4).',
      'Rundet auf N Dezimalen mit Banker’s Rounding: .5 geht zur geraden Zahl (Round(2.5) = 2, Round(3.5) = 4).',
    ),
    examples: [
      {
        title: L('Mezzi verso il pari', 'Halves toward even', 'Mitades hacia el par', 'Moitiés vers le pair', 'Hälften zur geraden Zahl'),
        code: `Sub DemoRound()
    Debug.Print Round(2.5, 0), Round(3.5, 0)
    Debug.Print Round(1.26, 1)
End Sub`,
        result: L('2 4 / 1.3', '2 4 / 1.3', '2 4 / 1.3', '2 4 / 1.3', '2 4 / 1.3'),
      },
    ],
    notes: L(
      'CInt e CLng usano lo stesso arrotondamento del banchiere. Per “sempre 0.5 in su” usa un metodo proprio.',
      'CInt and CLng use the same banker’s rounding. For “always 0.5 up” write your own method.',
      'CInt y CLng usan el mismo redondeo del banquero. Para “siempre 0.5 hacia arriba” escribe tu propio método.',
      'CInt et CLng utilisent le même arrondi du banquier. Pour « toujours 0,5 vers le haut », écrivez votre propre méthode.',
      'CInt und CLng nutzen dasselbe Banker’s Rounding. Für „0,5 immer aufwärts“ eigene Methode schreiben.',
    ),
    related: ['int-fix', 'c-numeric', 'abs-sgn'],
  },
  {
    id: 'sqr',
    name: 'Sqr',
    syntax: 'Sqr(number)',
    scope: ['core', 'excel', 'access'],
    category: 'fn-math',
    subcategory: 'trig',
    description: L(
      'Radice quadrata. number deve essere ≥ 0, altrimenti errore 5 (chiamata di procedura non valida).',
      'Square root. number must be ≥ 0, otherwise error 5 (invalid procedure call).',
      'Raíz cuadrada. number debe ser ≥ 0; si no, error 5 (llamada a procedimiento no válida).',
      'Racine carrée. number doit être ≥ 0, sinon erreur 5 (appel de procédure non valide).',
      'Quadratwurzel. number muss ≥ 0 sein, sonst Fehler 5 (ungültiger Prozeduraufruf).',
    ),
    examples: [
      {
        title: L('Radice di 9 e di 2', 'Root of 9 and of 2', 'Raíz de 9 y de 2', 'Racine de 9 et de 2', 'Wurzel von 9 und von 2'),
        code: `Sub DemoSqr()
    Debug.Print Sqr(9)
    Debug.Print Round(Sqr(2), 3)
End Sub`,
        result: L('3 poi 1.414', '3 then 1.414', '3 luego 1.414', '3 puis 1.414', '3 dann 1.414'),
      },
    ],
    related: ['trig', 'abs-sgn'],
  },
  {
    id: 'trig',
    name: 'Sin / Cos / Tan / Atn',
    syntax: 'Sin(angle) | Cos(angle) | Tan(angle) | Atn(number)',
    scope: ['core', 'excel', 'access'],
    category: 'fn-math',
    subcategory: 'trig',
    description: L(
      'Funzioni trigonometriche in radianti. Atn è l’arcotangente (restituisce radianti). π ≈ Atn(1) * 4.',
      'Trigonometric functions in radians. Atn is arctangent (returns radians). π ≈ Atn(1) * 4.',
      'Funciones trigonométricas en radianes. Atn es el arco tangente (devuelve radianes). π ≈ Atn(1) * 4.',
      'Fonctions trigonométriques en radians. Atn est l’arctangente (renvoie des radians). π ≈ Atn(1) * 4.',
      'Trigonometrische Funktionen im Bogenmaß. Atn ist Arkustangens (liefert Bogenmaß). π ≈ Atn(1) * 4.',
    ),
    examples: [
      {
        title: L('90° e π', '90° and π', '90° y π', '90° et π', '90° und π'),
        code: `Sub DemoTrig()
    Dim pi As Double
    pi = Atn(1) * 4
    Debug.Print Round(Sin(pi / 2), 6)
    Debug.Print Round(Cos(0), 6)
    Debug.Print Round(Tan(0), 6)
End Sub`,
        result: L('1, 1, 0', '1, 1, 0', '1, 1, 0', '1, 1, 0', '1, 1, 0'),
      },
    ],
    related: ['sqr', 'log-exp'],
  },
  {
    id: 'log-exp',
    name: 'Log / Exp',
    syntax: 'Log(number) | Exp(number)',
    scope: ['core', 'excel', 'access'],
    category: 'fn-math',
    subcategory: 'trig',
    description: L(
      'Log è il logaritmo naturale (base e). Exp è e^x. Per log in base 10: Log(x) / Log(10).',
      'Log is the natural logarithm (base e). Exp is e^x. For log base 10: Log(x) / Log(10).',
      'Log es el logaritmo natural (base e). Exp es e^x. Para log en base 10: Log(x) / Log(10).',
      'Log est le logarithme népérien (base e). Exp est e^x. Pour le log base 10 : Log(x) / Log(10).',
      'Log ist der natürliche Logarithmus (Basis e). Exp ist e^x. Für Log Basis 10: Log(x) / Log(10).',
    ),
    examples: [
      {
        title: L('e^0 e ln(e)', 'e^0 and ln(e)', 'e^0 y ln(e)', 'e^0 et ln(e)', 'e^0 und ln(e)'),
        code: `Sub DemoLogExp()
    Debug.Print Exp(0)
    Debug.Print Round(Log(Exp(1)), 6)
    Debug.Print Round(Log(100) / Log(10), 6)
End Sub`,
        result: L('1, 1, 2', '1, 1, 2', '1, 1, 2', '1, 1, 2', '1, 1, 2'),
      },
    ],
    related: ['trig', 'sqr'],
  },
  {
    id: 'rnd',
    name: 'Rnd / Randomize',
    syntax: 'Randomize [number] | Rnd[(number)]',
    scope: ['core', 'excel', 'access'],
    category: 'fn-math',
    subcategory: 'trig',
    description: L(
      'Rnd restituisce un Single in [0, 1). Randomize inizializza il seme (di solito con Timer). Senza Randomize la sequenza si ripete tra sessioni IDE.',
      'Rnd returns a Single in [0, 1). Randomize seeds the generator (usually from Timer). Without Randomize the sequence repeats across IDE sessions.',
      'Rnd devuelve un Single en [0, 1). Randomize inicia la semilla (normalmente con Timer). Sin Randomize la secuencia se repite entre sesiones del IDE.',
      'Rnd renvoie un Single dans [0, 1). Randomize initialise la graine (souvent avec Timer). Sans Randomize la séquence se répète entre sessions IDE.',
      'Rnd liefert ein Single in [0, 1). Randomize setzt den Samen (meist mit Timer). Ohne Randomize wiederholt sich die Folge zwischen IDE-Sitzungen.',
    ),
    examples: [
      {
        title: L('Dado 1..6', 'Die 1..6', 'Dado 1..6', 'Dé 1..6', 'Würfel 1..6'),
        code: `Sub DemoRnd()
    Randomize
    Debug.Print Int(Rnd() * 6) + 1
End Sub`,
        result: L('Un intero casuale tra 1 e 6 (cambia a ogni esecuzione).', 'A random integer from 1 to 6 (changes each run).', 'Un entero aleatorio entre 1 y 6 (cambia en cada ejecución).', 'Un entier aléatoire entre 1 et 6 (change à chaque exécution).', 'Eine Zufallszahl von 1 bis 6 (ändert sich bei jedem Lauf).'),
      },
    ],
    notes: L(
      'Rnd(0) ripete l’ultimo valore; Rnd negativo usa quel numero come seme.',
      'Rnd(0) repeats the last value; a negative Rnd uses that number as the seed.',
      'Rnd(0) repite el último valor; un Rnd negativo usa ese número como semilla.',
      'Rnd(0) répète la dernière valeur ; un Rnd négatif utilise ce nombre comme graine.',
      'Rnd(0) wiederholt den letzten Wert; ein negatives Rnd nutzt diese Zahl als Samen.',
    ),
    related: ['int-fix', 'timer'],
  },
  {
    id: 'c-numeric',
    name: 'CInt / CLng / CDbl / CSng / CByte / CCur',
    syntax: 'CInt(x) | CLng(x) | CDbl(x) | CSng(x) | CByte(x) | CCur(x)',
    scope: ['core', 'excel', 'access'],
    category: 'fn-convert',
    subcategory: 'cast',
    description: L(
      'Cast espliciti verso i tipi numerici. CInt/CLng usano banker’s rounding. Fuori range → overflow. Il locale vale per il decimale nelle stringhe.',
      'Explicit casts to numeric types. CInt/CLng use banker’s rounding. Out of range → overflow. Locale applies to the decimal in strings.',
      'Conversiones explícitas a tipos numéricos. CInt/CLng usan redondeo del banquero. Fuera de rango → desbordamiento. La configuración regional afecta al decimal en cadenas.',
      'Conversions explicites vers les types numériques. CInt/CLng utilisent l’arrondi du banquier. Hors plage → dépassement. La locale s’applique au séparateur décimal des chaînes.',
      'Explizite Umwandlungen in Zahlentypen. CInt/CLng nutzen Banker’s Rounding. Außerhalb des Bereichs → Überlauf. Locale gilt für das Dezimalzeichen in Zeichenfolgen.',
    ),
    examples: [
      {
        title: L('Arrotondamento e valuta', 'Rounding and currency', 'Redondeo y moneda', 'Arrondi et devise', 'Runden und Währung'),
        code: `Sub DemoCNumeric()
    Debug.Print CInt(2.5), CInt(3.5)
    Debug.Print TypeName(CLng(100000)), CLng(100000)
    Debug.Print TypeName(CCur(10.1))
End Sub`,
        result: L('2 4 / Long 100000 / Currency', '2 4 / Long 100000 / Currency', '2 4 / Long 100000 / Currency', '2 4 / Long 100000 / Currency', '2 4 / Long 100000 / Currency'),
      },
    ],
    related: ['c-other', 'integer', 'long', 'round'],
  },
  {
    id: 'c-other',
    name: 'CBool / CDate / CStr / CVar',
    syntax: 'CBool(x) | CDate(x) | CStr(x) | CVar(x)',
    scope: ['core', 'excel', 'access'],
    category: 'fn-convert',
    subcategory: 'cast',
    description: L(
      'CBool è True se x ≠ 0. CDate interpreta date secondo il locale. CStr converte in testo. CVar wrappa in Variant.',
      'CBool is True if x ≠ 0. CDate parses dates according to the locale. CStr converts to text. CVar wraps in a Variant.',
      'CBool es True si x ≠ 0. CDate interpreta fechas según la configuración regional. CStr convierte a texto. CVar envuelve en Variant.',
      'CBool est True si x ≠ 0. CDate interprète les dates selon la locale. CStr convertit en texte. CVar enveloppe en Variant.',
      'CBool ist True wenn x ≠ 0. CDate liest Daten gemäß Locale. CStr wandelt in Text. CVar packt in Variant.',
    ),
    examples: [
      {
        title: L('Zero, data e testo', 'Zero, date and text', 'Cero, fecha y texto', 'Zéro, date et texte', 'Null, Datum und Text'),
        code: `Sub DemoCOther()
    Debug.Print CBool(0), CBool(5)
    Debug.Print Year(CDate(#3/15/2024#))
    Debug.Print CStr(12) & "x"
End Sub`,
        result: L('False True / 2024 / 12x', 'False True / 2024 / 12x', 'False True / 2024 / 12x', 'False True / 2024 / 12x', 'False True / 2024 / 12x'),
      },
    ],
    notes: L(
      'Per una data indipendente dal locale usa DateSerial, non CDate su stringa ambigua.',
      'For a locale-independent date use DateSerial, not CDate on an ambiguous string.',
      'Para una fecha independiente de la configuración usa DateSerial, no CDate sobre una cadena ambigua.',
      'Pour une date indépendante de la locale, utilisez DateSerial, pas CDate sur une chaîne ambiguë.',
      'Für ein locale-unabhängiges Datum DateSerial verwenden, nicht CDate bei mehrdeutiger Zeichenfolge.',
    ),
    related: ['c-numeric', 'boolean', 'date-type', 'dateserial-timeserial'],
  },
  {
    id: 'is-functions',
    name: 'Funzioni Is*',
    syntax: 'IsEmpty | IsNull | IsNumeric | IsDate | IsArray | IsObject | IsError | IsMissing',
    scope: ['core', 'excel', 'access'],
    category: 'fn-convert',
    subcategory: 'info',
    description: L(
      'Test sul contenuto di un Variant: vuoto, Null, numerico, data, array, oggetto, errore CVErr, argomento Optional omesso.',
      'Tests on a Variant’s contents: empty, Null, numeric, date, array, object, CVErr error, omitted Optional argument.',
      'Pruebas sobre el contenido de un Variant: vacío, Null, numérico, fecha, matriz, objeto, error CVErr, argumento Optional omitido.',
      'Tests sur le contenu d’un Variant : vide, Null, numérique, date, tableau, objet, erreur CVErr, argument Optional omis.',
      'Tests zum Inhalt eines Variant: leer, Null, numerisch, Datum, Array, Objekt, CVErr-Fehler, weggelassenes Optional-Argument.',
    ),
    examples: [
      {
        title: L('Empty, numero e Optional', 'Empty, number and Optional', 'Empty, número y Optional', 'Empty, nombre et Optional', 'Empty, Zahl und Optional'),
        code: `Function Check(Optional ByVal x As Variant) As String
    Check = IsMissing(x) & " " & IsEmpty(x) & " " & IsNumeric(x)
End Function

Sub DemoIsFunctions()
    Dim v As Variant
    Debug.Print IsEmpty(v), IsDate(#3/15/2024#), IsArray(Array(1))
    Debug.Print Check(), Check(10)
End Sub`,
        result: L('True True True / True False False / False False True', 'True True True / True False False / False False True', 'True True True / True False False / False False True', 'True True True / True False False / False False True', 'True True True / True False False / False False True'),
      },
    ],
    notes: L(
      'IsEmpty è True solo su Variant non assegnato. IsMissing richiede Optional As Variant (non As Long).',
      'IsEmpty is True only on an unassigned Variant. IsMissing requires Optional As Variant (not As Long).',
      'IsEmpty es True solo en un Variant no asignado. IsMissing exige Optional As Variant (no As Long).',
      'IsEmpty n’est True que sur un Variant non assigné. IsMissing exige Optional As Variant (pas As Long).',
      'IsEmpty ist nur bei unzugewiesenem Variant True. IsMissing verlangt Optional As Variant (nicht As Long).',
    ),
    related: ['variant', 'optional', 'typename-vartype'],
  },
  {
    id: 'typename-vartype',
    name: 'TypeName / VarType',
    syntax: 'TypeName(varname) As String | VarType(varname) As Integer',
    scope: ['core', 'excel', 'access'],
    category: 'fn-convert',
    subcategory: 'info',
    description: L(
      'TypeName è il nome leggibile ("Integer", "Nothing", "Collection"). VarType è il codice vbInteger, vbLong, vbArray + tipo, …',
      'TypeName is the readable name ("Integer", "Nothing", "Collection"). VarType is the code vbInteger, vbLong, vbArray + type, …',
      'TypeName es el nombre legible ("Integer", "Nothing", "Collection"). VarType es el código vbInteger, vbLong, vbArray + tipo, …',
      'TypeName est le nom lisible ("Integer", "Nothing", "Collection"). VarType est le code vbInteger, vbLong, vbArray + type, …',
      'TypeName ist der lesbare Name ("Integer", "Nothing", "Collection"). VarType ist der Code vbInteger, vbLong, vbArray + Typ, …',
    ),
    examples: [
      {
        title: L('Tipi di letterali', 'Types of literals', 'Tipos de literales', 'Types des littéraux', 'Typen von Literalen'),
        code: `Sub DemoTypeNameVarType()
    Debug.Print TypeName(1), VarType(1)
    Debug.Print TypeName("a"), VarType("a")
    Debug.Print TypeName(Nothing)
End Sub`,
        result: L('Integer 2 / String 8 / Nothing', 'Integer 2 / String 8 / Nothing', 'Integer 2 / String 8 / Nothing', 'Integer 2 / String 8 / Nothing', 'Integer 2 / String 8 / Nothing'),
      },
    ],
    related: ['is-functions', 'variant', 'c-numeric'],
  },
  {
    id: 'rgb',
    name: 'RGB',
    syntax: 'RGB(red, green, blue) As Long',
    scope: ['core', 'excel', 'access'],
    category: 'fn-convert',
    subcategory: 'info',
    description: L(
      'Compatta tre canali 0–255 in un Long (R + G*256 + B*65536), usato da BackColor, Color, vbRed, …',
      'Packs three 0–255 channels into a Long (R + G*256 + B*65536), used by BackColor, Color, vbRed, …',
      'Empaqueta tres canales 0–255 en un Long (R + G*256 + B*65536), usado por BackColor, Color, vbRed, …',
      'Assemble trois canaux 0–255 en Long (R + G*256 + B*65536), utilisé par BackColor, Color, vbRed, …',
      'Packt drei Kanäle 0–255 in ein Long (R + G*256 + B*65536), genutzt von BackColor, Color, vbRed, …',
    ),
    examples: [
      {
        title: L('Arancione 255,128,0', 'Orange 255,128,0', 'Naranja 255,128,0', 'Orange 255,128,0', 'Orange 255,128,0'),
        code: `Sub DemoRGB()
    Debug.Print RGB(255, 128, 0)
    Debug.Print RGB(255, 0, 0) = vbRed
End Sub`,
        result: L('33023 poi True', '33023 then True', '33023 luego True', '33023 puis True', '33023 dann True'),
      },
    ],
    related: ['byte', 'c-numeric'],
  },
  {
    id: 'choose-switch',
    name: 'Choose / Switch',
    syntax: 'Choose(index, expr1[, exprN...]) | Switch(expr1, val1[, exprN, valN...])',
    scope: ['core', 'excel', 'access'],
    category: 'fn-convert',
    subcategory: 'info',
    description: L(
      'Choose seleziona l’argomento 1-based. Switch restituisce il valore della prima espressione True. Entrambi valutano tutti gli argomenti.',
      'Choose picks the 1-based argument. Switch returns the value of the first True expression. Both evaluate every argument.',
      'Choose elige el argumento 1-based. Switch devuelve el valor de la primera expresión True. Ambos evalúan todos los argumentos.',
      'Choose sélectionne l’argument 1-based. Switch renvoie la valeur de la première expression True. Les deux évaluent tous les arguments.',
      'Choose wählt das 1-basierte Argument. Switch liefert den Wert des ersten True-Ausdrucks. Beide werten alle Argumente aus.',
    ),
    examples: [
      {
        title: L('Indice e prima condizione vera', 'Index and first true condition', 'Índice y primera condición verdadera', 'Index et première condition vraie', 'Index und erste wahre Bedingung'),
        code: `Sub DemoChooseSwitch()
    Debug.Print Choose(2, "a", "b", "c")
    Debug.Print Switch(False, "no", True, "yes", True, "later")
End Sub`,
        result: L('b poi yes', 'b then yes', 'b luego yes', 'b puis yes', 'b dann yes'),
      },
    ],
    related: ['iif', 'select-case'],
  },
  {
    id: 'array-bounds',
    name: 'Array / LBound / UBound',
    syntax: 'Array(arglist) | LBound(array, [dimension]) | UBound(array, [dimension])',
    scope: ['core', 'excel', 'access'],
    category: 'fn-array-file',
    subcategory: 'arrays',
    description: L(
      'Array crea un Variant() 0-based. LBound/UBound sono i limiti di una dimensione (default 1).',
      'Array builds a 0-based Variant(). LBound/UBound are the bounds of one dimension (default 1).',
      'Array crea un Variant() 0-based. LBound/UBound son los límites de una dimensión (predeterminada 1).',
      'Array crée un Variant() 0-based. LBound/UBound sont les bornes d’une dimension (défaut 1).',
      'Array erzeugt ein 0-basiertes Variant(). LBound/UBound sind die Grenzen einer Dimension (Standard 1).',
    ),
    examples: [
      {
        title: L('Limiti e secondo elemento', 'Bounds and second item', 'Límites y segundo elemento', 'Bornes et deuxième élément', 'Grenzen und zweites Element'),
        code: `Sub DemoArrayBounds()
    Dim a As Variant
    a = Array("uno", "due", "tre")
    Debug.Print LBound(a), UBound(a)
    Debug.Print a(1)
End Sub`,
        result: L('0 2 poi due', '0 2 then due', '0 2 luego due', '0 2 puis due', '0 2 dann due'),
      },
    ],
    related: ['split', 'filter', 'erase', 'option-base'],
  },
  {
    id: 'filter',
    name: 'Filter',
    syntax: 'Filter(sourcearray, match, [include], [compare]) As String()',
    scope: ['core', 'excel', 'access'],
    category: 'fn-array-file',
    subcategory: 'arrays',
    description: L(
      'Restituisce un array 0-based con gli elementi che contengono (o, se include=False, non contengono) match.',
      'Returns a 0-based array of items that contain (or, if include=False, do not contain) match.',
      'Devuelve una matriz 0-based con los elementos que contienen (o, si include=False, no contienen) match.',
      'Renvoie un tableau 0-based des éléments qui contiennent (ou, si include=False, ne contiennent pas) match.',
      'Liefert ein 0-basiertes Array der Elemente, die match enthalten (oder bei include=False nicht enthalten).',
    ),
    examples: [
      {
        title: L('Solo i file .txt', 'Only .txt files', 'Solo archivos .txt', 'Uniquement les fichiers .txt', 'Nur .txt-Dateien'),
        code: `Sub DemoFilter()
    Dim a As Variant, f As Variant
    a = Array("a.txt", "b.csv", "c.txt")
    f = Filter(a, ".txt")
    Debug.Print UBound(f) - LBound(f) + 1, f(0)
End Sub`,
        result: L('2    a.txt', '2    a.txt', '2    a.txt', '2    a.txt', '2    a.txt'),
      },
    ],
    related: ['array-bounds', 'instr', 'split'],
  },
  {
    id: 'erase',
    name: 'Erase',
    syntax: 'Erase arraylist',
    scope: ['core', 'excel', 'access'],
    category: 'fn-array-file',
    subcategory: 'arrays',
    description: L(
      'Svuota un array: sui dinamici dealloca la memoria (servirà ReDim); sui fissi azzera gli elementi.',
      'Clears an array: on dynamic arrays it frees memory (ReDim needed again); on fixed arrays it zeros the items.',
      'Vacía una matriz: en dinámicas libera memoria (hará falta ReDim); en fijas pone a cero los elementos.',
      'Vide un tableau : sur les dynamiques il libère la mémoire (ReDim à nouveau) ; sur les fixes il remet les éléments à zéro.',
      'Leert ein Array: bei dynamischen gibt es Speicher frei (erneutes ReDim nötig); bei festen setzt es Elemente auf null.',
    ),
    examples: [
      {
        title: L('Array dinamico dopo Erase', 'Dynamic array after Erase', 'Matriz dinámica tras Erase', 'Tableau dynamique après Erase', 'Dynamisches Array nach Erase'),
        code: `Sub DemoErase()
    Dim a() As Long
    ReDim a(2)
    a(0) = 9
    Erase a
    On Error Resume Next
    Debug.Print a(0)
    Debug.Print Err.Number
    On Error GoTo 0
End Sub`,
        result: L('Errore 9 (indice fuori intervallo) perché l’array non è più allocato.', 'Error 9 (subscript out of range) because the array is no longer allocated.', 'Error 9 (índice fuera de intervalo) porque la matriz ya no está asignada.', 'Erreur 9 (indice hors limites) car le tableau n’est plus alloué.', 'Fehler 9 (Index außerhalb des Bereichs), weil das Array nicht mehr alloziert ist.'),
      },
    ],
    related: ['redim-preserve', 'array-bounds'],
  },
  {
    id: 'open-close',
    name: 'Open / Close',
    syntax: 'Open pathname For mode [Access access] As [#]filenumber | Close [#]filenumber',
    scope: ['core', 'excel', 'access'],
    category: 'fn-array-file',
    subcategory: 'io',
    description: L(
      'Apre un file sequenziale (Input, Output, Append), binario o Random e lo associa a un numero. Close lo chiude e svuota il buffer.',
      'Opens a sequential (Input, Output, Append), binary or Random file and binds it to a number. Close closes it and flushes the buffer.',
      'Abre un archivo secuencial (Input, Output, Append), binario o Random y lo asocia a un número. Close lo cierra y vacía el búfer.',
      'Ouvre un fichier séquentiel (Input, Output, Append), binaire ou Random et l’associe à un numéro. Close le ferme et vide le tampon.',
      'Öffnet eine sequentielle (Input, Output, Append), binäre oder Random-Datei und bindet sie an eine Nummer. Close schließt sie und leert den Puffer.',
    ),
    params: [
      { name: 'mode', description: L('Input, Output, Append, Binary, Random.', 'Input, Output, Append, Binary, Random.', 'Input, Output, Append, Binary, Random.', 'Input, Output, Append, Binary, Random.', 'Input, Output, Append, Binary, Random.') },
      { name: 'filenumber', description: L('Intero da FreeFile, usato come #n.', 'Integer from FreeFile, used as #n.', 'Entero de FreeFile, usado como #n.', 'Entier issu de FreeFile, utilisé comme #n.', 'Ganzzahl von FreeFile, verwendet als #n.') },
    ],
    examples: [
      {
        title: L('Crea e chiude un file in TEMP', 'Create and close a file in TEMP', 'Crear y cerrar un archivo en TEMP', 'Créer et fermer un fichier dans TEMP', 'Datei in TEMP anlegen und schließen'),
        code: `Sub DemoOpenClose()
    Dim path As String, fn As Integer
    path = Environ("TEMP") & "\\vba-open.txt"
    fn = FreeFile
    Open path For Output As #fn
    Print #fn, "ciao"
    Close #fn
    Debug.Print FileLen(path) > 0
    Kill path
End Sub`,
        result: L('True: il file è stato scritto e poi eliminato.', 'True: the file was written and then deleted.', 'True: el archivo se escribió y luego se eliminó.', 'True : le fichier a été écrit puis supprimé.', 'True: die Datei wurde geschrieben und danach gelöscht.'),
      },
    ],
    related: ['freefile-eof', 'line-input-print-write', 'file-ops'],
  },
  {
    id: 'line-input-print-write',
    name: 'Line Input / Print # / Write #',
    syntax: 'Line Input #fn, var | Print #fn, [expr] | Write #fn, [expr]',
    scope: ['core', 'excel', 'access'],
    category: 'fn-array-file',
    subcategory: 'io',
    description: L(
      'Line Input legge una riga. Print # scrive testo libero. Write # scrive campi delimitati da virgola e virgolette, indipendenti dal locale.',
      'Line Input reads one line. Print # writes free text. Write # writes comma-quoted fields, locale-independent.',
      'Line Input lee una línea. Print # escribe texto libre. Write # escribe campos entre comillas y comas, independientes de la configuración regional.',
      'Line Input lit une ligne. Print # écrit du texte libre. Write # écrit des champs entre guillemets et virgules, indépendants de la locale.',
      'Line Input liest eine Zeile. Print # schreibt freien Text. Write # schreibt kommagetrennte, in Anführungszeichen gesetzte Felder, locale-unabhängig.',
    ),
    examples: [
      {
        title: L('Print vs Write', 'Print vs Write', 'Print frente a Write', 'Print contre Write', 'Print gegen Write'),
        code: `Sub DemoPrintWrite()
    Dim path As String, fn As Integer, riga As String
    path = Environ("TEMP") & "\\vba-io.txt"
    fn = FreeFile
    Open path For Output As #fn
    Print #fn, "ciao"; 1
    Write #fn, "ciao", 1
    Close #fn
    fn = FreeFile
    Open path For Input As #fn
    Line Input #fn, riga
    Debug.Print riga
    Line Input #fn, riga
    Debug.Print riga
    Close #fn
    Kill path
End Sub`,
        result: L('Prima riga tipo ciao1; seconda "ciao",1', 'First line like ciao1; second "ciao",1', 'Primera línea tipo ciao1; segunda "ciao",1', 'Première ligne du type ciao1 ; seconde "ciao",1', 'Erste Zeile wie ciao1; zweite "ciao",1'),
      },
    ],
    related: ['open-close', 'freefile-eof', 'split'],
  },
  {
    id: 'freefile-eof',
    name: 'FreeFile / EOF',
    syntax: 'FreeFile[(rangenumber)] | EOF(filenumber)',
    scope: ['core', 'excel', 'access'],
    category: 'fn-array-file',
    subcategory: 'io',
    description: L(
      'FreeFile restituisce il prossimo numero di file libero (1–255, o 256–511 con rangenumber=1). EOF è True a fine file in Input.',
      'FreeFile returns the next free file number (1–255, or 256–511 with rangenumber=1). EOF is True at end of file in Input.',
      'FreeFile devuelve el siguiente número de archivo libre (1–255, o 256–511 con rangenumber=1). EOF es True al final del archivo en Input.',
      'FreeFile renvoie le prochain numéro de fichier libre (1–255, ou 256–511 avec rangenumber=1). EOF est True en fin de fichier en Input.',
      'FreeFile liefert die nächste freie Dateinummer (1–255 oder 256–511 mit rangenumber=1). EOF ist True am Dateiende bei Input.',
    ),
    examples: [
      {
        title: L('Leggi fino a EOF', 'Read until EOF', 'Leer hasta EOF', 'Lire jusqu’à EOF', 'Lesen bis EOF'),
        code: `Sub DemoFreeFileEof()
    Dim path As String, fn As Integer, riga As String, n As Long
    path = Environ("TEMP") & "\\vba-eof.txt"
    fn = FreeFile
    Open path For Output As #fn
    Print #fn, "a"
    Print #fn, "b"
    Close #fn
    fn = FreeFile
    Open path For Input As #fn
    Do Until EOF(fn)
        Line Input #fn, riga
        n = n + 1
    Loop
    Close #fn
    Kill path
    Debug.Print n
End Sub`,
        result: L('Finestra Immediata: 2', 'Immediate Window: 2', 'Ventana Inmediato: 2', 'Fenêtre Exécution : 2', 'Direktbereich: 2'),
      },
    ],
    related: ['open-close', 'line-input-print-write', 'do-loop'],
  },
  {
    id: 'dir',
    name: 'Dir',
    syntax: 'Dir[(pathname, [attributes])] As String',
    scope: ['core', 'excel', 'access'],
    category: 'fn-array-file',
    subcategory: 'fs',
    description: L(
      'Prima chiamata con un percorso (jolly * ?) restituisce il primo nome; le successive Dir() senza argomenti scorrono. "" = fine elenco.',
      'First call with a path (wildcards * ?) returns the first name; later Dir() calls with no args walk the list. "" = end of list.',
      'La primera llamada con una ruta (comodines * ?) devuelve el primer nombre; las Dir() siguientes sin argumentos recorren. "" = fin de lista.',
      'Le premier appel avec un chemin (jokers * ?) renvoie le premier nom ; les Dir() suivants sans argument parcourent. "" = fin de liste.',
      'Erster Aufruf mit Pfad (Wildcards * ?) liefert den ersten Namen; spätere Dir() ohne Argumente gehen die Liste durch. "" = Listenende.',
    ),
    examples: [
      {
        title: L('Esiste TEMP?', 'Does TEMP exist?', '¿Existe TEMP?', 'TEMP existe-t-il ?', 'Existiert TEMP?'),
        code: `Sub DemoDir()
    Dim p As String
    p = Environ("TEMP")
    Debug.Print Len(Dir(p, vbDirectory)) > 0
End Sub`,
        result: L('True se la cartella TEMP esiste.', 'True if the TEMP folder exists.', 'True si existe la carpeta TEMP.', 'True si le dossier TEMP existe.', 'True, wenn der TEMP-Ordner existiert.'),
      },
    ],
    related: ['folder-ops', 'file-ops', 'environ'],
  },
  {
    id: 'file-ops',
    name: 'Kill / Name / FileCopy',
    syntax: 'Kill pathname | Name oldpath As newpath | FileCopy source, destination',
    scope: ['core', 'excel', 'access'],
    category: 'fn-array-file',
    subcategory: 'fs',
    description: L(
      'Kill elimina file (accetta * ?). Name rinomina o sposta. FileCopy copia. Tutti falliscono se il file è aperto o il percorso non esiste.',
      'Kill deletes files (accepts * ?). Name renames or moves. FileCopy copies. All fail if the file is open or the path does not exist.',
      'Kill elimina archivos (acepta * ?). Name renombra o mueve. FileCopy copia. Todos fallan si el archivo está abierto o la ruta no existe.',
      'Kill supprime des fichiers (accepte * ?). Name renomme ou déplace. FileCopy copie. Tous échouent si le fichier est ouvert ou le chemin absent.',
      'Kill löscht Dateien (akzeptiert * ?). Name benennt um oder verschiebt. FileCopy kopiert. Alle scheitern, wenn die Datei offen ist oder der Pfad fehlt.',
    ),
    examples: [
      {
        title: L('Copia, rinomina, elimina', 'Copy, rename, delete', 'Copiar, renombrar, eliminar', 'Copier, renommer, supprimer', 'Kopieren, umbenennen, löschen'),
        code: `Sub DemoFileOps()
    Dim src As String, dst As String, ren As String, fn As Integer
    src = Environ("TEMP") & "\\vba-a.txt"
    dst = Environ("TEMP") & "\\vba-b.txt"
    ren = Environ("TEMP") & "\\vba-c.txt"
    fn = FreeFile
    Open src For Output As #fn
    Print #fn, "x"
    Close #fn
    FileCopy src, dst
    Name dst As ren
    Debug.Print FileLen(ren) > 0
    Kill src
    Kill ren
End Sub`,
        result: L('True: la copia rinominata esisteva; poi entrambi i file vengono cancellati.', 'True: the renamed copy existed; then both files are deleted.', 'True: la copia renombrada existía; luego se borran ambos archivos.', 'True : la copie renommée existait ; puis les deux fichiers sont effacés.', 'True: die umbenannte Kopie existierte; danach werden beide Dateien gelöscht.'),
      },
    ],
    related: ['dir', 'folder-ops', 'open-close'],
  },
  {
    id: 'folder-ops',
    name: 'MkDir / RmDir / FileLen / GetAttr',
    syntax: 'MkDir path | RmDir path | FileLen(pathname) | GetAttr(pathname)',
    scope: ['core', 'excel', 'access'],
    category: 'fn-array-file',
    subcategory: 'fs',
    description: L(
      'MkDir crea una cartella, RmDir la elimina se vuota. FileLen è la dimensione in byte. GetAttr restituisce flag (vbDirectory, vbReadOnly, vbHidden, …).',
      'MkDir creates a folder, RmDir deletes it if empty. FileLen is the size in bytes. GetAttr returns flags (vbDirectory, vbReadOnly, vbHidden, …).',
      'MkDir crea una carpeta, RmDir la elimina si está vacía. FileLen es el tamaño en bytes. GetAttr devuelve flags (vbDirectory, vbReadOnly, vbHidden, …).',
      'MkDir crée un dossier, RmDir le supprime s’il est vide. FileLen est la taille en octets. GetAttr renvoie des drapeaux (vbDirectory, vbReadOnly, vbHidden, …).',
      'MkDir erstellt einen Ordner, RmDir löscht ihn wenn leer. FileLen ist die Größe in Bytes. GetAttr liefert Flags (vbDirectory, vbReadOnly, vbHidden, …).',
    ),
    examples: [
      {
        title: L('Cartella temporanea e attributi', 'Temp folder and attributes', 'Carpeta temporal y atributos', 'Dossier temporaire et attributs', 'Temp-Ordner und Attribute'),
        code: `Sub DemoFolderOps()
    Dim p As String
    p = Environ("TEMP") & "\\vba-demo-dir"
    On Error Resume Next
    RmDir p
    On Error GoTo 0
    MkDir p
    Debug.Print (GetAttr(p) And vbDirectory) = vbDirectory
    RmDir p
End Sub`,
        result: L('True: GetAttr segnala vbDirectory; poi la cartella viene rimossa.', 'True: GetAttr reports vbDirectory; then the folder is removed.', 'True: GetAttr indica vbDirectory; luego se elimina la carpeta.', 'True : GetAttr signale vbDirectory ; puis le dossier est supprimé.', 'True: GetAttr meldet vbDirectory; danach wird der Ordner entfernt.'),
      },
    ],
    related: ['dir', 'file-ops', 'environ'],
  },
]









