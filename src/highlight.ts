const KEYWORDS =
  /\b(And|As|Boolean|ByRef|ByVal|Byte|Call|Case|CBool|CByte|CCur|CDate|CDbl|CDec|Const|CInt|CLng|CLngLng|CLngPtr|Close|Collection|Const|CSng|CStr|CVar|Currency|Date|Decimal|Dim|Do|Double|Each|Else|ElseIf|End|Enum|Erase|Error|Event|Exit|False|For|Function|Get|GoSub|GoTo|If|Imp|Integer|Is|Let|Lib|Like|Long|LongLong|LongPtr|Loop|LSet|Me|Mod|New|Next|Not|Nothing|Object|On|Open|Option|Optional|Or|ParamArray|Preserve|Print|Private|Property|Public|RaiseEvent|ReDim|Rem|Resume|Return|RSet|Select|Set|Single|Static|Stop|String|Sub|Then|To|True|Type|Until|Variant|Wend|While|With|Xor)\b/gi

const OBJECTS =
  /\b(Application|ThisWorkbook|ActiveWorkbook|Workbooks|Worksheets|Sheets|ActiveSheet|ActiveCell|Selection|Range|Cells|CurrentDb|DoCmd|Err|Debug|Forms|Reports|Screen|Me|DAO|ADODB|Recordset|QueryDef|TableDef|CurrentProject|WorksheetFunction|ChartObjects|PivotTables|Names|PageSetup|FileSystemObject)\b/g

export function highlightVba(code: string): string {
  const escaped = code
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')

  const parts = escaped.split(/("(?:[^"]|"")*")/g)
  return parts
    .map((part, i) => {
      if (i % 2 === 1) return `<span class="tok-str">${part}</span>`
      return part
        .replace(/('.*)$/gm, '<span class="tok-cmt">$1</span>')
        .replace(KEYWORDS, '<span class="tok-kw">$1</span>')
        .replace(OBJECTS, '<span class="tok-obj">$1</span>')
        .replace(/\b(\d+\.?\d*)\b/g, '<span class="tok-num">$1</span>')
    })
    .join('')
}
