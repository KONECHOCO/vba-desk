import type { Command, Lang, Scope } from '../types'
import { categories } from './categories'
import { coreCommands } from './core'
import { excelCommands } from './excel'
import { accessCommands } from './access'

export { categories }

export const allCommands: Command[] = [...coreCommands, ...excelCommands, ...accessCommands]

const byId = new Map(allCommands.map((c) => [c.id, c]))

export function getCommand(id: string): Command | undefined {
  return byId.get(id)
}

export function filterCommands(
  query: string,
  scope: Scope | 'all',
  categoryId: string | null,
  subcategoryId: string | null,
  lang: Lang,
): Command[] {
  const q = query.trim().toLowerCase()
  return allCommands.filter((cmd) => {
    if (scope === 'core' && !cmd.scope.includes('core')) return false
    if (scope === 'excel') {
      const excelOnly = cmd.scope.includes('excel') && !cmd.scope.includes('core')
      if (q ? !cmd.scope.includes('excel') : !excelOnly) return false
    }
    if (scope === 'access') {
      const accessOnly = cmd.scope.includes('access') && !cmd.scope.includes('core')
      if (q ? !cmd.scope.includes('access') : !accessOnly) return false
    }
    if (!q) {
      if (categoryId && cmd.category !== categoryId) return false
      if (subcategoryId && cmd.subcategory !== subcategoryId) return false
      return true
    }
    const hay = [
      cmd.name,
      cmd.syntax,
      cmd.id,
      cmd.description[lang],
      cmd.description.en,
      ...(cmd.params ?? []).map((p) => p.name),
    ]
      .join('\n')
      .toLowerCase()
    return hay.includes(q)
  })
}

export function relatedCommands(cmd: Command): Command[] {
  return (cmd.related ?? []).map((id) => byId.get(id)).filter((c): c is Command => Boolean(c))
}

export const stats = {
  total: allCommands.length,
  core: coreCommands.length,
  excel: excelCommands.length,
  access: accessCommands.length,
  categories: categories.length,
}
