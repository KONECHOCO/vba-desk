export type Lang = 'it' | 'en' | 'es' | 'fr' | 'de'

export type I18n = Record<Lang, string>

export type Scope = 'core' | 'excel' | 'access'

export interface Param {
  name: string
  optional?: boolean
  description: I18n
}

export interface Example {
  title: I18n
  code: string
  result: I18n
}

export interface Command {
  id: string
  name: string
  syntax: string
  scope: Scope[]
  category: string
  subcategory: string
  description: I18n
  params?: Param[]
  examples: Example[]
  notes?: I18n
  related?: string[]
}

export interface Subcategory {
  id: string
  label: I18n
}

export interface Category {
  id: string
  scope: Scope
  label: I18n
  blurb: I18n
  subcategories: Subcategory[]
}

export const LANGS: { id: Lang; label: string }[] = [
  { id: 'it', label: 'Italiano' },
  { id: 'en', label: 'English' },
  { id: 'es', label: 'Español' },
  { id: 'fr', label: 'Français' },
  { id: 'de', label: 'Deutsch' },
]

export function L(it: string, en: string, es: string, fr: string, de: string): I18n {
  return { it, en, es, fr, de }
}

export function t(text: I18n, lang: Lang): string {
  return text[lang] || text.en
}
