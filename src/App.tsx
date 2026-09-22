import { useEffect, useMemo, useRef, useState } from 'react'
import type { Lang, Scope } from './types'
import { LANGS, t } from './types'
import { ui } from './i18n'
import { highlightVba } from './highlight'
import {
  adsAvailable,
  initAds,
  isAdFreeActive,
  onAdFreeGranted,
  onCommandOpened,
  watchAdForAdFreeWindow,
} from './ads'
import { RemoveAdsButton } from './monetization/RemoveAdsButton'
import { privacyIntro, privacySections } from './privacyPolicy'
import {
  allCommands,
  categories,
  filterCommands,
  getCommand,
  relatedCommands,
  stats,
} from './data/catalog'

type ScopeFilter = Scope | 'all'
type Theme = 'light' | 'dark'
type View = 'browse' | 'detail' | 'nav'

const LANG_KEY = 'vba-desk-lang'
const THEME_KEY = 'vba-desk-theme'

function readHash(): { cmd?: string; cat?: string; sub?: string; privacy?: boolean } {
  const raw = decodeURIComponent(location.hash.replace(/^#\/?/, ''))
  if (!raw) return {}
  if (raw === 'privacy') return { privacy: true }
  const parts = raw.split('/')
  if (parts[0] === 'c' && parts[1]) return { cmd: parts[1] }
  if (parts[0] === 'cat' && parts[1]) return { cat: parts[1], sub: parts[2] }
  return {}
}

function Privacy({ lang }: { lang: Lang }) {
  return (
    <div className="overview">
      <a className="back-btn copy-btn" href="#/">← {t(ui.backHome, lang)}</a>
      <div className="kicker">{t(ui.privacyPolicy, lang)}</div>
      <h1 className="display">{t(ui.privacyPolicy, lang)}</h1>
      <p className="lead">{t(privacyIntro, lang)}</p>
      {privacySections.map((s, i) => (
        <div key={i}>
          <h2 className="section-title">{t(s.heading, lang)}</h2>
          <p className="prose">{t(s.body, lang)}</p>
        </div>
      ))}
    </div>
  )
}

function writeHash(cmd?: string | null, cat?: string | null, sub?: string | null) {
  let next = '#/'
  if (cmd) next = `#/c/${cmd}`
  else if (cat && sub) next = `#/cat/${cat}/${sub}`
  else if (cat) next = `#/cat/${cat}`
  if (location.hash !== next) history.replaceState(null, '', next)
}

export default function App() {
  const [lang, setLang] = useState<Lang>(() => {
    const saved = localStorage.getItem(LANG_KEY) as Lang | null
    return LANGS.some((l) => l.id === saved) ? saved! : 'it'
  })
  const [theme, setTheme] = useState<Theme>(() => (localStorage.getItem(THEME_KEY) as Theme) || 'light')
  const [scope, setScope] = useState<ScopeFilter>('all')
  const [categoryId, setCategoryId] = useState<string | null>(null)
  const [subId, setSubId] = useState<string | null>(null)
  const [selectedId, setSelectedId] = useState<string | null>(null)
  const [query, setQuery] = useState('')
  const [copied, setCopied] = useState<string | null>(null)
  const [view, setView] = useState<View>('browse')
  const [showPrivacy, setShowPrivacy] = useState(() => readHash().privacy === true)
  const [adFree, setAdFree] = useState(() => isAdFreeActive())
  const listRef = useRef<HTMLElement | null>(null)

  useEffect(() => {
    const apply = () => {
      const h = readHash()
      setShowPrivacy(h.privacy === true)
      if (h.cmd && getCommand(h.cmd)) {
        const cmd = getCommand(h.cmd)!
        setSelectedId(cmd.id)
        setCategoryId(cmd.category)
        setSubId(cmd.subcategory)
        setView('detail')
      } else if (h.cat) {
        setSelectedId(null)
        setCategoryId(h.cat)
        setSubId(h.sub ?? null)
        setView('browse')
      }
    }
    apply()
    window.addEventListener('hashchange', apply)
    return () => window.removeEventListener('hashchange', apply)
  }, [])

  const initialLangRef = useRef(lang)
  useEffect(() => {
    initAds(initialLangRef.current)
  }, [])

  useEffect(() => onAdFreeGranted(() => setAdFree(isAdFreeActive())), [])

  useEffect(() => {
    document.documentElement.lang = lang
    document.documentElement.dataset.theme = theme
    localStorage.setItem(LANG_KEY, lang)
    localStorage.setItem(THEME_KEY, theme)
  }, [lang, theme])

  useEffect(() => {
    listRef.current?.scrollTo({ top: 0 })
  }, [categoryId, subId, query, scope])

  const visibleCats = useMemo(
    () => categories.filter((c) => scope === 'all' || c.scope === scope),
    [scope],
  )

  const commands = useMemo(
    () => filterCommands(query, scope, categoryId, subId, lang),
    [query, scope, categoryId, subId, lang],
  )

  const selected = selectedId ? getCommand(selectedId) : undefined
  const activeCat = categories.find((c) => c.id === categoryId)
  const activeSub = activeCat?.subcategories.find((s) => s.id === subId)
  const isHome = !query && !categoryId && !selectedId && scope === 'all'

  function openCommand(id: string) {
    const cmd = getCommand(id)
    if (!cmd) return
    setSelectedId(id)
    setCategoryId(cmd.category)
    setSubId(cmd.subcategory)
    setScope(cmd.scope.includes('core') ? 'core' : cmd.scope[0])
    setView('detail')
    writeHash(id)
    onCommandOpened()
  }

  async function removeAdsForAWhile() {
    await watchAdForAdFreeWindow()
  }

  function openCategory(id: string, sub?: string | null) {
    const cat = categories.find((c) => c.id === id)
    setCategoryId(id)
    setSubId(sub ?? null)
    setSelectedId(null)
    setQuery('')
    if (cat) setScope(cat.scope)
    setView('browse')
    writeHash(null, id, sub)
  }

  function openScope(next: ScopeFilter) {
    setScope(next)
    setCategoryId(null)
    setSubId(null)
    setSelectedId(null)
    setQuery('')
    setView('browse')
    writeHash()
  }

  function goHome() {
    setCategoryId(null)
    setSubId(null)
    setSelectedId(null)
    setQuery('')
    setScope('all')
    setView('browse')
    writeHash()
  }

  async function copyCode(code: string, key: string) {
    try {
      await navigator.clipboard.writeText(code)
      setCopied(key)
      window.setTimeout(() => setCopied((cur) => (cur === key ? null : cur)), 1600)
    } catch {
      /* ignore */
    }
  }

  const listTitle = query
    ? `${t(ui.resultsFor, lang)} “${query}”`
    : activeSub
      ? t(activeSub.label, lang)
      : activeCat
        ? t(activeCat.label, lang)
        : t(ui.overview, lang)

  if (showPrivacy) {
    return (
      <div className="app" data-view="browse">
        <Privacy lang={lang} />
      </div>
    )
  }

  return (
    <div className="app" data-view={view} data-ad-banner={adsAvailable && !adFree}>
      <header className="topbar">
        <a className="brand" href="#/" onClick={(e) => { e.preventDefault(); goHome() }}>
          <span className="logo" aria-hidden="true"><i /><i /></span>
          <span>
            <span className="brand-mark">VBA Desk</span>
            <span className="brand-sub">Excel · Access</span>
          </span>
        </a>
        <div className="search-wrap">
          <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
            <circle cx="7" cy="7" r="4.5" fill="none" stroke="currentColor" strokeWidth="1.4" />
            <path d="M10.4 10.4 L14 14" stroke="currentColor" strokeWidth="1.4" />
          </svg>
          <input
            value={query}
            onChange={(e) => {
              setQuery(e.target.value)
              setSelectedId(null)
              setView('browse')
            }}
            placeholder={t(ui.search, lang)}
            aria-label={t(ui.search, lang)}
          />
        </div>
        <div className="top-actions">
          {adsAvailable && (
            adFree ? (
              <span className="ad-free-badge">{t(ui.adsFreeActive, lang)}</span>
            ) : (
              <button className="icon-btn remove-ads-btn" type="button" onClick={removeAdsForAWhile} title={t(ui.removeAds, lang)} aria-label={t(ui.removeAds, lang)}>
                <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
                  <path d="M4 4l8 8M12 4l-8 8" stroke="currentColor" strokeWidth="1.5" fill="none" />
                </svg>
              </button>
            )
          )}
          {!isHome && (
            <button className="icon-btn menu-btn" type="button" onClick={() => setView('nav')} aria-label={t(ui.categories, lang)}>
              <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
                <path d="M2 4h12M2 8h12M2 12h12" fill="none" stroke="currentColor" strokeWidth="1.5" />
              </svg>
            </button>
          )}
          <select
            className="lang-select"
            value={lang}
            onChange={(e) => setLang(e.target.value as Lang)}
            aria-label={t(ui.language, lang)}
          >
            {LANGS.map((l) => (
              <option key={l.id} value={l.id}>{l.label}</option>
            ))}
          </select>
          <button
            className="icon-btn"
            type="button"
            onClick={() => setTheme((th) => (th === 'light' ? 'dark' : 'light'))}
            aria-label={theme === 'light' ? t(ui.themeDark, lang) : t(ui.themeLight, lang)}
            title={theme === 'light' ? t(ui.themeDark, lang) : t(ui.themeLight, lang)}
          >
            {theme === 'light' ? (
              <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
                <path d="M10.5 2.2A6 6 0 1 0 13.8 10 5 5 0 0 1 10.5 2.2z" fill="currentColor" />
              </svg>
            ) : (
              <svg width="16" height="16" viewBox="0 0 16 16" aria-hidden="true">
                <circle cx="8" cy="8" r="3" fill="currentColor" />
                <path d="M8 1.5v2M8 12.5v2M1.5 8h2M12.5 8h2M3.2 3.2l1.4 1.4M11.4 11.4l1.4 1.4M3.2 12.8l1.4-1.4M11.4 4.6l1.4-1.4" stroke="currentColor" strokeWidth="1.3" />
              </svg>
            )}
          </button>
        </div>
      </header>

      {isHome ? (
        <Home
          lang={lang}
          onOpenCategory={openCategory}
          onOpenCommand={openCommand}
          onOpenScope={openScope}
        />
      ) : (
        <div>
          <div className="scope-bar" role="tablist">
            {(['all', 'core', 'excel', 'access'] as ScopeFilter[]).map((s) => (
              <button
                key={s}
                type="button"
                className="scope-chip"
                data-on={scope === s}
                data-scope={s}
                onClick={() => openScope(s)}
              >
                {s === 'all' ? t(ui.all, lang) : s === 'core' ? t(ui.language, lang) : s === 'excel' ? t(ui.excel, lang) : t(ui.access, lang)}
              </button>
            ))}
          </div>

          <div className="shell">
            <nav className="nav" aria-label={t(ui.categories, lang)}>
              <button className="home-link" type="button" onClick={goHome}>{t(ui.backHome, lang)}</button>
              <div className="nav-label">{t(ui.categories, lang)}</div>
              {visibleCats.map((cat) => {
                const count = allCommands.filter((c) => c.category === cat.id).length
                const open = categoryId === cat.id
                return (
                  <div className="cat" key={cat.id}>
                    <button className="cat-btn" type="button" data-on={open} onClick={() => openCategory(cat.id)}>
                      <span>
                        <i className={`scope-dot ${cat.scope}`} />
                        {t(cat.label, lang)}
                      </span>
                      <span className="cat-count">{count}</span>
                    </button>
                    {open &&
                      cat.subcategories.map((sub) => (
                        <button
                          key={sub.id}
                          className="sub-btn"
                          type="button"
                          data-on={subId === sub.id}
                          onClick={() => openCategory(cat.id, sub.id)}
                        >
                          {t(sub.label, lang)}
                        </button>
                      ))}
                  </div>
                )
              })}
              <RemoveAdsButton locale={lang} />
            </nav>

            <aside className="list" ref={listRef}>
              <div className="list-head">
                <h2>{listTitle}</h2>
                <small>{commands.length} {t(ui.commands, lang)}</small>
              </div>
              {commands.length === 0 ? (
                <p className="empty">{t(ui.noResults, lang)}</p>
              ) : (
                commands.map((cmd) => (
                  <button
                    key={cmd.id}
                    type="button"
                    className="cmd-row"
                    data-on={selectedId === cmd.id}
                    onClick={() => openCommand(cmd.id)}
                  >
                    <span className="cmd-name">{cmd.name}</span>
                    <span className="badges">
                      {cmd.scope.includes('core') && <span className="badge core">{t(ui.coreShort, lang)}</span>}
                      {cmd.scope.includes('excel') && !cmd.scope.includes('core') && <span className="badge excel">Excel</span>}
                      {cmd.scope.includes('access') && !cmd.scope.includes('core') && <span className="badge access">Access</span>}
                    </span>
                    <span className="cmd-desc">{t(cmd.description, lang)}</span>
                  </button>
                ))
              )}
            </aside>

            <main className="detail">
              {selected ? (
                <article>
                  <button className="back-btn copy-btn" type="button" onClick={() => setView('browse')}>
                    ← {t(ui.close, lang)}
                  </button>
                  <div className="lemma-head">
                    <div>
                      <div className="crumb">
                        {t(ui.appliesTo, lang)} · {selected.scope.includes('core') ? t(ui.both, lang) : selected.scope.map((s) => (s === 'excel' ? 'Excel' : 'Access')).join(' · ')}
                        {activeCat ? ` · ${t(activeCat.label, lang)}` : ''}
                        {activeSub ? ` / ${t(activeSub.label, lang)}` : ''}
                      </div>
                      <h1 className="lemma-name">{selected.name}</h1>
                    </div>
                    <div className="badges">
                      {selected.scope.includes('core') && <span className="badge core">{t(ui.language, lang)}</span>}
                      {selected.scope.includes('excel') && <span className="badge excel">Excel</span>}
                      {selected.scope.includes('access') && <span className="badge access">Access</span>}
                    </div>
                  </div>

                  <p className="prose">{t(selected.description, lang)}</p>
                  <div className="kicker">{t(ui.syntax, lang)}</div>
                  <div className="syntax">{selected.syntax}</div>

                  {selected.params && selected.params.length > 0 && (
                    <>
                      <h2 className="section-title">{t(ui.parameters, lang)}</h2>
                      <table className="params">
                        <thead>
                          <tr>
                            <th>{t(ui.parameters, lang)}</th>
                            <th></th>
                          </tr>
                        </thead>
                        <tbody>
                          {selected.params.map((p) => (
                            <tr key={p.name}>
                              <td>
                                <code>{p.name}</code>{' '}
                                <span className="opt">{p.optional ? t(ui.optional, lang) : t(ui.required, lang)}</span>
                              </td>
                              <td>{t(p.description, lang)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </>
                  )}

                  <h2 className="section-title">{t(ui.examples, lang)}</h2>
                  {selected.examples.map((ex, i) => {
                    const key = `${selected.id}-${i}`
                    return (
                      <section className="example" key={key}>
                        <div className="example-top">
                          <h3>{t(ex.title, lang)}</h3>
                          <button type="button" className="copy-btn" onClick={() => copyCode(ex.code, key)}>
                            {copied === key ? t(ui.copied, lang) : t(ui.copy, lang)}
                          </button>
                        </div>
                        <pre dangerouslySetInnerHTML={{ __html: highlightVba(ex.code) }} />
                        <div className="result">
                          <span className="result-label">{t(ui.immediate, lang)}</span>
                          {t(ex.result, lang)}
                        </div>
                      </section>
                    )
                  })}

                  {selected.notes && (
                    <>
                      <h2 className="section-title">{t(ui.notes, lang)}</h2>
                      <p className="notes">{t(selected.notes, lang)}</p>
                    </>
                  )}

                  {relatedCommands(selected).length > 0 && (
                    <>
                      <h2 className="section-title">{t(ui.related, lang)}</h2>
                      <div className="related">
                        {relatedCommands(selected).map((rel) => (
                          <button key={rel.id} type="button" onClick={() => openCommand(rel.id)}>
                            {rel.name}
                          </button>
                        ))}
                      </div>
                    </>
                  )}
                </article>
              ) : activeCat ? (
                <div className="overview">
                  <button className="back-btn copy-btn" type="button" onClick={() => setView('browse')}>
                    ← {t(ui.close, lang)}
                  </button>
                  <div className="kicker">{t(ui.browse, lang)}</div>
                  <h1 className="display">{t(activeCat.label, lang)}</h1>
                  <p className="lead">{t(activeCat.blurb, lang)}</p>
                  <p className="prose">{t(ui.selectHint, lang)}</p>
                  <div className="related">
                    {activeCat.subcategories.map((sub) => (
                      <button key={sub.id} type="button" onClick={() => openCategory(activeCat.id, sub.id)}>
                        {t(sub.label, lang)}
                      </button>
                    ))}
                  </div>
                </div>
              ) : (
                <div className="overview">
                  <div className="kicker">{t(ui.browse, lang)}</div>
                  <h1 className="display">{listTitle}</h1>
                  <p className="lead">{t(ui.selectHint, lang)}</p>
                </div>
              )}
            </main>
          </div>
        </div>
      )}
    </div>
  )
}

function Home({
  lang,
  onOpenCategory,
  onOpenCommand,
  onOpenScope,
}: {
  lang: Lang
  onOpenCategory: (id: string, sub?: string | null) => void
  onOpenCommand: (id: string) => void
  onOpenScope: (scope: ScopeFilter) => void
}) {
  const grouped = {
    core: categories.filter((c) => c.scope === 'core'),
    excel: categories.filter((c) => c.scope === 'excel'),
    access: categories.filter((c) => c.scope === 'access'),
  }
  const starters = ['msgbox', 'dim', 'xl-range-value', 'ac-dlookup', 'for-next', 'xl-range-find']
    .map((id) => getCommand(id))
    .filter((c): c is NonNullable<typeof c> => Boolean(c))

  return (
    <div className="home">
      <div className="home-inner">
        <section className="hero">
          <div>
            <div className="kicker">{t(ui.appName, lang)}</div>
            <h1 className="display">{t(ui.tagline, lang)}</h1>
            <p className="lead">{t(ui.homeLead, lang)}</p>
          </div>
          <div className="stat-row">
            <div className="stat"><b>{stats.total}</b><span>{t(ui.commands, lang)}</span></div>
            <div className="stat"><b>{stats.categories}</b><span>{t(ui.categories, lang)}</span></div>
            <div className="stat"><b>{stats.excel}</b><span>Excel</span></div>
            <div className="stat"><b>{stats.access}</b><span>Access</span></div>
          </div>
        </section>

        <h2 className="section-title">{t(ui.startHere, lang)}</h2>
        <div className="scope-tiles">
          <button type="button" className="scope-tile" data-scope="core" onClick={() => onOpenScope('core')}>
            <h2>{t(ui.language, lang)}</h2>
            <p>{t(ui.languageDesc, lang)}</p>
            <span className="count">{stats.core} {t(ui.commands, lang)}</span>
          </button>
          <button type="button" className="scope-tile" data-scope="excel" onClick={() => onOpenScope('excel')}>
            <h2>{t(ui.excel, lang)}</h2>
            <p>{t(ui.excelDesc, lang)}</p>
            <span className="count">{stats.excel} {t(ui.commands, lang)}</span>
          </button>
          <button type="button" className="scope-tile" data-scope="access" onClick={() => onOpenScope('access')}>
            <h2>{t(ui.access, lang)}</h2>
            <p>{t(ui.accessDesc, lang)}</p>
            <span className="count">{stats.access} {t(ui.commands, lang)}</span>
          </button>
        </div>

        <h2 className="section-title">{t(ui.browse, lang)}</h2>
        <div className="cat-mosaic">
          {(['core', 'excel', 'access'] as Scope[]).map((scope) => (
            <div className="mosaic-col" key={scope}>
              <h3>
                <i className={`scope-dot ${scope}`} />
                {scope === 'core' ? t(ui.language, lang) : scope === 'excel' ? t(ui.excel, lang) : t(ui.access, lang)}
              </h3>
              {grouped[scope].map((cat) => {
                const count = allCommands.filter((c) => c.category === cat.id).length
                return (
                  <button key={cat.id} type="button" className="cat-card" onClick={() => onOpenCategory(cat.id)}>
                    <strong>
                      {t(cat.label, lang)}
                      <span>{count}</span>
                    </strong>
                    <p>{t(cat.blurb, lang)}</p>
                  </button>
                )
              })}
            </div>
          ))}
        </div>

        <h2 className="section-title">{t(ui.quickStart, lang)}</h2>
        <div className="quick">
          {starters.map((cmd) => (
            <button key={cmd.id} type="button" onClick={() => onOpenCommand(cmd.id)}>
              {cmd.name}
            </button>
          ))}
        </div>

        <RemoveAdsButton locale={lang} />

        <p className="footer-note">
          {t(ui.footer, lang)} · <a href="#/privacy">{t(ui.privacyPolicy, lang)}</a>
        </p>
      </div>
    </div>
  )
}
