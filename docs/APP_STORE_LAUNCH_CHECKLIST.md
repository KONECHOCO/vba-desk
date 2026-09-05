# VBA Desk — App Store launch checklist

Everything code-side is done: Capacitor iOS/Android wrap, Unity LevelPlay ads (banner +
interstitial + rewarded), the `#/privacy` page, `.github/workflows/deploy-pages.yml`, and
`codemagic.yaml`. What's left only you can do (dashboard access, real credentials).

## 1. Unity LevelPlay (ads)

1. Create the app at https://platform.ironsrc.com (Grow → Apps → Add app) for iOS, bundle id
   `com.konechoco.vbadesk`.
2. Create three ad units: Banner, Interstitial, Rewarded Video.
3. Copy the **iOS App Key** and the three **ad unit IDs**, then replace the placeholders in
   [src/ads.ts](../src/ads.ts):
   - `LEVELPLAY_APP_KEY_IOS`
   - `AD_UNIT_BANNER_IOS`
   - `AD_UNIT_INTERSTITIAL_IOS`
   - `AD_UNIT_REWARDED_IOS`
4. Optional, later: add mediation networks (AdMob, AppLovin, …) for better fill by setting
   `levelplay.networks` in `package.json` and re-running `npx cap sync` — each network needs
   its own account/app-id first.

## 2. Apple Developer

- Register the Bundle ID `com.konechoco.vbadesk` at
  https://developer.apple.com/account/resources/identifiers/list (or let Codemagic's
  automatic signing register it on first build).

## 3. App Store Connect — create the app record

App Store Connect → My Apps → **+** → New App:

- Platform: iOS
- Name: `VBA Desk`
- Primary language: Italian
- Bundle ID: `com.konechoco.vbadesk`
- SKU: e.g. `vbadesk001`

Then fill in:

- **Privacy Policy URL**: `https://konechoco.github.io/vba-desk/#/privacy` (see step 6)
- **App Privacy** (data collection questionnaire): declare **Identifiers → Device ID** and
  **Usage Data** as collected, linked to the user, used for **Third-Party Advertising** — this
  matches what Unity LevelPlay actually collects. See LevelPlay's own disclosure guidance:
  https://developers.is.com/ironsource-mobile/general/app-privacy-details-ios/
- Age rating questionnaire: answer "Yes" to *Unrestricted Web Access* = No, but mark
  *Infrequent/Mild* ads if asked, since third-party ads may serve age-inappropriate content.
- Pricing: Free.

## 4. Codemagic

1. Team settings → Integrations → **App Store Connect** → add your API key (if not already
   present from your other apps) and name the integration `codemagic_appstore` (or update
   the name in [codemagic.yaml](../codemagic.yaml) to match your existing one).
2. Confirm the app is connected to the `KONECHOCO/vba-desk` repo.
3. Trigger the `ios-workflow` (push to `main`, or run manually) — it builds, signs, and
   uploads to **TestFlight** automatically (`submit_to_app_store: false`, so public review
   submission stays a separate, deliberate step).

## 5. TestFlight

Install the build via TestFlight and click through the app once for a sanity check —
banner shows, an interstitial appears after a handful of command opens, the rewarded button
works, language switch still works, dark mode still works.

## 6. GitHub Pages (privacy policy hosting)

Repo → **Settings → Pages → Source: GitHub Actions**. Once enabled, every push to `main`
publishes to `https://konechoco.github.io/vba-desk/` via
[.github/workflows/deploy-pages.yml](../.github/workflows/deploy-pages.yml). The privacy page
is at the `/#/privacy` route.

## 7. Screenshots

Reuse [scripts/capture-screens.mjs](../scripts/capture-screens.mjs) or capture fresh ones from
the iOS Simulator for the required sizes (6.9" and 6.5" iPhone at minimum). The existing
`docs/screenshots/*.png` can be resized/cropped as a starting point.

## 8. Submit for review

Once the app record, metadata, screenshots and a processed TestFlight build are all in place,
use the `appstore-connect-manager` skill (`asc_client.py`) to review everything one more time
and submit — always with your explicit go-ahead first.

---

## Drafted App Store metadata (all 5 languages)

Paste directly into App Store Connect → App Information / Version Information.

### Name (all locales)
`VBA Desk`

### Subtitle (30 chars max)

| Locale | Subtitle |
|---|---|
| it | Guida VBA per Excel e Access |
| en | VBA guide for Excel and Access |
| es | Guia VBA para Excel y Access |
| fr | Guide VBA pour Excel et Access |
| de | VBA-Leitfaden für Excel/Access |

### Promotional text (170 chars max, editable anytime without review)

- **it**: Oltre 100 comandi VBA per Excel e Access con sintassi, parametri, esempi copiabili e
  il risultato atteso. Il riferimento da tenere sempre a portata di mano.
- **en**: 100+ VBA commands for Excel and Access with syntax, parameters, copyable examples
  and the expected result. The desk reference you'll keep coming back to.
- **es**: Más de 100 comandos VBA para Excel y Access con sintaxis, parámetros, ejemplos
  copiables y el resultado esperado. La referencia que siempre tendrás a mano.
- **fr**: Plus de 100 commandes VBA pour Excel et Access avec syntaxe, paramètres, exemples
  à copier et le résultat attendu. La référence à garder toujours sous la main.
- **de**: Über 100 VBA-Befehle für Excel und Access mit Syntax, Parametern, kopierbaren
  Beispielen und dem erwarteten Ergebnis. Das Nachschlagewerk für jeden Tag.

### Description

- **it**: VBA Desk è una guida di consultazione per chi scrive macro in Excel e Access. Ogni
  comando è raggruppato per funzione — linguaggio, Excel, Access — e mostra sintassi,
  parametri, almeno un esempio di codice pronto da copiare e il risultato che otterrai nella
  Finestra Immediata o nel foglio di lavoro. Naviga per categoria o cerca direttamente il
  comando che ti serve, cambia lingua tra italiano, inglese, spagnolo, francese e tedesco, e
  passa dal tema chiaro a quello scuro in un tocco. Nessun account richiesto: il catalogo è
  incluso nell'app e funziona anche offline.
- **en**: VBA Desk is a desk reference for anyone writing macros in Excel and Access. Every
  command is grouped by function — language, Excel, Access — and shows its syntax,
  parameters, at least one ready-to-copy code example, and the result you'll get in the
  Immediate Window or worksheet. Browse by category or search directly for the command you
  need, switch between Italian, English, Spanish, French and German, and flip from light to
  dark theme in one tap. No account required — the catalog ships with the app and works
  offline.
- **es**: VBA Desk es una referencia de escritorio para quienes escriben macros en Excel y
  Access. Cada comando está agrupado por función — lenguaje, Excel, Access — y muestra su
  sintaxis, parámetros, al menos un ejemplo de código listo para copiar y el resultado que
  obtendrás en la Ventana Inmediato o en la hoja de cálculo. Explora por categoría o busca
  directamente el comando que necesitas, cambia entre italiano, inglés, español, francés y
  alemán, y pasa del tema claro al oscuro con un toque. No requiere cuenta: el catálogo viene
  incluido en la app y funciona sin conexión.
- **fr**: VBA Desk est une référence de bureau pour quiconque écrit des macros dans Excel et
  Access. Chaque commande est regroupée par fonction — langage, Excel, Access — et affiche sa
  syntaxe, ses paramètres, au moins un exemple de code prêt à copier et le résultat obtenu
  dans la fenêtre Exécution ou la feuille de calcul. Parcourez par catégorie ou recherchez
  directement la commande dont vous avez besoin, changez entre italien, anglais, espagnol,
  français et allemand, et passez du thème clair au thème sombre en un geste. Aucun compte
  requis : le catalogue est inclus dans l'application et fonctionne hors ligne.
- **de**: VBA Desk ist ein Nachschlagewerk für alle, die Makros in Excel und Access schreiben.
  Jeder Befehl ist nach Funktion gruppiert — Sprache, Excel, Access — und zeigt Syntax,
  Parameter, mindestens ein kopierbares Codebeispiel und das Ergebnis, das Sie im
  Direktbereich oder im Arbeitsblatt erhalten. Durchsuchen Sie nach Kategorie oder suchen Sie
  direkt nach dem benötigten Befehl, wechseln Sie zwischen Italienisch, Englisch, Spanisch,
  Französisch und Deutsch, und schalten Sie mit einem Tipp zwischen hellem und dunklem Design
  um. Kein Konto erforderlich: Der Katalog ist in der App enthalten und funktioniert offline.

### Keywords (100 chars max, comma-separated, no spaces)

| Locale | Keywords |
|---|---|
| it | `vba,macro,excel,access,funzioni,programmazione,guida,riferimento,office,codice` |
| en | `vba,macro,excel,access,vba functions,programming,reference,guide,office,code` |
| es | `vba,macro,excel,access,funciones,programacion,guia,referencia,office,codigo` |
| fr | `vba,macro,excel,access,fonctions,programmation,guide,reference,office,code` |
| de | `vba,makro,excel,access,funktionen,programmierung,leitfaden,referenz,office,code` |

### What's New (first version)

- **it**: Prima versione.
- **en**: First release.
- **es**: Primera versión.
- **fr**: Première version.
- **de**: Erste Version.
