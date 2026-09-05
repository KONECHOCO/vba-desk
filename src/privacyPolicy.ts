import type { I18n } from './types'
import { L } from './types'

export interface PrivacySection {
  heading: I18n
  body: I18n
}

export const privacyIntro: I18n = L(
  'VBA Desk non richiede un account e non raccoglie dati personali per il proprio funzionamento: il catalogo di comandi VBA è incluso nell’app e funziona offline.',
  'VBA Desk requires no account and collects no personal data for its own operation: the VBA command catalog ships with the app and works offline.',
  'VBA Desk no requiere cuenta y no recopila datos personales para su propio funcionamiento: el catálogo de comandos VBA viene incluido en la app y funciona sin conexión.',
  'VBA Desk ne nécessite aucun compte et ne collecte aucune donnée personnelle pour son propre fonctionnement : le catalogue de commandes VBA est inclus dans l’application et fonctionne hors ligne.',
  'VBA Desk erfordert kein Konto und erhebt für den eigenen Betrieb keine personenbezogenen Daten: Der VBA-Befehlskatalog ist in der App enthalten und funktioniert offline.',
)

export const privacySections: PrivacySection[] = [
  {
    heading: L('Pubblicità', 'Advertising', 'Publicidad', 'Publicité', 'Werbung'),
    body: L(
      'L’app mostra pubblicità tramite Unity LevelPlay (già noto come ironSource) per rimanere gratuita. Unity LevelPlay e le reti pubblicitarie con cui media possono raccogliere identificativi pubblicitari del dispositivo (es. IDFA su iOS), indirizzo IP, informazioni sul dispositivo e sull’uso dell’app per mostrare e misurare annunci, anche personalizzati se hai dato il consenso. Alla prima apertura ti chiediamo di accettare o rifiutare la pubblicità personalizzata; su iOS ti verrà inoltre chiesta l’autorizzazione al tracciamento (App Tracking Transparency). Puoi cambiare idea in qualsiasi momento dalle impostazioni del dispositivo.',
      'The app shows ads via Unity LevelPlay (formerly ironSource) to stay free. Unity LevelPlay and the ad networks it mediates may collect device advertising identifiers (e.g. IDFA on iOS), IP address, and device/usage information to show and measure ads, including personalized ones if you consent. On first launch we ask you to accept or decline personalized advertising; on iOS you will also be asked for App Tracking Transparency authorization. You can change your choice at any time from your device settings.',
      'La app muestra anuncios mediante Unity LevelPlay (antes ironSource) para seguir siendo gratuita. Unity LevelPlay y las redes publicitarias que media pueden recopilar identificadores publicitarios del dispositivo (p. ej. IDFA en iOS), dirección IP e información del dispositivo/uso para mostrar y medir anuncios, incluidos los personalizados si has dado tu consentimiento. En el primer inicio te pedimos aceptar o rechazar la publicidad personalizada; en iOS también se te pedirá la autorización de App Tracking Transparency. Puedes cambiar tu elección en cualquier momento desde los ajustes del dispositivo.',
      'L’application affiche des publicités via Unity LevelPlay (anciennement ironSource) pour rester gratuite. Unity LevelPlay et les réseaux publicitaires qu’il utilise peuvent collecter des identifiants publicitaires de l’appareil (ex. IDFA sur iOS), l’adresse IP et des informations sur l’appareil/l’utilisation afin d’afficher et de mesurer les publicités, y compris personnalisées si vous y consentez. Au premier lancement, nous vous demandons d’accepter ou de refuser la publicité personnalisée ; sur iOS, l’autorisation App Tracking Transparency vous sera également demandée. Vous pouvez changer d’avis à tout moment depuis les réglages de votre appareil.',
      'Die App zeigt Werbung über Unity LevelPlay (früher ironSource), um kostenlos zu bleiben. Unity LevelPlay und die von ihm vermittelten Werbenetzwerke können Werbe-IDs des Geräts (z. B. IDFA unter iOS), IP-Adresse sowie Geräte-/Nutzungsinformationen erheben, um Anzeigen zu zeigen und zu messen, einschließlich personalisierter Anzeigen bei erteilter Zustimmung. Beim ersten Start bitten wir Sie, personalisierter Werbung zuzustimmen oder sie abzulehnen; unter iOS wird zusätzlich die App-Tracking-Transparency-Berechtigung abgefragt. Sie können Ihre Wahl jederzeit in den Geräteeinstellungen ändern.',
    ),
  },
  {
    heading: L('Dati raccolti', 'Data collected', 'Datos recopilados', 'Données collectées', 'Erhobene Daten'),
    body: L(
      'Non chiediamo né conserviamo nome, email o altri dati identificativi. I soli dati trattati sono quelli tecnici raccolti dall’SDK pubblicitario descritti sopra, secondo le rispettive informative dei partner pubblicitari.',
      'We do not ask for or store your name, email, or other identifying data. The only data processed is the technical data collected by the advertising SDK described above, per the ad partners’ own policies.',
      'No solicitamos ni almacenamos tu nombre, correo electrónico u otros datos identificativos. Los únicos datos tratados son los datos técnicos recopilados por el SDK publicitario descrito arriba, según las políticas de cada socio publicitario.',
      'Nous ne demandons ni ne conservons votre nom, votre e-mail ou d’autres données identifiantes. Les seules données traitées sont les données techniques collectées par le SDK publicitaire décrit ci-dessus, selon les politiques propres à chaque partenaire publicitaire.',
      'Wir fragen weder Name noch E-Mail-Adresse oder andere identifizierende Daten ab und speichern sie nicht. Es werden ausschließlich die oben beschriebenen technischen Daten des Werbe-SDK gemäß den jeweiligen Richtlinien der Werbepartner verarbeitet.',
    ),
  },
  {
    heading: L('Minori', 'Children', 'Menores', 'Mineurs', 'Minderjährige'),
    body: L(
      'L’app non è rivolta a bambini sotto i 13 anni e non richiede pubblicità personalizzata a utenti che si dichiarano minorenni.',
      'The app is not directed at children under 13 and does not request personalized advertising for users who identify as minors.',
      'La app no está dirigida a menores de 13 años y no solicita publicidad personalizada a usuarios que se identifiquen como menores.',
      'L’application ne s’adresse pas aux enfants de moins de 13 ans et ne demande pas de publicité personnalisée aux utilisateurs se déclarant mineurs.',
      'Die App richtet sich nicht an Kinder unter 13 Jahren und fordert bei Nutzern, die sich als minderjährig ausweisen, keine personalisierte Werbung an.',
    ),
  },
  {
    heading: L('Contatti', 'Contact', 'Contacto', 'Contact', 'Kontakt'),
    body: L(
      'Per domande su questa informativa scrivi a konechoco@gmail.com.',
      'For questions about this policy, contact konechoco@gmail.com.',
      'Para preguntas sobre esta política, escribe a konechoco@gmail.com.',
      'Pour toute question concernant cette politique, contactez konechoco@gmail.com.',
      'Bei Fragen zu dieser Richtlinie wenden Sie sich an konechoco@gmail.com.',
    ),
  },
]
