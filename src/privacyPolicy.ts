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
      'L’app mostra pubblicità non personalizzata per rimanere gratuita: video tramite Unity Ads (Unity Technologies) e un banner tramite Google AdMob. Questi SDK possono trattare indirizzo IP, informazioni sul dispositivo e sull’uso dell’app per mostrare e misurare gli annunci. L’app non richiede il tracciamento (App Tracking Transparency) e non usa l’identificativo pubblicitario per seguirti tra app di altre aziende. Nello Spazio economico europeo e nel Regno Unito ti chiediamo il consenso con il modulo certificato di Google. Puoi rimuovere la pubblicità con un acquisto una tantum.',
      'The app shows non-personalized ads to stay free: videos via Unity Ads (Unity Technologies) and a banner via Google AdMob. These SDKs may process IP address and device/usage information to show and measure ads. The app does not request tracking (App Tracking Transparency) and does not use the advertising identifier to follow you across other companies’ apps. In the EEA and UK we ask for consent with Google’s certified consent form. You can remove ads with a one-time purchase.',
      'La app muestra anuncios no personalizados para seguir siendo gratuita: vídeos mediante Unity Ads (Unity Technologies) y un banner mediante Google AdMob. Estos SDK pueden tratar la dirección IP e información del dispositivo/uso para mostrar y medir anuncios. La app no solicita el seguimiento (App Tracking Transparency) ni usa el identificador publicitario para seguirte en apps de otras empresas. En el EEE y el Reino Unido pedimos el consentimiento con el formulario certificado de Google. Puedes quitar los anuncios con una compra única.',
      'L’application affiche des publicités non personnalisées pour rester gratuite : vidéos via Unity Ads (Unity Technologies) et une bannière via Google AdMob. Ces SDK peuvent traiter l’adresse IP et des informations sur l’appareil/l’utilisation afin d’afficher et de mesurer les publicités. L’application ne demande pas le suivi (App Tracking Transparency) et n’utilise pas l’identifiant publicitaire pour vous suivre dans les apps d’autres sociétés. Dans l’EEE et au Royaume-Uni, nous demandons votre consentement via le formulaire certifié de Google. Vous pouvez supprimer les publicités par un achat unique.',
      'Die App zeigt nicht personalisierte Werbung, um kostenlos zu bleiben: Videos über Unity Ads (Unity Technologies) und ein Banner über Google AdMob. Diese SDKs können IP-Adresse sowie Geräte-/Nutzungsinformationen verarbeiten, um Anzeigen zu zeigen und zu messen. Die App fragt kein Tracking an (App Tracking Transparency) und nutzt die Werbe-ID nicht, um Sie über Apps anderer Unternehmen hinweg zu verfolgen. Im EWR und im Vereinigten Königreich fragen wir die Einwilligung über das zertifizierte Formular von Google ab. Sie können die Werbung mit einem einmaligen Kauf entfernen.',
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
