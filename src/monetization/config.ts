// Per-app monetization settings. Everything else in src/monetization/ is
// identical across the KONECHOCO guide apps — only this file differs.
export const monetizationConfig = {
  appName: 'VBA Desk',
  accentColor: '#1f9d55',
  accentTextColor: '#ffffff',
  // Non-consumable "Remove ads" product. Must exist with this exact ID in
  // App Store Connect (In-App Purchases) and Google Play Console (One-time products).
  removeAdsProductId: 'com.konechoco.vbadesk.removeads',
  // Must match `levelplay.networks` in package.json. Add 'admob' to both (plus
  // `levelplay.admob.appId`) once the AdMob app exists — without the App ID
  // the AdMob adapter crashes the app at launch.
  adNetworks: ['unityads'],
  requestTracking: false,
  // Localized at runtime by initAds() in src/ads.ts.
  consentCopy: undefined as Record<string, string> | undefined,
}
