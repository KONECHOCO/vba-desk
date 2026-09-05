import { Capacitor } from '@capacitor/core'
import { AdEvent, LevelPlayAds, type ConsentOptions } from 'capacitor-levelplay-ads'
import type { Lang } from './types'
import { t } from './types'
import { ui } from './i18n'

// TODO: replace with your Unity LevelPlay dashboard values
// (Grow > Apps > VBA Desk > iOS, and its Banner / Interstitial / Rewarded ad units).
const LEVELPLAY_APP_KEY_IOS = 'REPLACE_WITH_IOS_APP_KEY'
const AD_UNIT_BANNER_IOS = 'REPLACE_WITH_IOS_BANNER_AD_UNIT_ID'
const AD_UNIT_INTERSTITIAL_IOS = 'REPLACE_WITH_IOS_INTERSTITIAL_AD_UNIT_ID'
const AD_UNIT_REWARDED_IOS = 'REPLACE_WITH_IOS_REWARDED_AD_UNIT_ID'

// Public URL of the Privacy Policy page (docs/APP_STORE_LAUNCH_CHECKLIST.md tracks
// enabling GitHub Pages so this resolves).
const PRIVACY_POLICY_URL = 'https://konechoco.github.io/vba-desk/#/privacy'

const AD_FREE_KEY = 'vba-desk-ads-free-until'
const AD_FREE_MINUTES = 30
const INTERSTITIAL_EVERY_N_OPENS = 6
const INTERSTITIAL_MIN_GAP_MS = 3 * 60 * 1000

let initialized = false
let openCount = 0
let lastInterstitialAt = 0
const adFreeListeners = new Set<() => void>()

/** Notified when a rewarded ad actually grants the ad-free window (see `initAds`). */
export function onAdFreeGranted(cb: () => void): () => void {
  adFreeListeners.add(cb)
  return () => adFreeListeners.delete(cb)
}

export const adsAvailable = Capacitor.isNativePlatform()

export function isAdFreeActive(): boolean {
  const until = Number(localStorage.getItem(AD_FREE_KEY) || 0)
  return until > Date.now()
}

function consentOptions(lang: Lang): ConsentOptions {
  // The bundled "custom" consent alert reads title/message/button text straight
  // off the native options dict; these aren't in the plugin's public TS type but
  // are read by both native implementations (verified against their source).
  return {
    privacyPolicyUrl: PRIVACY_POLICY_URL,
    locale: lang,
    appName: 'VBA Desk',
    title: t(ui.consentTitle, lang),
    message: t(ui.consentMessage, lang),
    acceptButtonText: t(ui.consentAccept, lang),
    declineButtonText: t(ui.consentDecline, lang),
  } as ConsentOptions
}

export async function initAds(lang: Lang): Promise<void> {
  if (!adsAvailable || initialized) return
  try {
    await LevelPlayAds.requestConsentInfo(consentOptions(lang))
    await LevelPlayAds.initialize({ appKey: LEVELPLAY_APP_KEY_IOS, isTesting: false })
    initialized = true
    await LevelPlayAds.loadInterstitial({ adUnitId: AD_UNIT_INTERSTITIAL_IOS })
    await LevelPlayAds.loadRewarded({ adUnitId: AD_UNIT_REWARDED_IOS })
    await LevelPlayAds.addListener(AdEvent.RewardedRewarded, () => {
      const until = Date.now() + AD_FREE_MINUTES * 60 * 1000
      localStorage.setItem(AD_FREE_KEY, String(until))
      LevelPlayAds.hideBanner().catch(() => {})
      LevelPlayAds.loadRewarded({ adUnitId: AD_UNIT_REWARDED_IOS }).catch(() => {})
      adFreeListeners.forEach((cb) => cb())
    })
    if (!isAdFreeActive()) {
      await LevelPlayAds.createBanner({ adUnitId: AD_UNIT_BANNER_IOS, position: 'BOTTOM' })
    }
  } catch {
    // Ads are a monetization layer, never block the reference content on failure.
  }
}

export async function onCommandOpened(): Promise<void> {
  if (!adsAvailable || !initialized || isAdFreeActive()) return
  openCount += 1
  if (openCount < INTERSTITIAL_EVERY_N_OPENS) return
  const now = Date.now()
  if (now - lastInterstitialAt < INTERSTITIAL_MIN_GAP_MS) return
  openCount = 0
  try {
    const { isReady } = await LevelPlayAds.isInterstitialReady()
    if (isReady) {
      lastInterstitialAt = now
      await LevelPlayAds.showInterstitial()
      await LevelPlayAds.loadInterstitial({ adUnitId: AD_UNIT_INTERSTITIAL_IOS })
    }
  } catch {
    // ignore — reload on next open attempt
  }
}

/**
 * Triggers the rewarded ad. The ad-free window is granted only once
 * `AdEvent.RewardedRewarded` actually fires (see the listener in `initAds`) —
 * a completed `showRewarded()` call does not by itself mean the user finished
 * watching and earned the reward.
 */
export async function watchAdForAdFreeWindow(): Promise<boolean> {
  if (!adsAvailable) return false
  try {
    const { isReady } = await LevelPlayAds.isRewardedReady()
    if (!isReady) return false
    await LevelPlayAds.showRewarded()
    return true
  } catch {
    return false
  }
}

export function restoreBannerIfExpired(): void {
  if (!adsAvailable || !initialized || isAdFreeActive()) return
  LevelPlayAds.showBanner().catch(() => {})
}
