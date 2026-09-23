import type { Lang } from './types'
import {
  adsEnabled,
  bootstrapAds,
  showInterstitialAfterNavigation,
  showRewardedForAdFree,
} from './monetization/ads'
import { isAdFree, subscribePremium } from './monetization/premium'

// Thin adapter over the shared src/monetization/ module (launch video,
// interstitials, banner, rewarded 30-min ad-free window, "Remove ads" purchase).
// LevelPlay keys come from Codemagic (VITE_LEVELPLAY_* variables).

export const adsAvailable = adsEnabled

export function isAdFreeActive(): boolean {
  return isAdFree()
}

/** Notified whenever the ad-free state changes (purchase, rewarded window, expiry). */
export function onAdFreeGranted(cb: () => void): () => void {
  return subscribePremium(cb)
}

export async function initAds(_lang: Lang): Promise<void> {
  await bootstrapAds()
}

export async function onCommandOpened(): Promise<void> {
  await showInterstitialAfterNavigation()
}

export function watchAdForAdFreeWindow(): Promise<boolean> {
  return showRewardedForAdFree()
}
