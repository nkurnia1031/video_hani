# Short-form portrait video: composition and UI-safe areas

**Researched:** 2026-11-28
**Primary destination:** Instagram Reels
**Scope:** Instagram Reels, TikTok, and YouTube Shorts. The platform guidance below is based on current first-party sources. The proposed storyboard layout is an implementation recommendation, not a platform guarantee.

## Findings from platform sources

### Instagram Reels

Meta's Reels **video-ad** specifications recommend a vertical 9:16 canvas and 1440 × 2560 pixels. They advise leaving approximately 14% at the top, 35% at the bottom, and 6% on each side free of text, logos, and other important creative elements so the Reels interface does not cover them. Meta's guide concerns ads; it does not establish a universal safe-area measurement for every organic Reel or every device UI. [Meta Ads Guide: Instagram Reels video](https://business-auth.facebook.com/business/ads-guide/update/video/instagram-reels/outcome-engagement)

For the storyboard's current 1080 × 1920 canvas, those percentages translate to approximate margins of:

| Area | Share | Approx. pixels |
|---|---:|---:|
| Top | 14% | 269 px |
| Bottom | 35% | 672 px |
| Left and right | 6% each | 65 px |

The resulting central safe rectangle is approximately 950 × 979 px, from x=65 to x=1015 and y=269 to y=1248. The calculation applies Meta's ad recommendation proportionally to the current board canvas; it is not a separate Meta-published pixel specification.

### TikTok

TikTok recommends vertical 9:16 and at least 720p for ad creative, and describes a hook, body, and close as a reliable content structure. Its ad specifications say the safe-zone dimensions depend on the asset orientation, ad-caption length, and interactive add-ons; longer captions reduce the safe zone. TikTok provides matching safe-zone files and recommends previewing the actual placement. There is therefore no single percentage that applies to all TikTok formats. [TikTok Creative Codes](https://ads.tiktok.com/business/en-US/creative-codes), [In-Feed Ads specifications and safe-zone notes](https://ads.tiktok.com/resources/help/article/tiktok-reservation-in-feed-ads-reach-frequency?lang=en-0)

TikTok's small-business video guide recommends 9:16, 1080 × 1920 or higher, keeping text and captions within safe zones, and making content legible without zooming. These are ad-oriented recommendations and should not be mistaken for a fixed organic-post UI template. [TikTok video advertising guide](https://ads.tiktok.com/business/en/guides/video-advertising-guide?redirected=1)

### YouTube Shorts

YouTube accepts square or vertical videos for Shorts; its upload editor can trim footage, and content outside the selected crop will not appear in the Short. [Upload YouTube Shorts](https://support.google.com/youtube/answer/12779649?hl=en)

YouTube's Shorts editor supplies visual guides showing where viewer UI such as Like, comments, and the description may appear. It marks overlays near edges as non-safe and can signal when viewer icons may cover them. The public guidance does not publish one fixed safe-area percentage. [Enhance your Shorts: visual guides](https://support.google.com/youtube/answer/16215842?hl=en)

## Composition recommendation for this storyboard

1. Keep the existing 9:16 composition and design around the Instagram Reels viewer as the primary destination.
2. Use Meta's 14% top / 35% bottom / 6% side guidance as a **conservative reference for critical content**, while labeling it as ad-derived. Treat the top and bottom bands as visually quiet UI reserve areas: background may continue through them, but keep faces, hands doing an action, phones, dialogue, labels, and essential props out of them.
3. Keep key material away from the right edge as well. TikTok and Shorts expose interactive UI there, but their exact placement and safe dimensions vary. Use the platform's current preview/template when exporting to those apps.
4. Keep action and core on-screen text in the central viewing area, with enough separation to read quickly on a phone. Treat each scene as one composition and one page in the storyboard; do not break a single scene across pages.
5. Recheck final exports in the target app's preview. Ad-oriented templates are useful references, but UI, caption length, device crop, and organic/paid placement can change the available area.

Recommendations 2–5 are production decisions inferred from the platform guidance. They are not a guarantee that an organic post will display identical UI on every device.

## Source access

All links above are first-party platform documentation. Accessed 2026-11-28. No universal numeric safe area for organic Reels, TikTok posts, and YouTube Shorts was found; the numeric margins above come specifically from Meta's Reels ad guidance.
