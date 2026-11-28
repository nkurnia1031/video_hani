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
2. Meta's 14% top / 35% bottom / 6% side guidance remains the ad-derived reference for important content. For this storyboard, the requested composition keeps the top 14% and bottom 20% background-only. At 1080 × 1920 those bands are y=0–269 and y=1536–1920.
3. The remaining 15% between y=1248 and y=1536 is an **illustration-only band**: it can hold secondary people, furniture, or props, but no captions, dialogue, labels, or key messages. Important text remains in x=65–1015 and y=269–1248. That extra band may intersect app UI in some placements, so content there must stay nonessential.
4. Keep key material away from the right edge as well. TikTok and Shorts expose interactive UI there, but their exact placement and safe dimensions vary. Use the platform's current preview/template when exporting to those apps.
5. Keep action and core on-screen text in the central viewing area, with enough separation to read quickly on a phone. Treat each scene as one composition and one page in the storyboard; do not break a single scene across pages.
6. Recheck final exports in the target app's preview. Ad-oriented templates are useful references, but UI, caption length, device crop, and organic/paid placement can change the available area.

### Board's selected layout bands

| Region on 1080 × 1920 frame | Bounds | Content allowed |
|---|---:|---|
| Top reserve · 14% | y=0–269 | Background only |
| Main text-safe area | x=65–1015, y=269–1248 | Faces, actions, phones, dialogue, labels, and key messages |
| Lower illustration band · 15% | y=1248–1536 | Secondary illustration and props only; no text or key messages |
| Bottom UI reserve · 20% | y=1536–1920 | Background only |
| Side margins · 6% each | x=0–65 and x=1015–1080 | Background only |

This layout follows the requested 20% empty lower reserve and uses the extra 15% only for nonessential illustration. It is a production choice, not Meta's published 35% bottom recommendation. The final composition still needs an in-app placement preview, especially when repurposed to paid Reels, TikTok, or Shorts.

Recommendations 2–5 are production decisions inferred from the platform guidance. They are not a guarantee that an organic post will display identical UI on every device.

## Source access

All links above are first-party platform documentation. Accessed 2026-11-28. No universal numeric safe area for organic Reels, TikTok posts, and YouTube Shorts was found; the numeric margins above come specifically from Meta's Reels ad guidance.
