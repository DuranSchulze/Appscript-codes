# FilePino Website — QA & Web Audit Checklist

> **URL:** https://www.filepino.com/  
> **Stack:** WordPress + Elementor  
> **Last Updated:** 2026-02-19

---

## 🐛 Known Issues

- **Responsiveness issue:** Layout breaks at certain resolutions (specifically near tablet sizes, not on mobile phones)

---

### Manual Testing Checklist

#### Responsive Design Testing

- **Tested all breakpoints:** mobile (320px–480px), tablet (768px–1024px), small laptop (1024px–1366px), desktop (1440px+)
  - ✅ Working correctly across all screen sizes
- **Horizontal scrolling:** No unwanted horizontal scroll at any viewport
- **Mobile navigation:** Menu collapses properly with hamburger menu behavior
- **Sticky header:** Fixed header behaves correctly on scroll across all breakpoints
- **Cross-browser compatibility:** Tested on Chrome, Safari, Firefox, Edge (latest versions)
- **Real device testing:** Tested on actual iOS Safari and Android Chrome devices
- **Media scaling:** All images and cards scale without overflow or cropping
- **Touch targets:** Verify touch targets are large enough on mobile (minimum 44×44px)

#### Navigation & Links

- ⚠️ **Compliance Calendar page:** The page at `/compliance-calendar/` appears to be outdated and needs review
- ⚠️ **Download button:** Not functioning properly across pages, particularly in blog sections
- ⚠️ **Navigation links:** Generally working across pages and blog articles

#### Footer

- [ ] **Footer links (Email, Phone, Map, Social Media):** Most links are functional
  - ⚠️ **WhatsApp button:** Not working on mobile screens

#### Mobile Navigation Issue

- ⚠️ **Dropdown menus on mobile:** When users tap on navigation items with dropdown menus, the parent link activates immediately instead of showing the dropdown. This makes it difficult for mobile users to access submenu items. The dropdown toggle should be separate from the link action.

---

## 3. Forms & Contact

- [ ] **Contact form:** Not fully tested, but appears functional based on the volume of client inquiries received
- [ ] **Phone number validation:** Form accepts any characters in the phone number field (should restrict to numbers and valid phone formats)

---

---

# 📊 Website Audit Findings Report

> **Audit Date:** 2026-02-23  
> **Tool Used:** SquirrelScan (squirrelscan.com)  
> **Pages Analyzed:** 100 pages

---

## Executive Summary

A comprehensive automated audit was conducted on the FilePino website to identify issues affecting search engine visibility, user experience, website performance, and accessibility compliance. The audit revealed **significant opportunities for improvement**, with the most critical concerns being:

1. **Accessibility barriers** that may prevent some users from fully accessing the site
2. **Performance bottlenecks** causing slow page loads
3. **Missing trust signals** that affect credibility and search rankings

---

## 📈 Overall Health Score

| Category                          |       Issues Found        | Impact Level |
| --------------------------------- | :-----------------------: | :----------: |
| **Accessibility**                 | 799 errors + 741 warnings | 🔴 Critical  |
| **Performance**                   |  6 errors + 895 warnings  |   🟠 High    |
| **Images**                        |       836 warnings        |   🟠 High    |
| **Links**                         |       212 warnings        |  🟡 Medium   |
| **Content**                       |        56 warnings        |  🟡 Medium   |
| **Structured Data**               |        100 errors         |  🟡 Medium   |
| **SEO**                           |        12 warnings        |    🟢 Low    |
| **URL Structure**                 |        13 warnings        |    🟢 Low    |
| **Trust & Credibility (E-E-A-T)** |        5 warnings         |  🟡 Medium   |

---

## 🔴 Critical Finding: Accessibility Issues

### What This Means

Accessibility issues prevent people with disabilities (using screen readers, keyboard navigation, or other assistive technologies) from using the website effectively. This affects approximately **15-20% of potential visitors** and may expose the business to legal compliance risks.

### Key Problems Identified

| Problem                                       |   How Many?   | Why It Matters                                        |
| --------------------------------------------- | :-----------: | ----------------------------------------------------- |
| **Empty links** (buttons/links with no text)  |     133+      | Screen readers cannot tell users where the link leads |
| **Form fields without proper labels**         |   9 fields    | Users cannot understand what information to enter     |
| **Missing main content area**                 |   98 pages    | Screen reader users cannot skip to main content       |
| **Buttons without descriptive names**         | 6-11 per page | Users don't know what action the button performs      |
| **Hidden elements that can still be focused** | 16 instances  | Keyboard users get "lost" in invisible areas          |

### Real-World Impact

> A blind user navigating with a keyboard and screen reader encounters a button with no label. They have no way of knowing whether it submits a form, opens a menu, or links to another page. They may skip it entirely or accidentally trigger an unwanted action.

---

## 🟠 High Priority: Performance Issues

### What This Means

Performance issues cause the website to load slowly, leading to poor user experience, higher bounce rates (visitors leaving before the page loads), and lower search engine rankings.

### Key Problems Identified

| Problem                     | Current State | Target State         |
| --------------------------- | ------------- | -------------------- |
| **Server response time**    | 650-1,270ms   | Under 200ms          |
| **Total page weight**       | 29,210 KB     | Under 1,000 KB       |
| **Caching**                 | Not enabled   | Should be enabled    |
| **JavaScript optimization** | Not minified  | Should be compressed |

### Real-World Impact

> A potential client visits the website. The page takes 3-4 seconds to load. Research shows that **53% of mobile users abandon sites that take longer than 3 seconds to load**. That visitor may never return.

### Quick Wins

- **Enabling caching** could make repeat visits 2-3x faster
- **Compressing JavaScript files** would save approximately 220 KB per page load
- **Optimizing images** would significantly reduce load times

---

## 🟠 High Priority: Image Issues

### What This Means

Images play a crucial role in both user experience and search engine optimization. Problems with images affect how quickly pages load, how search engines understand the content, and whether all users can access the information.

### Key Problems Identified

| Problem                  |    Count    | Business Impact                                                                  |
| ------------------------ | :---------: | -------------------------------------------------------------------------------- |
| **Missing alt text**     | 488 images  | Search engines cannot understand image content; screen reader users miss context |
| **Oversized images**     |  49 images  | Pages load slower than necessary                                                 |
| **No lazy loading**      | 240+ images | Below-the-fold images load unnecessarily, slowing initial page load              |
| **Images not optimized** |  All pages  | Wasted bandwidth and slower performance                                          |

### Real-World Impact

> A Google Images searcher looking for "business registration Philippines" cannot find FilePino's helpful diagrams because the images lack descriptive alt text. A potential client is lost to a competitor.

---

## 🟡 Medium Priority: Link Issues

### What This Means

Links connect pages together and help search engines understand the website structure. Problems with links can hurt search rankings and create confusing navigation for users.

### Key Problems Identified

| Problem                   |   Count   | Explanation                                                                             |
| ------------------------- | :-------: | --------------------------------------------------------------------------------------- |
| **Weak internal linking** | 43 pages  | Some pages have only 1 link pointing to them, making them hard to discover              |
| **Orphan pages**          | 44 pages  | Pages with almost no incoming links, essentially "hidden" from users and search engines |
| **Empty anchor text**     | 129 links | Links that don't describe where they lead                                               |
| **Redirect chains**       | 19 pages  | Links that go through unnecessary intermediate stops before reaching the destination    |
| **Broken external links** | Multiple  | Links to other websites that no longer work                                             |

---

## 🟡 Medium Priority: Trust & Credibility (E-E-A-T)

### What This Means

Google evaluates websites based on Experience, Expertise, Authoritativeness, and Trustworthiness (E-E-A-T). Missing these elements can hurt search rankings, especially for service businesses.

### Missing Elements

| What's Missing         | Why It Matters                                      |
| ---------------------- | --------------------------------------------------- |
| **About page**         | Visitors want to know who they're dealing with      |
| **Contact page**       | Clients need ways to reach the business             |
| **Privacy Policy**     | Required by law in many jurisdictions; builds trust |
| **Author attribution** | Shows expertise and credibility of content creators |
| **Publication dates**  | Helps visitors assess content freshness             |

### Real-World Impact

> A potential corporate client researches service providers. They visit FilePino but cannot find information about the team, company background, or when articles were written. They choose a competitor with a detailed About page and dated, attributed articles.

---

## 🟡 Medium Priority: Structured Data Issues

### What This Means

Structured data is hidden code that helps search engines understand page content. When properly implemented, it enables rich snippets in search results (star ratings, event dates, FAQ accordions, etc.).

### Current Status

All 100 pages audited have **invalid or incomplete structured data**, meaning the website is missing opportunities to stand out in search results.

### What's Missing

- Company logo in search results
- Publisher information for articles
- Article images in search previews

---

## 🟢 Lower Priority Issues

### Content Quality

- **44 pages** may have excessive repetition of certain keywords (business, visa, tax, bir, etc.), which can appear "spammy" to search engines
- Some pages have **heading structure issues** (skipping levels), making content harder to follow

### URL Structure

- **13 URLs** are longer than 100 characters, which can be difficult for users to share and remember

---

## ✅ Recommended Action Plan

### Phase 1: Critical Fixes (Week 1-2)

| Action                           | Benefit                                      |
| -------------------------------- | -------------------------------------------- |
| Fix empty links and buttons      | All users can navigate the site              |
| Add proper labels to form fields | Forms become usable for everyone             |
| Add main content landmarks       | Screen reader users can navigate efficiently |
| Fix duplicate form field IDs     | Eliminates form confusion                    |

### Phase 2: Performance Improvements (Week 2-4)

| Action                            | Expected Improvement        |
| --------------------------------- | --------------------------- |
| Enable browser caching            | 50-70% faster repeat visits |
| Compress JavaScript files         | ~220 KB saved per page      |
| Optimize image file sizes         | 30-50% faster page loads    |
| Implement lazy loading for images | Faster initial page display |

### Phase 3: SEO & Trust Building (Week 4-6)

| Action                                 | Expected Outcome                                  |
| -------------------------------------- | ------------------------------------------------- |
| Add alt text to all images             | Better search visibility; accessible to all users |
| Create About, Contact, Privacy pages   | Improved trust and credibility                    |
| Add author names and dates to articles | Demonstrates expertise; content appears fresher   |
| Fix structured data                    | Rich snippets in search results                   |
| Improve internal linking               | Better crawlability; users discover more content  |

### Phase 4: Ongoing Maintenance

- Regular link audits to catch broken links early
- Image optimization as part of content workflow
- Periodic accessibility testing as new content is added

---

## 📋 Summary Checklist

### Must Fix Now

- [ ] 133+ empty links need descriptive text
- [ ] 9 form fields need proper labels
- [ ] 98 pages need main content area defined
- [ ] Duplicate form IDs need to be made unique

### Should Fix Soon

- [ ] 488 images need alt text descriptions
- [ ] Enable caching on all pages
- [ ] Compress JavaScript files
- [ ] 49 oversized images need optimization

### Plan to Fix

- [ ] Create About, Contact, and Privacy Policy pages
- [ ] Add author names to articles
- [ ] Add publication dates to content
- [ ] Fix broken external links
- [ ] Improve internal linking structure

---
