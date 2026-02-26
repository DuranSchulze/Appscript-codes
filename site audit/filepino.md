# FilePino Website — QA & Web Audit Checklist

> **URL:** https://www.filepino.com/  
> **Stack:** WordPress + Elementor  
> **Last Updated:** 2026-02-19

---

## 🐛 Known Issues

- Responsiveness breaks on low resolution (not mobile phone sizes but near tablet sizes)



- [x] Test all breakpoints: mobile (320px–480px), tablet (768px–1024px), small laptop (1024px–1366px), desktop (1440px+)
  - All Working in each screen sizes
- [x] Check horizontal scrolling doesn't appear at any viewport
- [x] Verify navigation menu collapses properly on mobile (hamburger menu behavior)
- [x] Test sticky/fixed header behavior on scroll across all breakpoints
- [x] Cross-browser testing: Chrome, Safari, Firefox, Edge (latest versions)
- [x] Test on actual iOS Safari and Android Chrome devices
- [x] Verify all images and cards scale without overflow or cropping issues
- [ ] Check touch targets are large enough on mobile (minimum 44×44px)
- https://www.filepino.com/compliance-calendar/ what is this page and this seems outdated.
- Navigations and links seems working across ther pages especially in blogs. Download Button is not working too.
- [ ] Footer links (Email, Phone, Map, Social Media) are working, But in thw footer for socials, The whatsapp button is not working tested in mobile screen
- Navigation bar seems good and working for desktop, BUT in mobile its very hard because some navigation dropdowns when pressed already opens another link, It seems like the dropdown button directly have a link making it hard for users to navigate the links inside the dropdown.



---


## 3. Forms & Contact

- [ ] Contact form , I didnt try to test it and maybe its working based on the volume of requests or clients we are getting.
- [ ] Form validation works but not on Phone Number. It accepts any characters.

---

## 4. Performance & Page Speed

- [ ] Run **Google PageSpeed Insights** (target: 90+ on mobile, 90+ on desktop)
- [ ] Run **GTmetrix** analysis
- [ ] Check **Largest Contentful Paint (LCP)** < 2.5s
- [ ] Check **First Input Delay (FID)** / **Interaction to Next Paint (INP)** < 200ms
- [ ] Check **Cumulative Layout Shift (CLS)** < 0.1
- [ ] Images are optimized (WebP format, lazy loading, proper sizing)
- [ ] CSS and JS files are minified and bundled
- [ ] Unused CSS/JS is eliminated or deferred
- [ ] Check for render-blocking resources
- [ ] Browser caching headers are properly set
- [ ] Enable GZIP/Brotli compression
- [ ] Check total page weight (target: under 3MB)

---

## 5. SEO Audit

- [ ] Each page has a unique `<title>` tag (50–60 characters)
- [ ] Each page has a unique `<meta description>` (150–160 characters)
- [ ] Proper heading hierarchy (single `<h1>` per page, logical `<h2>`–`<h6>`)
- [ ] All images have descriptive `alt` attributes
- [ ] URL structure is clean and SEO-friendly (no query strings, lowercase)
- [ ] **Canonical tags** are set on all pages
- [ ] **Open Graph (OG)** tags are present for social sharing
- [ ] **Twitter Card** meta tags are present
- [ ] `robots.txt` exists and is correctly configured
- [ ] `sitemap.xml` exists and is submitted to Google Search Console
- [ ] Structured data / JSON-LD schema is implemented (e.g., LocalBusiness, Organization)
- [ ] No duplicate content across pages
- [ ] Internal linking strategy is logical and well-connected
- [ ] 301 redirects are in place for old/changed URLs
- [ ] Google Search Console is set up and monitoring

---

## 6. Accessibility (WCAG 2.1)

- [ ] Color contrast ratios meet WCAG AA (minimum 4.5:1 for text)
- [ ] All interactive elements are keyboard-navigable (Tab, Enter, Escape)
- [ ] Focus indicators are visible on all focusable elements
- [ ] ARIA labels on icons, buttons, and non-text elements
- [ ] Skip-to-content link is available
- [ ] Forms have proper `<label>` elements associated to inputs
- [ ] Error messages are announced to screen readers
- [ ] No content relies solely on color to convey meaning
- [ ] Video/media has captions or transcripts (if applicable)
- [ ] Run **WAVE** or **axe DevTools** accessibility audit

---

## 7. Security

- [ ] SSL certificate is valid and active (HTTPS everywhere)
- [ ] HTTP → HTTPS redirect is enforced
- [ ] Mixed content warnings are resolved (no HTTP resources on HTTPS pages)
- [ ] WordPress is on the latest version
- [ ] All plugins and themes are updated
- [ ] Admin login URL is not exposed at default `/wp-admin` or `/wp-login.php`
- [ ] Strong password policy for admin accounts
- [ ] Security headers are set:
  - [ ] `X-Content-Type-Options: nosniff`
  - [ ] `X-Frame-Options: SAMEORIGIN`
  - [ ] `Content-Security-Policy`
  - [ ] `Strict-Transport-Security` (HSTS)
  - [ ] `Referrer-Policy`
- [ ] Cookie consent banner works correctly (Accept All, Reject All, Customise)
- [ ] File upload restrictions are in place (if applicable)
- [ ] Rate limiting on forms to prevent abuse

---

## 8. Content & Copy

- [ ] No spelling or grammar errors on any page
- [ ] All placeholder/lorem ipsum text is replaced with real content
- [ ] Contact information is accurate and consistent across all pages
  - [ ] Email: info@filepino.com
  - [ ] Phone: +63 917 892 2337, (02) 7116 3477, (02) 7000 7500
  - [ ] Address: 1212 High Street South Corporate Plaza Tower 2, BGC, Taguig
- [ ] Client testimonials section displays properly and content is real
- [ ] Copyright year in footer is current (2026)
- [ ] Privacy Policy and Terms of Service pages exist and are linked
- [ ] Blog/articles (if any) have proper dates and author attribution

---

## 9. Cookie Consent & Legal Compliance

- [ ] Cookie consent banner appears on first visit
- [ ] "Accept All" and "Reject All" buttons work correctly
- [ ] "Customise" option allows granular cookie control
- [ ] Non-essential cookies are blocked until consent is given
- [ ] Privacy Policy page is accessible and up to date
- [ ] Cookie policy details which cookies are used and why
- [ ] Consent preference is remembered on return visits
- [ ] Compliant with **Philippine Data Privacy Act (RA 10173)**

---

## 10. Analytics & Tracking

- [ ] Google Analytics (GA4) is installed and tracking
- [ ] Google Tag Manager is properly configured (if used)
- [ ] Conversion tracking is set up for form submissions and CTAs
- [ ] Facebook Pixel / Meta Pixel is installed (if used for ads)
- [ ] No duplicate tracking scripts
- [ ] Tracking respects cookie consent preferences

---

## 11. WordPress / Elementor Specific

- [ ] Elementor editor does not leave visual artifacts on the frontend
- [ ] No broken Elementor widgets or missing plugin warnings
- [ ] WordPress debug mode is OFF on production (`WP_DEBUG = false`)
- [ ] Caching plugin is active and configured (e.g., WP Rocket, LiteSpeed Cache)
- [ ] Database is optimized (clean post revisions, transients, spam comments)
- [ ] Unused themes and plugins are deactivated and removed
- [ ] WordPress REST API is restricted if not needed
- [ ] XML-RPC is disabled if not used (security risk)
- [ ] Automatic updates are configured for minor releases

---

## 12. Email Deliverability

- [ ] Form submission emails are not going to spam
- [ ] SMTP plugin is configured (not using default `wp_mail()`)
- [ ] SPF, DKIM, and DMARC records are set for the domain
- [ ] Reply-to address is correct on automated emails
- [ ] Email templates are branded and professional

---

## 13. Backup & Disaster Recovery

- [ ] Automated backups are scheduled (daily or weekly)
- [ ] Backup includes both files and database
- [ ] Backups are stored off-site (not just on the hosting server)
- [ ] Backup restoration has been tested at least once

---

## 14. Uptime & Monitoring

- [ ] Uptime monitoring is active (e.g., UptimeRobot, Pingdom)
- [ ] Alerts are configured for downtime
- [ ] Server error logs are being monitored
- [ ] 5xx error pages are custom and user-friendly
- [ ] 404 error page is custom and provides navigation back to the site


