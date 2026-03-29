"""Markdown report generator for SEO analysis."""

from datetime import datetime
from typing import List, Dict, Optional
from pathlib import Path

from crawler.page_parser import PageData


class MarkdownReporter:
    def __init__(self, base_url: str):
        self.base_url = base_url
        self.timestamp = datetime.now()
        self.report_lines = []
    
    def generate_report(
        self,
        pages: List[PageData],
        gsc_data: Dict[str, Dict],
        ga4_data: Dict[str, Dict],
        ai_recommendations: Dict[str, Dict],
        gsc_site_summary: Optional[Dict] = None,
        ga4_site_summary: Optional[Dict] = None,
        ai_provider: str = "",
        keyword_opportunities: Optional[List[Dict]] = None
    ) -> str:
        """Generate full Markdown report."""
        
        self.report_lines = []
        
        self._add_header()
        self._add_executive_summary(
            pages, 
            gsc_site_summary, 
            ga4_site_summary,
            bool(gsc_data),
            bool(ga4_data),
            ai_provider,
            bool(keyword_opportunities)
        )
        self._add_priority_actions(pages, gsc_data, ai_recommendations)
        
        if keyword_opportunities:
            self._add_keyword_opportunities(keyword_opportunities)
        
        self._add_page_analysis(pages, gsc_data, ga4_data, ai_recommendations)
        
        return "\n".join(self.report_lines)
    
    def _add_header(self):
        """Add report header."""
        self.report_lines.extend([
            f"# SEO Analysis Report: {self.base_url}",
            "",
            f"**Generated:** {self.timestamp.strftime('%Y-%m-%d %H:%M:%S')}",
            "",
            "---",
            ""
        ])
    
    def _add_executive_summary(
        self,
        pages: List[PageData],
        gsc_summary: Optional[Dict],
        ga4_summary: Optional[Dict],
        has_gsc: bool,
        has_ga4: bool,
        ai_provider: str,
        has_keyword_metrics: bool = False
    ):
        """Add executive summary section."""
        
        total_pages = len(pages)
        critical_issues = sum(1 for p in pages if not p.meta_description or not p.title)
        warnings = sum(1 for p in pages if len(p.title) > 60 or len(p.meta_description) > 160)
        
        self.report_lines.extend([
            "## Executive Summary",
            "",
            f"| Metric | Value |",
            f"|--------|-------|",
            f"| **Pages Analyzed** | {total_pages} |",
            f"| **Critical Issues** | {critical_issues} |",
            f"| **Warnings** | {warnings} |",
            f"| **Google Search Console** | {'✅ Connected' if has_gsc else '❌ Not connected'} |",
            f"| **Google Analytics 4** | {'✅ Connected' if has_ga4 else '❌ Not connected'} |",
            f"| **AI Provider** | {ai_provider or 'Not used'} |",
            f"| **Keyword Metrics** | {'✅ Available' if has_keyword_metrics else '❌ Not available'} |",
            ""
        ])
        
        if gsc_summary:
            self.report_lines.extend([
                "### Search Console Overview (Last 28 Days)",
                "",
                f"| Metric | Value |",
                f"|--------|-------|",
                f"| Total Clicks | {gsc_summary.get('clicks', 0):,} |",
                f"| Total Impressions | {gsc_summary.get('impressions', 0):,} |",
                f"| Average CTR | {gsc_summary.get('ctr', 0):.2%} |",
                f"| Average Position | {gsc_summary.get('avg_position', 0):.1f} |",
                ""
            ])
        
        if ga4_summary:
            self.report_lines.extend([
                "### Analytics Overview (Last 28 Days)",
                "",
                f"| Metric | Value |",
                f"|--------|-------|",
                f"| Total Sessions | {ga4_summary.get('sessions', 0):,} |",
                f"| Total Users | {ga4_summary.get('users', 0):,} |",
                f"| New Users | {ga4_summary.get('new_users', 0):,} |",
                f"| Page Views | {ga4_summary.get('page_views', 0):,} |",
                ""
            ])
        
        self.report_lines.append("---\n")
    
    def _add_keyword_opportunities(self, keywords: List[Dict]):
        """Add keyword opportunities section with GSC + Ads data."""
        
        self.report_lines.extend([
            "## 🎯 Keyword Opportunities",
            "",
            f"**Top {len(keywords)} keywords from Search Console, enriched with Google Ads metrics:**",
            "",
            "*Opportunity Score = Based on current position, search volume, and competition*",
            "",
        ])
        
        from keyword_metrics_integration import format_keyword_opportunities_table
        
        table = format_keyword_opportunities_table(keywords)
        self.report_lines.append(table)
        
        # Add quick wins section
        quick_wins = [k for k in keywords if 5 <= k.get('gsc_avg_position', 100) <= 15 and k.get('gsc_impressions', 0) > 100]
        
        if quick_wins:
            self.report_lines.extend([
                "",
                "### 🚀 Quick Win Keywords (Positions 5-15)",
                "",
                "These keywords are ranking well but with small improvements could drive significantly more traffic:",
                "",
            ])
            
            for kw in quick_wins[:10]:
                self.report_lines.append(
                    f"- **{kw['keyword']}** - Position {kw['gsc_avg_position']:.1f}, "
                    f"{kw['gsc_impressions']:,} impressions, "
                    f"{kw.get('search_volume', 0):,} monthly searches"
                )
        
        # Add high-opportunity keywords
        high_opp = [k for k in keywords if k.get('opportunity_score', 0) >= 70]
        if high_opp:
            self.report_lines.extend([
                "",
                "### 💎 High Opportunity Keywords (Score 70+)",
                "",
            ])
            
            for kw in high_opp[:10]:
                self.report_lines.append(
                    f"- **{kw['keyword']}** - Score: {kw.get('opportunity_score', 0)}/100, "
                    f"Position: {kw['gsc_avg_position']:.1f}, "
                    f"Competition: {kw.get('competition', 'N/A')}"
                )
        
        self.report_lines.append("\n---\n")
    
    def _add_priority_actions(
        self,
        pages: List[PageData],
        gsc_data: Dict[str, Dict],
        ai_recommendations: Dict[str, Dict]
    ):
        """Add priority actions section."""
        
        actions = []
        
        for page in pages:
            gsc = gsc_data.get(page.url, {})
            
            if not page.meta_description and gsc.get('impressions', 0) > 100:
                actions.append({
                    "priority": "HIGH",
                    "page": page.url,
                    "issue": "Missing meta description",
                    "impact": f"{gsc.get('impressions', 0):,} impressions but 0% CTR potential"
                })
            
            if len(page.title) > 60:
                actions.append({
                    "priority": "MEDIUM",
                    "page": page.url,
                    "issue": f"Title too long ({len(page.title)} chars)",
                    "impact": "Title truncated in search results"
                })
            
            if not page.h1:
                actions.append({
                    "priority": "MEDIUM",
                    "page": page.url,
                    "issue": "Missing H1 heading",
                    "impact": "Poor content structure"
                })
            
            if gsc.get('avg_position', 100) < 15 and gsc.get('ctr', 0) < 0.02:
                actions.append({
                    "priority": "HIGH",
                    "page": page.url,
                    "issue": "Low CTR despite good position",
                    "impact": f"Position {gsc.get('avg_position', 0):.1f} but only {gsc.get('ctr', 0):.2%} CTR"
                })
        
        actions.sort(key=lambda x: (x["priority"] != "HIGH", x["priority"]))
        
        self.report_lines.extend([
            "## Priority Actions",
            "",
            f"**Found {len(actions)} items requiring attention:**",
            ""
        ])
        
        for i, action in enumerate(actions[:15], 1):
            emoji = "🔴" if action["priority"] == "HIGH" else "🟡"
            self.report_lines.extend([
                f"{emoji} **{i}. [{action['priority']}]** {action['issue']}",
                f"   - Page: `{action['page']}`",
                f"   - Impact: {action['impact']}",
                ""
            ])
        
        if len(actions) > 15:
            self.report_lines.append(f"*... and {len(actions) - 15} more issues*\n")
        
        self.report_lines.append("---\n")
    
    def _add_page_analysis(
        self,
        pages: List[PageData],
        gsc_data: Dict[str, Dict],
        ga4_data: Dict[str, Dict],
        ai_recommendations: Dict[str, Dict]
    ):
        """Add detailed page analysis section."""
        
        self.report_lines.extend([
            "## Page Analysis",
            ""
        ])
        
        for page in pages:
            self._add_page_section(page, gsc_data.get(page.url), ga4_data.get(page.url), ai_recommendations.get(page.url))
        
        self.report_lines.append("---\n")
    
    def _add_page_section(
        self,
        page: PageData,
        gsc: Optional[Dict],
        ga4: Optional[Dict],
        ai_rec: Optional[Dict]
    ):
        """Add a single page section."""
        
        from urllib.parse import urlparse
        path = urlparse(page.url).path or "/"
        
        status_icon = "✅" if page.status_code == 200 else "❌"
        
        self.report_lines.extend([
            f"### {status_icon} {path}",
            "",
            "#### On-Page SEO",
            "",
            f"| Element | Value | Status |",
            f"|---------|-------|--------|",
        ])
        
        title_status = "✅" if page.title and 20 <= len(page.title) <= 60 else "⚠️"
        desc_status = "✅" if page.meta_description and 50 <= len(page.meta_description) <= 160 else "⚠️"
        h1_status = "✅" if page.h1 else "❌"
        
        self.report_lines.extend([
            f"| Title | {self._truncate(page.title, 50)} ({len(page.title)} chars) | {title_status} |",
            f"| Meta Description | {self._truncate(page.meta_description, 40)} ({len(page.meta_description)} chars) | {desc_status} |",
            f"| H1 | {self._truncate(page.h1, 50)} | {h1_status} |",
            f"| Status | {page.status_code} | {'✅' if page.status_code == 200 else '❌'} |",
            f"| Word Count | {page.word_count:,} | - |",
            f"| Internal Links | {len(page.internal_links)} | - |",
            f"| Images w/o Alt | {page.images_without_alt} | {'⚠️' if page.images_without_alt > 0 else '✅'} |",
            f"| Schema Markup | {'✅' if page.has_schema_markup else '❌'} | {'✅' if page.has_schema_markup else '-'} |",
            ""
        ])
        
        if gsc:
            self.report_lines.extend([
                "#### Google Search Console (Last 28 Days)",
                "",
                f"| Metric | Value |",
                f"|--------|-------|",
                f"| Clicks | {gsc.get('clicks', 0):,} |",
                f"| Impressions | {gsc.get('impressions', 0):,} |",
                f"| CTR | {gsc.get('ctr', 0):.2%} |",
                f"| Avg Position | {gsc.get('avg_position', 0):.1f} |",
                ""
            ])
            
            if gsc.get('top_queries'):
                self.report_lines.extend([
                    "**Top Queries:**",
                    ""
                ])
                for q in gsc['top_queries'][:3]:
                    self.report_lines.append(f"- `{q['query']}`: {q['clicks']} clicks, position {q['position']:.1f}")
                self.report_lines.append("")
        
        if ga4:
            self.report_lines.extend([
                "#### Google Analytics 4 (Last 28 Days)",
                "",
                f"| Metric | Value |",
                f"|--------|-------|",
                f"| Sessions | {ga4.get('sessions', 0):,} |",
                f"| Users | {ga4.get('users', 0):,} |",
                f"| Page Views | {ga4.get('page_views', 0):,} |",
                f"| Engagement Rate | {ga4.get('engagement_rate', 0):.2%} |",
                ""
            ])
        
        if ai_rec and not ai_rec.get("error"):
            self.report_lines.extend([
                "#### 🤖 AI Keyword Recommendations",
                "",
            ])
            
            if ai_rec.get("primary_keyword"):
                self.report_lines.append(f"**Primary Keyword:** {ai_rec['primary_keyword']}")
            
            if ai_rec.get("secondary_keywords"):
                self.report_lines.append(f"**Secondary:** {', '.join(ai_rec['secondary_keywords'][:4])}")
            
            if ai_rec.get("search_intent"):
                self.report_lines.append(f"**Intent:** {ai_rec['search_intent']}")
            
            if ai_rec.get("content_gaps"):
                self.report_lines.append(f"**Content Gap:** {ai_rec['content_gaps']}")
            
            if ai_rec.get("optimization_tip"):
                self.report_lines.append(f"**Tip:** {ai_rec['optimization_tip']}")
            
            self.report_lines.append("")
        
        self.report_lines.append("---\n")
    
    def _truncate(self, text: str, length: int) -> str:
        """Truncate text for display."""
        if not text:
            return "-"
        if len(text) <= length:
            return text
        return text[:length - 3] + "..."
    
    def save_report(self, output_path: Optional[Path] = None) -> Path:
        """Save report to file."""
        if output_path is None:
            timestamp = self.timestamp.strftime('%Y%m%d-%H%M%S')
            output_path = Path(f"seo-report-{timestamp}.md")
        
        output_path.write_text("\n".join(self.report_lines), encoding='utf-8')
        return output_path
