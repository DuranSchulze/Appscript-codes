"""SEO Scanner - Main CLI entry point.

A Python tool that crawls websites, analyzes SEO, fetches Google data,
and generates AI-powered keyword recommendations.
"""

from pathlib import Path
from typing import Optional
from datetime import datetime
import time

import typer
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn
from rich.panel import Panel
from rich.text import Text

import config
from crawler.site_crawler import SiteCrawler
from google.search_console import SearchConsoleClient
from google.ga4 import GA4Client
from google.auth import clear_credentials
from ai.keyword_recommender import KeywordRecommender
from reporter.markdown_reporter import MarkdownReporter
from keyword_metrics_integration import (
    KeywordMetricsIntegration,
    format_keyword_opportunities_table,
    get_quick_wins
)

app = typer.Typer(help="SEO Scanner - Analyze website SEO with AI recommendations")
console = Console()


@app.command()
def scan(
    url: str = typer.Argument(..., help="Website URL to scan (e.g., https://example.com)"),
    max_pages: int = typer.Option(100, "--max-pages", "-p", help="Maximum pages to crawl"),
    max_depth: int = typer.Option(2, "--max-depth", "-d", help="Maximum crawl depth"),
    output: Optional[Path] = typer.Option(None, "--output", "-o", help="Output file path (default: auto-generated)"),
    use_gsc: bool = typer.Option(False, "--gsc", help="Connect to Google Search Console"),
    use_ga4: bool = typer.Option(False, "--ga4", help="Connect to Google Analytics 4"),
    use_keyword_metrics: bool = typer.Option(False, "--keyword-metrics", "-km", help="Enrich GSC keywords with search volume data"),
    keyword_provider: str = typer.Option("google-ads", "--keyword-provider", "-kp", help="Data provider: google-ads, semrush, or both"),
    ai_provider: Optional[str] = typer.Option(None, "--ai-provider", help="AI provider: gemini or openai"),
    no_ai: bool = typer.Option(False, "--no-ai", help="Skip AI keyword recommendations"),
    delay: float = typer.Option(0.5, "--delay", help="Delay between requests (seconds)"),
):
    """Scan a website and generate SEO analysis report."""
    
    console.print(Panel.fit(
        Text("🔍 SEO Scanner", style="bold blue"),
        subtitle="Analyzing: " + url
    ))
    
    # Phase 1: Crawl website
    console.print("\n[bold cyan]Phase 1: Crawling website...[/bold cyan]")
    
    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        console=console,
    ) as progress:
        task = progress.add_task("Crawling pages...", total=None)
        
        crawler = SiteCrawler(
            base_url=url,
            max_pages=max_pages,
            max_depth=max_depth,
            delay=delay
        )
        pages = crawler.crawl()
        progress.update(task, completed=True)
    
    console.print(f"[green]✓[/green] Crawled {len(pages)} pages")
    
    # Phase 2: Connect to Google APIs
    gsc_client = None
    ga4_client = None
    gsc_data = {}
    ga4_data = {}
    gsc_summary = None
    ga4_summary = None
    
    if use_gsc:
        console.print("\n[bold cyan]Phase 2a: Connecting to Google Search Console...[/bold cyan]")
        gsc_client = SearchConsoleClient()
        if gsc_client.connect():
            console.print("[green]✓[/green] Search Console connected")
            
            with Progress(SpinnerColumn(), TextColumn("[progress.description]{task.description}"), console=console) as progress:
                task = progress.add_task("Fetching GSC data...", total=None)
                
                gsc_summary = gsc_client.get_site_summary()
                for page in pages:
                    data = gsc_client.get_search_analytics(page.url)
                    if data:
                        gsc_data[page.url] = data
                    time.sleep(0.2)
                progress.update(task, completed=True)
            
            console.print(f"[green]✓[/green] Fetched GSC data for {len(gsc_data)} pages")
        else:
            console.print("[yellow]⚠[/yellow] Failed to connect to Search Console")
    
    if use_ga4:
        console.print("\n[bold cyan]Phase 2b: Connecting to Google Analytics 4...[/bold cyan]")
        ga4_client = GA4Client()
        if ga4_client.connect():
            console.print("[green]✓[/green] GA4 connected")
            
            with Progress(SpinnerColumn(), TextColumn("[progress.description]{task.description}"), console=console) as progress:
                task = progress.add_task("Fetching GA4 data...", total=None)
                
                ga4_summary = ga4_client.get_site_summary()
                for page in pages:
                    from urllib.parse import urlparse
                    path = urlparse(page.url).path or "/"
                    data = ga4_client.get_page_metrics(path)
                    if data:
                        ga4_data[page.url] = data
                    time.sleep(0.2)
                progress.update(task, completed=True)
            
            console.print(f"[green]✓[/green] Fetched GA4 data for {len(ga4_data)} pages")
        else:
            console.print("[yellow]⚠[/yellow] Failed to connect to GA4")
    
    # Phase 3: AI Keyword Recommendations
    ai_recommendations = {}
    ai_provider_used = ""
    
    if not no_ai:
        provider = ai_provider or config.DEFAULT_AI_PROVIDER
        ai_provider_used = provider
        
        console.print(f"\n[bold cyan]Phase 3: Generating AI keyword recommendations ({provider})...[/bold cyan]")
        
        try:
            recommender = KeywordRecommender(provider)
            
            if recommender.test_connection():
                console.print(f"[green]✓[/green] {provider} API connected")
                
                with Progress(SpinnerColumn(), TextColumn("[progress.description]{task.description}"), console=console) as progress:
                    task = progress.add_task("Generating recommendations...", total=len(pages))
                    
                    for i, page in enumerate(pages):
                        current_queries = None
                        if page.url in gsc_data:
                            current_queries = gsc_data[page.url].get('top_queries', [])
                        
                        rec = recommender.get_recommendations(
                            page_title=page.title,
                            meta_description=page.meta_description,
                            h1=page.h1,
                            content_snippet=page.content_snippet,
                            current_queries=current_queries
                        )
                        ai_recommendations[page.url] = rec
                        progress.update(task, advance=1)
                        time.sleep(0.5)
                
                console.print(f"[green]✓[/green] Generated recommendations for {len(ai_recommendations)} pages")
            else:
                console.print(f"[yellow]⚠[/yellow] Could not connect to {provider} API. Check your API key.")
                ai_provider_used = ""
                
        except Exception as e:
            console.print(f"[red]✗[/red] AI recommendation error: {e}")
            ai_provider_used = ""
    
    # Phase 3.5: Keyword Metrics Integration (GSC + Search Volume Data)
    keyword_opportunities = []
    keyword_metrics_available = False
    
    if use_keyword_metrics and use_gsc and gsc_client and gsc_client.service:
        console.print(f"\n[bold cyan]Phase 3.5: Fetching keyword metrics (via {keyword_provider})...[/bold cyan]")
        
        # Get keywords from GSC first
        try:
            integration = KeywordMetricsIntegration(
                search_console_client=gsc_client,
                ads_client=None  # We'll handle this manually
            )
            
            with Progress(SpinnerColumn(), TextColumn("[progress.description]{task.description}"), console=console) as progress:
                task = progress.add_task("Extracting keywords from Search Console...", total=None)
                
                gsc_keywords = integration.extract_keywords_from_gsc(
                    days=28,
                    min_impressions=10,
                    limit=100
                )
                progress.update(task, completed=True)
            
            if not gsc_keywords:
                console.print("[yellow]⚠[/yellow] No keywords found in Search Console")
            else:
                console.print(f"[green]✓[/green] Found {len(gsc_keywords)} keywords in Search Console")
                
                # Enrich with provider data
                if keyword_provider in ["google-ads", "both"] and config.GOOGLE_ADS_DEVELOPER_TOKEN:
                    try:
                        from google.ads_keywords import GoogleAdsKeywordClient
                        
                        ads_client = GoogleAdsKeywordClient()
                        if ads_client.connect():
                            integration.ads_client = ads_client
                            
                            with Progress(SpinnerColumn(), TextColumn("[progress.description]{task.description}"), console=console) as progress:
                                task = progress.add_task("Enriching with Google Ads data...", total=None)
                                
                                keyword_opportunities = integration.enrich_keywords_with_ads_metrics(
                                    gsc_keywords,
                                    location_id="2840",
                                    language_id="1000"
                                )
                                progress.update(task, completed=True)
                            
                            if keyword_opportunities:
                                console.print(f"[green]✓[/green] Enriched {len(keyword_opportunities)} keywords with Google Ads data")
                                keyword_metrics_available = True
                    except Exception as e:
                        console.print(f"[red]✗[/red] Google Ads enrichment error: {e}")
                
                # SEMrush enrichment
                if keyword_provider in ["semrush", "both"] and config.SEMRUSH_API_KEY:
                    try:
                        from google.semrush_keywords import SEMrushKeywordClient
                        
                        semrush_client = SEMrushKeywordClient()
                        if semrush_client.test_connection():
                            keyword_list = [k['keyword'] for k in gsc_keywords]
                            
                            with Progress(SpinnerColumn(), TextColumn("[progress.description]{task.description}"), console=console) as progress:
                                task = progress.add_task("Enriching with SEMrush data...", total=None)
                                
                                semrush_metrics = semrush_client.get_keyword_metrics(
                                    keywords=keyword_list,
                                    database="us"
                                )
                                progress.update(task, completed=True)
                            
                            if semrush_metrics:
                                console.print(f"[green]✓[/green] Enriched {len(semrush_metrics)} keywords with SEMrush data")
                                
                                # Merge SEMrush data with GSC data
                                semrush_lookup = {m['keyword']: m for m in semrush_metrics}
                                for k in gsc_keywords:
                                    if k['keyword'] in semrush_lookup:
                                        sem_data = semrush_lookup[k['keyword']]
                                        k['search_volume'] = sem_data['search_volume']
                                        k['cpc'] = sem_data['cpc']
                                        k['competition'] = sem_data['competition_level']
                                        k['source'] = 'semrush'
                                
                                if not keyword_opportunities:
                                    keyword_opportunities = gsc_keywords
                                keyword_metrics_available = True
                    except Exception as e:
                        console.print(f"[red]✗[/red] SEMrush enrichment error: {e}")
                
                # If no enrichment happened, use GSC data only
                if not keyword_opportunities and gsc_keywords:
                    keyword_opportunities = gsc_keywords
                    keyword_metrics_available = True
                    console.print("[green]✓[/green] Using Search Console data (no search volume enrichment)")
                
                # Show quick wins
                if keyword_opportunities:
                    quick_wins = get_quick_wins(keyword_opportunities)
                    if quick_wins:
                        console.print(f"[green]→[/green] Found {len(quick_wins)} 'quick win' keywords (positions 5-15)")
                    
        except Exception as e:
            console.print(f"[red]✗[/red] Keyword metrics error: {e}")
    
    # Phase 4: Generate Report
    console.print("\n[bold cyan]Phase 4: Generating Markdown report...[/bold cyan]")
    
    reporter = MarkdownReporter(base_url=url)
    report_content = reporter.generate_report(
        pages=pages,
        gsc_data=gsc_data,
        ga4_data=ga4_data,
        ai_recommendations=ai_recommendations,
        gsc_site_summary=gsc_summary,
        ga4_site_summary=ga4_summary,
        ai_provider=ai_provider_used,
        keyword_opportunities=keyword_opportunities
    )
    
    # Save report
    output_path = output
    if output_path is None:
        from datetime import datetime
        timestamp = datetime.now().strftime('%Y%m%d-%H%M%S')
        url_slug = url.replace('https://', '').replace('http://', '').replace('/', '_')[:30]
        output_path = config.OUTPUT_DIR / f"seo-report-{url_slug}-{timestamp}.md"
    else:
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
    
    output_path.write_text(report_content, encoding='utf-8')
    
    console.print(f"[green]✓[/green] Report saved to: {output_path}")
    
    # Summary
    console.print("\n" + "=" * 60)
    console.print("[bold green]Scan Complete![/bold green]")
    console.print(f"Pages analyzed: {len(pages)}")
    console.print(f"GSC data: {'✅' if gsc_data else '❌'}")
    console.print(f"GA4 data: {'✅' if ga4_data else '❌'}")
    console.print(f"AI recommendations: {'✅' if ai_recommendations else '❌'}")
    console.print(f"Keyword metrics: {'✅' if keyword_metrics_available else '❌'}")
    console.print(f"Report: [cyan]{output_path}[/cyan]")


@app.command()
def auth_setup():
    """Set up Google API authentication."""
    console.print("[bold cyan]Google API Authentication Setup[/bold cyan]\n")
    
    if not config.GOOGLE_CLIENT_SECRET_FILE.exists():
        console.print("[red]Error:[/red] client_secret.json not found!")
        console.print(f"Please download your OAuth2 credentials from Google Cloud Console")
        console.print(f"and save them to: {config.GOOGLE_CLIENT_SECRET_FILE}")
        raise typer.Exit(1)
    
    console.print("This will open a browser for Google authentication.\n")
    
    # Test Search Console auth
    console.print("[bold]Testing Search Console authentication...[/bold]")
    try:
        from google.auth import get_search_console_credentials
        creds = get_search_console_credentials()
        console.print("[green]✓[/green] Search Console authentication successful!")
    except Exception as e:
        console.print(f"[red]✗[/red] Search Console authentication failed: {e}")
    
    # Test GA4 auth
    console.print("\n[bold]Testing GA4 authentication...[/bold]")
    try:
        from google.auth import get_ga4_credentials
        creds = get_ga4_credentials()
        console.print("[green]✓[/green] GA4 authentication successful!")
    except Exception as e:
        console.print(f"[red]✗[/red] GA4 authentication failed: {e}")
    
    # Test Google Ads auth if developer token is configured
    if config.GOOGLE_ADS_DEVELOPER_TOKEN:
        console.print("\n[bold]Testing Google Ads authentication...[/bold]")
        try:
            from google.ads_keywords import GoogleAdsKeywordClient
            ads_client = GoogleAdsKeywordClient()
            if ads_client.connect():
                console.print("[green]✓[/green] Google Ads authentication successful!")
            else:
                console.print("[yellow]⚠[/yellow] Google Ads connection failed - check developer token")
        except Exception as e:
            console.print(f"[red]✗[/red] Google Ads authentication failed: {e}")
    
    console.print("\n[green]Setup complete! You can now use --gsc, --ga4 flags, and keyword-metrics command.[/green]")


@app.command()
def clear_auth():
    """Clear stored Google authentication tokens."""
    console.print("[bold cyan]Clearing stored credentials...[/bold cyan]")
    clear_credentials()
    console.print("[green]✓[/green] Credentials cleared. Run 'auth-setup' to re-authenticate.")


@app.command()
def test_ai(
    provider: str = typer.Argument("gemini", help="AI provider to test (gemini or openai)")
):
    """Test AI provider connection."""
    console.print(f"[bold cyan]Testing {provider} connection...[/bold cyan]\n")
    
    try:
        recommender = KeywordRecommender(provider)
        if recommender.test_connection():
            console.print(f"[green]✓[/green] {provider} API connection successful!")
            
            # Test a simple recommendation
            console.print("\n[bold]Testing keyword recommendation...[/bold]")
            rec = recommender.get_recommendations(
                page_title="SEO Best Practices Guide",
                meta_description="Learn the best SEO practices for 2024",
                h1="Complete SEO Guide",
                content_snippet="This guide covers everything about SEO...",
                current_queries=[{"query": "seo guide", "clicks": 100, "impressions": 1000, "position": 8.5}]
            )
            
            if "error" not in rec:
                console.print(f"[green]✓[/green] Sample recommendation generated:")
                if rec.get("primary_keyword"):
                    console.print(f"  Primary: {rec['primary_keyword']}")
            else:
                console.print(f"[yellow]⚠[/yellow] Recommendation test returned: {rec['error']}")
        else:
            console.print(f"[red]✗[/red] {provider} API connection failed!")
            console.print(f"Check your {provider.upper()}_API_KEY in .env file")
            
    except Exception as e:
        console.print(f"[red]✗[/red] Error: {e}")


@app.command()
def keyword_metrics(
    keywords: list[str] = typer.Argument(..., help="Keywords to analyze (space-separated)"),
    provider: str = typer.Option("google-ads", "--provider", "-p", help="Data provider: google-ads, semrush, or both"),
    database: str = typer.Option("us", "--database", "-db", help="SEMrush database (us, uk, ca, au, etc.)"),
    location: str = typer.Option("2840", "--location", "-l", help="Google Ads location ID (2840=USA, 2826=UK, etc.)"),
    language: str = typer.Option("1000", "--language", "-lang", help="Language ID (1000=English)"),
    output: Optional[Path] = typer.Option(None, "--output", "-o", help="Output file for metrics"),
):
    """Get search volume and competition metrics for keywords via Google Ads or SEMrush API."""
    
    console.print(Panel.fit(
        Text("📊 Keyword Metrics", style="bold blue"),
        subtitle=f"Analyzing {len(keywords)} keywords via {provider}"
    ))
    
    all_metrics = []
    
    # Google Ads
    if provider in ["google-ads", "both"]:
        if not config.GOOGLE_ADS_DEVELOPER_TOKEN:
            console.print("[yellow]⚠[/yellow] GOOGLE_ADS_DEVELOPER_TOKEN not configured, skipping Google Ads")
        else:
            console.print("\n[bold cyan]Connecting to Google Ads API...[/bold cyan]")
            
            from google.ads_keywords import GoogleAdsKeywordClient, format_keyword_metrics_table
            
            ads_client = GoogleAdsKeywordClient()
            if ads_client.connect():
                console.print("[green]✓[/green] Google Ads connected")
                
                with Progress(
                    SpinnerColumn(),
                    TextColumn("[progress.description]{task.description}"),
                    console=console,
                ) as progress:
                    task = progress.add_task("Fetching Google Ads data...", total=None)
                    
                    metrics = ads_client.get_keyword_metrics(
                        keywords=list(keywords),
                        location_id=location,
                        language_id=language
                    )
                    progress.update(task, completed=True)
                
                if metrics:
                    console.print(f"[green]✓[/green] Fetched {len(metrics)} keywords from Google Ads")
                    for m in metrics:
                        m['source'] = 'google-ads'
                    all_metrics.extend(metrics)
                else:
                    console.print("[yellow]⚠[/yellow] No data from Google Ads")
            else:
                console.print("[red]✗[/red] Failed to connect to Google Ads API")
    
    # SEMrush
    if provider in ["semrush", "both"]:
        if not config.SEMRUSH_API_KEY:
            console.print("[yellow]⚠[/yellow] SEMRUSH_API_KEY not configured, skipping SEMrush")
        else:
            console.print("\n[bold cyan]Connecting to SEMrush API...[/bold cyan]")
            
            from google.semrush_keywords import SEMrushKeywordClient, format_semrush_metrics_table
            
            semrush_client = SEMrushKeywordClient()
            if semrush_client.test_connection():
                console.print("[green]✓[/green] SEMrush connected")
                
                with Progress(
                    SpinnerColumn(),
                    TextColumn("[progress.description]{task.description}"),
                    console=console,
                ) as progress:
                    task = progress.add_task("Fetching SEMrush data...", total=None)
                    
                    metrics = semrush_client.get_keyword_metrics(
                        keywords=list(keywords),
                        database=database
                    )
                    progress.update(task, completed=True)
                
                if metrics:
                    console.print(f"[green]✓[/green] Fetched {len(metrics)} keywords from SEMrush")
                    all_metrics.extend(metrics)
                else:
                    console.print("[yellow]⚠[/yellow] No data from SEMrush")
            else:
                console.print("[red]✗[/red] Failed to connect to SEMrush API")
    
    if not all_metrics:
        console.print("[red]No keyword metrics retrieved from any provider[/red]")
        raise typer.Exit(1)
    
    # Display results
    console.print("\n[bold green]Keyword Metrics Results:[/bold green]\n")
    
    # Group by source
    google_ads_metrics = [m for m in all_metrics if m.get('source') == 'google-ads']
    semrush_metrics = [m for m in all_metrics if m.get('source') == 'semrush']
    
    if google_ads_metrics:
        from google.ads_keywords import format_keyword_metrics_table
        console.print("**Google Ads Results:**")
        console.print(format_keyword_metrics_table(google_ads_metrics))
        console.print("")
    
    if semrush_metrics:
        from google.semrush_keywords import format_semrush_metrics_table
        console.print("**SEMrush Results:**")
        console.print(format_semrush_metrics_table(semrush_metrics))
        console.print("")
    
    # Save to file if requested
    if output:
        output_path = Path(output)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        report_lines = [
            "# Keyword Metrics Report",
            f"\nGenerated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"Provider: {provider}",
        ]
        
        if google_ads_metrics:
            from google.ads_keywords import format_keyword_metrics_table
            report_lines.extend(["\n## Google Ads Results\n", format_keyword_metrics_table(google_ads_metrics)])
        
        if semrush_metrics:
            from google.semrush_keywords import format_semrush_metrics_table
            report_lines.extend(["\n## SEMrush Results\n", format_semrush_metrics_table(semrush_metrics)])
        
        # Summary statistics
        report_lines.append("\n## Summary Statistics\n")
        if google_ads_metrics:
            total_searches = sum(m['avg_monthly_searches'] for m in google_ads_metrics)
            report_lines.append(f"**Google Ads:** {len(google_ads_metrics)} keywords, {total_searches:,} total monthly searches")
        
        if semrush_metrics:
            total_searches = sum(m['search_volume'] for m in semrush_metrics)
            report_lines.append(f"**SEMrush:** {len(semrush_metrics)} keywords, {total_searches:,} total monthly searches")
        
        output_path.write_text("\n".join(report_lines), encoding='utf-8')
        console.print(f"\n[green]✓[/green] Metrics saved to: {output_path}")
    
    # Summary
    console.print("\n" + "=" * 60)
    console.print(f"[bold]Summary:[/bold]")
    console.print(f"Total keywords analyzed: {len(all_metrics)}")
    if google_ads_metrics:
        total_searches = sum(m['avg_monthly_searches'] for m in google_ads_metrics)
        avg_competition = sum(m['competition_index'] for m in google_ads_metrics) / len(google_ads_metrics)
        console.print(f"Google Ads: {len(google_ads_metrics)} keywords, {total_searches:,} monthly searches, {avg_competition:.1f}/100 avg competition")
    if semrush_metrics:
        total_searches = sum(m['search_volume'] for m in semrush_metrics)
        avg_competition = sum(m['competition'] for m in semrush_metrics) / len(semrush_metrics) * 100
        console.print(f"SEMrush: {len(semrush_metrics)} keywords, {total_searches:,} monthly searches, {avg_competition:.1f}/100 avg competition")


if __name__ == "__main__":
    app()
