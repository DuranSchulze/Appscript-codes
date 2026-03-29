"""HTML page parser for extracting SEO elements."""

from dataclasses import dataclass
from typing import List, Optional
from bs4 import BeautifulSoup
import requests
from urllib.parse import urljoin, urlparse


@dataclass
class PageData:
    url: str
    status_code: int
    title: str = ""
    meta_description: str = ""
    canonical_url: str = ""
    robots_meta: str = ""
    h1: str = ""
    h2_list: List[str] = None
    h3_list: List[str] = None
    word_count: int = 0
    internal_links: List[str] = None
    images_without_alt: int = 0
    has_schema_markup: bool = False
    content_snippet: str = ""
    error: Optional[str] = None
    
    def __post_init__(self):
        if self.h2_list is None:
            self.h2_list = []
        if self.h3_list is None:
            self.h3_list = []
        if self.internal_links is None:
            self.internal_links = []


class PageParser:
    def __init__(self, base_url: str, headers: dict):
        self.base_url = base_url
        self.base_domain = urlparse(base_url).netloc
        self.headers = headers
    
    def fetch_and_parse(self, url: str) -> PageData:
        try:
            response = requests.get(
                url, 
                headers=self.headers, 
                timeout=30,
                allow_redirects=True
            )
            response.raise_for_status()
            
            soup = BeautifulSoup(response.content, 'lxml')
            
            return self._extract_data(url, response.status_code, soup)
            
        except requests.exceptions.RequestException as e:
            return PageData(
                url=url,
                status_code=getattr(e.response, 'status_code', 0),
                error=str(e)
            )
    
    def _extract_data(self, url: str, status_code: int, soup: BeautifulSoup) -> PageData:
        title = self._get_title(soup)
        meta_description = self._get_meta_description(soup)
        canonical = self._get_canonical(soup)
        robots = self._get_robots_meta(soup)
        h1 = self._get_h1(soup)
        h2_list = self._get_headings(soup, 'h2')
        h3_list = self._get_headings(soup, 'h3')
        word_count = self._get_word_count(soup)
        internal_links = self._get_internal_links(soup, url)
        images_without_alt = self._count_images_without_alt(soup)
        has_schema = self._has_schema_markup(soup)
        content_snippet = self._get_content_snippet(soup)
        
        return PageData(
            url=url,
            status_code=status_code,
            title=title,
            meta_description=meta_description,
            canonical_url=canonical,
            robots_meta=robots,
            h1=h1,
            h2_list=h2_list,
            h3_list=h3_list,
            word_count=word_count,
            internal_links=internal_links,
            images_without_alt=images_without_alt,
            has_schema_markup=has_schema,
            content_snippet=content_snippet
        )
    
    def _get_title(self, soup: BeautifulSoup) -> str:
        title_tag = soup.find('title')
        return title_tag.get_text(strip=True) if title_tag else ""
    
    def _get_meta_description(self, soup: BeautifulSoup) -> str:
        meta = soup.find('meta', attrs={'name': 'description'})
        if meta:
            return meta.get('content', '')
        meta = soup.find('meta', attrs={'property': 'og:description'})
        return meta.get('content', '') if meta else ""
    
    def _get_canonical(self, soup: BeautifulSoup) -> str:
        canonical = soup.find('link', attrs={'rel': 'canonical'})
        return canonical.get('href', '') if canonical else ""
    
    def _get_robots_meta(self, soup: BeautifulSoup) -> str:
        robots = soup.find('meta', attrs={'name': 'robots'})
        return robots.get('content', '') if robots else ""
    
    def _get_h1(self, soup: BeautifulSoup) -> str:
        h1 = soup.find('h1')
        return h1.get_text(strip=True) if h1 else ""
    
    def _get_headings(self, soup: BeautifulSoup, tag: str) -> List[str]:
        return [h.get_text(strip=True) for h in soup.find_all(tag, limit=10)]
    
    def _get_word_count(self, soup: BeautifulSoup) -> int:
        text = soup.get_text(separator=' ', strip=True)
        words = text.split()
        return len(words)
    
    def _get_internal_links(self, soup: BeautifulSoup, current_url: str) -> List[str]:
        links = []
        for a_tag in soup.find_all('a', href=True):
            href = a_tag['href']
            full_url = urljoin(current_url, href)
            parsed = urlparse(full_url)
            
            if parsed.netloc == self.base_domain and full_url not in links:
                links.append(full_url)
        
        return links[:50]
    
    def _count_images_without_alt(self, soup: BeautifulSoup) -> int:
        images = soup.find_all('img')
        return sum(1 for img in images if not img.get('alt'))
    
    def _has_schema_markup(self, soup: BeautifulSoup) -> bool:
        script_tags = soup.find_all('script', type='application/ld+json')
        return len(script_tags) > 0
    
    def _get_content_snippet(self, soup: BeautifulSoup) -> str:
        for tag in ['article', 'main', '[role="main"]']:
            content = soup.find(tag)
            if content:
                text = content.get_text(separator=' ', strip=True)
                return text[:500]
        
        text = soup.get_text(separator=' ', strip=True)
        return text[:500]
