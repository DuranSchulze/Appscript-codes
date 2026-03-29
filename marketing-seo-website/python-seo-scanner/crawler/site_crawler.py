"""Site crawler for discovering and crawling website pages."""

import time
import xml.etree.ElementTree as ET
from urllib.parse import urljoin, urlparse
from typing import List, Set, Optional
import requests
from urllib.robotparser import RobotFileParser

from config import USER_AGENT, DEFAULT_MAX_PAGES, DEFAULT_MAX_DEPTH, DEFAULT_CRAWL_DELAY
from crawler.page_parser import PageParser, PageData


class SiteCrawler:
    def __init__(
        self,
        base_url: str,
        max_pages: int = DEFAULT_MAX_PAGES,
        max_depth: int = DEFAULT_MAX_DEPTH,
        delay: float = DEFAULT_CRAWL_DELAY
    ):
        self.base_url = base_url.rstrip('/')
        self.base_domain = urlparse(base_url).netloc
        self.max_pages = max_pages
        self.max_depth = max_depth
        self.delay = delay
        self.headers = {'User-Agent': USER_AGENT}
        
        self.parser = PageParser(base_url, self.headers)
        self.crawled_urls: Set[str] = set()
        self.pages: List[PageData] = []
        self.robots_parser: Optional[RobotFileParser] = None
        
        self._load_robots_txt()
    
    def _load_robots_txt(self):
        robots_url = f"{self.base_url}/robots.txt"
        try:
            self.robots_parser = RobotFileParser(robots_url)
            self.robots_parser.read()
        except Exception:
            self.robots_parser = None
    
    def can_fetch(self, url: str) -> bool:
        if not self.robots_parser:
            return True
        return self.robots_parser.can_fetch(USER_AGENT, url)
    
    def discover_from_sitemap(self) -> List[str]:
        sitemap_urls = [
            f"{self.base_url}/sitemap.xml",
            f"{self.base_url}/sitemap_index.xml",
            f"{self.base_url}/sitemap-index.xml",
        ]
        
        for sitemap_url in sitemap_urls:
            try:
                response = requests.get(sitemap_url, headers=self.headers, timeout=30)
                if response.status_code == 200:
                    return self._parse_sitemap(response.content)
            except requests.exceptions.RequestException:
                continue
        
        return []
    
    def _parse_sitemap(self, content: bytes) -> List[str]:
        urls = []
        try:
            root = ET.fromstring(content)
            ns = {'ns': 'http://www.sitemaps.org/schemas/sitemap/0.9'}
            
            for url_elem in root.findall('.//ns:url/ns:loc', ns):
                if url_elem.text:
                    urls.append(url_elem.text.strip())
            
            if not urls:
                for url_elem in root.findall('.//url/loc'):
                    if url_elem.text:
                        urls.append(url_elem.text.strip())
            
            for sitemap_elem in root.findall('.//ns:sitemap/ns:loc', ns):
                if sitemap_elem.text:
                    try:
                        response = requests.get(sitemap_elem.text, headers=self.headers, timeout=30)
                        if response.status_code == 200:
                            urls.extend(self._parse_sitemap(response.content))
                    except requests.exceptions.RequestException:
                        pass
        
        except ET.ParseError:
            pass
        
        return urls[:self.max_pages]
    
    def crawl(self) -> List[PageData]:
        sitemap_urls = self.discover_from_sitemap()
        
        if sitemap_urls:
            return self._crawl_urls(sitemap_urls)
        else:
            return self._crawl_recursive()
    
    def _crawl_urls(self, urls: List[str]) -> List[PageData]:
        for url in urls:
            if len(self.pages) >= self.max_pages:
                break
            
            if not self.can_fetch(url):
                continue
            
            page_data = self.parser.fetch_and_parse(url)
            self.pages.append(page_data)
            self.crawled_urls.add(url)
            
            time.sleep(self.delay)
        
        return self.pages
    
    def _crawl_recursive(self) -> List[PageData]:
        to_crawl = [(self.base_url, 0)]
        
        while to_crawl and len(self.pages) < self.max_pages:
            url, depth = to_crawl.pop(0)
            
            if url in self.crawled_urls:
                continue
            
            if not self.can_fetch(url):
                continue
            
            page_data = self.parser.fetch_and_parse(url)
            self.pages.append(page_data)
            self.crawled_urls.add(url)
            
            if depth < self.max_depth and page_data.status_code == 200:
                for link in page_data.internal_links:
                    if link not in self.crawled_urls:
                        to_crawl.append((link, depth + 1))
            
            time.sleep(self.delay)
        
        return self.pages
