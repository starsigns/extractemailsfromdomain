
import sys
import os
import asyncio
import aiohttp
import re
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog,
    QProgressBar, QMessageBox, QTableWidget, QTableWidgetItem, QHeaderView, QHBoxLayout
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer

EMAIL_REGEX = r'\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\b'

# Invalid email patterns to exclude
INVALID_PATTERNS = [
    r'.*\.(png|jpg|jpeg|gif|svg|webp|ico|css|js|pdf|doc|docx|zip|rar|exe)$',
    r'.*@(example\.com|test\.com|localhost|127\.0\.0\.1)',
    r'^[^@]*@[^@]*\.(png|jpg|jpeg|gif|svg|webp|ico|css|js|pdf)$',
    r'.*@.*\.(png|jpg|jpeg|gif|svg|webp|ico|css|js)$',
    r'.*@sentry.*',
    r'.*@.*sentry.*',
    r'.*@ovh\.net$',
    r'.*@.*\.ovhcloud\.com$',
    r'^abuse@.*',
    r'^u003e.*',
    r'.*@exemple\.',
    r'.*@email\.',
    r'.*@domain\.',
    r'.*@mail\.com$',
    r'.*@.*\.abc\.xyz$',
    r'.*@.*wixpress\.com$',
    r'.*@.*doctolib\.fr$',
    r'.*@.*doctolib\.com$',
    r'.*@.*\.mssante\.fr$',
    r'.*@.*\.apicrypt\.org$',
    r'.*exceptions\..*',
    r'.*\.ingest\.sentry\..*',
    r'^[a-f0-9]{32}@.*',
    r'^[a-f0-9]{40}@.*',
    r'.*@.*\.local\.fr$',
    r'.*@prestashop\.com$',
    r'.*@themeisle\.com$',
    r'.*@cal\.com$',
    r'.*@tally\.so$',
    r'.*@linkeo\.com$',
    r'.*@doe\.com$',
    r'.*@.*\.website\.com$',
    r'.*@yourwebsite\.com$',
    r'.*@.*gmail\.com$',  # Too many generic gmail addresses
]

class ExtractWorker(QThread):
    progress = pyqtSignal(int, int)
    result = pyqtSignal(list)
    status = pyqtSignal(str)
    auto_save = pyqtSignal()
    partial_results = pyqtSignal(list)

    def __init__(self, domains, auto_export_path=None):
        super().__init__()
        self.domains = domains
        self.results = []
        self.auto_export_path = auto_export_path
        self.processed_count = 0
        self.seen_emails = set()  # Track unique emails per domain

    def run(self):
        asyncio.run(self.async_run())

    async def async_run(self):
        connector = aiohttp.TCPConnector(limit=200, limit_per_host=20, ttl_dns_cache=300)
        timeout = aiohttp.ClientTimeout(total=3, connect=1)
        async with aiohttp.ClientSession(
            connector=connector, 
            timeout=timeout,
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        ) as session:
            # Process domains in larger batches for better performance
            batch_size = 100
            update_frequency = 2  # Update GUI every 2 batches to reduce hanging
            
            for i in range(0, len(self.domains), batch_size):
                batch = self.domains[i:i+batch_size]
                # Process batch domains concurrently
                tasks = [self.process_domain(session, domain) for domain in batch]
                await asyncio.gather(*tasks, return_exceptions=True)
                
                current_progress = min(i + batch_size, len(self.domains))
                
                # Only update GUI every few batches to prevent hanging
                if (i // batch_size) % update_frequency == 0 or current_progress >= len(self.domains):
                    self.progress.emit(current_progress, len(self.domains))
                    self.partial_results.emit(self.results.copy())
                
                # Auto-save every 500 processed domains (less frequent for speed)
                if self.auto_export_path and current_progress % 500 == 0:
                    self.auto_save.emit()
                    
            self.result.emit(self.results)
            self.status.emit('Extraction complete!')

    async def process_domain(self, session, domain):
        url = domain if domain.startswith('http') else f'http://{domain}'
        domain_emails = set()  # Track emails for this specific domain
        try:
            html = await self.fetch(session, url)
            if html:
                # Extract emails from home page
                home_emails = self.extract_emails(html)
                for email in home_emails:
                    email_domain_key = f"{email}|{domain}"
                    if email_domain_key not in self.seen_emails:
                        self.seen_emails.add(email_domain_key)
                        self.results.append({'domain': domain, 'email': email, 'source_url': url})
                        domain_emails.add(email)
                
                # Extract and process only priority links (faster filtering)
                links = self.extract_priority_links(html, url)
                limited_links = list(links)[:5]  # Reduced from 10 to 5 for speed
                
                if limited_links:
                    tasks = [self.fetch_and_extract_emails_with_source(session, link) for link in limited_links]
                    results = await asyncio.gather(*tasks, return_exceptions=True)
                    for result in results:
                        if isinstance(result, tuple) and len(result) == 2:
                            link_emails, source_url = result
                            for email in link_emails:
                                email_domain_key = f"{email}|{domain}"
                                if email_domain_key not in self.seen_emails:
                                    self.seen_emails.add(email_domain_key)
                                    self.results.append({'domain': domain, 'email': email, 'source_url': source_url})
                                    domain_emails.add(email)
        except Exception:
            pass

    async def fetch(self, session, url):
        try:
            async with session.get(url, timeout=3, allow_redirects=False) as resp:
                if resp.status == 200:
                    # Limit response size for speed (1MB max)
                    content = await resp.read()
                    if len(content) > 1024*1024:  # 1MB limit
                        content = content[:1024*1024]
                    return content.decode('utf-8', errors='ignore')
        except Exception:
            return None
        return None

    async def fetch_and_extract_emails_with_source(self, session, url):
        html = await self.fetch(session, url)
        if html:
            return self.extract_emails(html), url
        return set(), url

    def extract_emails(self, html):
        # Use faster regex with compiled pattern
        if not hasattr(self, '_email_pattern'):
            self._email_pattern = re.compile(EMAIL_REGEX, re.IGNORECASE)
        
        raw_emails = set(self._email_pattern.findall(html))
        valid_emails = set()
        
        for email in raw_emails:
            email = email.lower().strip()
            # Basic validation checks
            if self.is_valid_email(email):
                valid_emails.add(email)
                # Limit emails per page for speed
                if len(valid_emails) >= 20:
                    break
        
        return valid_emails

    def is_valid_email(self, email):
        # Check against invalid patterns
        for pattern in INVALID_PATTERNS:
            if re.match(pattern, email, re.IGNORECASE):
                return False
        
        # Additional validation checks
        if len(email) < 5 or len(email) > 254:
            return False
        
        # Check for common invalid patterns
        if email.count('@') != 1:
            return False
        
        local, domain = email.split('@')
        
        # Local part validation
        if len(local) == 0 or len(local) > 64:
            return False
        
        # Domain validation
        if len(domain) == 0 or len(domain) > 253:
            return False
        
        # Check if domain has at least one dot
        if '.' not in domain:
            return False
        
        # Check for common spam/invalid domains and patterns
        spam_domains = [
            'noreply', 'no-reply', 'donotreply', 'example.com', 'test.com',
            'sentry', 'ovh.net', 'ovhcloud.com', 'wixpress.com', 'doctolib',
            'mssante.fr', 'apicrypt.org', 'prestashop.com', 'themeisle.com',
            'linkeo.com', 'tally.so', 'cal.com'
        ]
        if any(spam in domain.lower() for spam in spam_domains):
            return False
        
        # Check for hash-like patterns (32+ hex characters)
        if re.match(r'^[a-f0-9]{32,}@', email):
            return False
        
        # Check for obvious test emails
        test_patterns = [
            'exemple', 'example', 'test@', 'demo@', 'sample@',
            'jean.dupont', 'marie.durand', 'martin.durand', 'john@doe',
            'nom@', 'votre@', 'your@', 'email@domain', 'user@mail'
        ]
        if any(test in email.lower() for test in test_patterns):
            return False
        
        # Check for u003e prefix (encoded characters)
        if email.startswith('u003e'):
            return False
        
        return True

    def extract_priority_links(self, html, base_url):
        """Extract only high-priority links likely to contain emails"""
        soup = BeautifulSoup(html, 'html.parser')
        links = set()
        
        # Priority keywords for email-likely pages
        priority_keywords = [
            'contact', 'about', 'team', 'staff', 'management', 'direction',
            'legal', 'mentions', 'imprint', 'impressum', 'support', 'help'
        ]
        
        for a in soup.find_all('a', href=True):
            href = a['href'].lower()
            link_text = a.get_text().lower()
            
            # Check if href or link text contains priority keywords
            if any(keyword in href or keyword in link_text for keyword in priority_keywords):
                if href.startswith('http'):
                    links.add(a['href'])
                elif href.startswith('/'):
                    links.add(self.join_url(base_url, a['href']))
                    
        return links

    def extract_links(self, html, base_url):
        soup = BeautifulSoup(html, 'html.parser')
        links = set()
        for a in soup.find_all('a', href=True):
            href = a['href']
            if href.startswith('http'):
                links.add(href)
            elif href.startswith('/'):
                links.add(self.join_url(base_url, href))
        return links

    def join_url(self, base, path):
        if base.endswith('/'):
            base = base[:-1]
        if path.startswith('/'):
            path = path[1:]
        return f'{base}/{path}'

class EmailExtractorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Email Extractor from Domains')
        self.resize(700, 500)
        self.domains = []
        self.results = []
        self.worker = None
        self.pending_results = []
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.process_pending_updates)
        self.update_timer.setSingleShot(True)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        self.file_label = QLabel('No file selected')
        layout.addWidget(self.file_label)

        btn_layout = QHBoxLayout()
        self.browse_btn = QPushButton('Browse File')
        self.browse_btn.clicked.connect(self.browse_file)
        btn_layout.addWidget(self.browse_btn)

        self.start_btn = QPushButton('Start Extraction')
        self.start_btn.setEnabled(False)
        self.start_btn.clicked.connect(self.start_extraction)
        btn_layout.addWidget(self.start_btn)
        layout.addLayout(btn_layout)

        self.progress = QProgressBar()
        layout.addWidget(self.progress)

        self.status = QLabel('')
        layout.addWidget(self.status)

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(['Domain', 'Email', 'Source URL'])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.table)

        self.export_btn = QPushButton('Export to Excel')
        self.export_btn.setEnabled(False)
        self.export_btn.clicked.connect(self.export_results)
        layout.addWidget(self.export_btn)

        self.setLayout(layout)

    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, 'Open file', '', 'Text Files (*.txt)')
        if file_path:
            self.file_label.setText(os.path.basename(file_path))
            with open(file_path, 'r', encoding='utf-8') as f:
                self.domains = [line.strip() for line in f if line.strip()]
            self.start_btn.setEnabled(True)
            self.status.setText(f'{len(self.domains)} domains loaded.')

    def start_extraction(self):
        self.start_btn.setEnabled(False)
        self.export_btn.setEnabled(False)
        self.progress.setValue(0)
        self.progress.setMaximum(len(self.domains))
        self.status.setText('Starting extraction...')
        self.table.setRowCount(0)
        self.results = []
        
        # Set up auto-export path
        auto_export_path = os.path.join(os.getcwd(), f'emails_auto_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
        
        self.worker = ExtractWorker(self.domains, auto_export_path)
        self.worker.progress.connect(self.update_progress)
        self.worker.result.connect(self.show_results)
        self.worker.status.connect(self.update_status)
        self.worker.partial_results.connect(self.update_table_live)
        self.worker.auto_save.connect(lambda: self.auto_save_results(auto_export_path))
        self.worker.start()

    def update_progress(self, value, maximum):
        self.progress.setValue(value)
        self.progress.setMaximum(maximum)
        self.status.setText(f'Processed {value}/{maximum} domains | Found {len(self.results)} emails')

    def show_results(self, results):
        self.results = results
        # Force final update without timer
        self.pending_results = results
        self.process_pending_updates()
        self.export_btn.setEnabled(True)

    def update_status(self, text):
        self.status.setText(text)
        self.start_btn.setEnabled(True)

    def update_table_live(self, results):
        """Update the table with live results as they come in"""
        self.pending_results = results
        # Use timer to batch updates and prevent GUI hanging
        if not self.update_timer.isActive():
            self.update_timer.start(500)  # Update every 500ms at most

    def process_pending_updates(self):
        """Process pending table updates in batches"""
        if not self.pending_results:
            return
            
        results = self.pending_results
        self.results = results
        
        # Limit table size to prevent performance issues
        max_display_rows = 1000
        display_results = results[-max_display_rows:] if len(results) > max_display_rows else results
        
        self.table.setRowCount(len(display_results))
        
        # Update in chunks to prevent hanging
        chunk_size = 50
        for chunk_start in range(0, len(display_results), chunk_size):
            chunk_end = min(chunk_start + chunk_size, len(display_results))
            for i in range(chunk_start, chunk_end):
                if i < len(display_results):
                    row = display_results[i]
                    self.table.setItem(i, 0, QTableWidgetItem(row['domain']))
                    self.table.setItem(i, 1, QTableWidgetItem(row['email']))
                    self.table.setItem(i, 2, QTableWidgetItem(row['source_url']))
            
            # Process events between chunks to keep GUI responsive
            QApplication.processEvents()
        
        # Scroll to bottom to show latest results
        self.table.scrollToBottom()

    def auto_save_results(self, file_path):
        if self.results:
            self.save_to_excel(file_path, self.results)
            self.status.setText(f'Auto-saved {len(self.results)} results to {os.path.basename(file_path)} | Processed {self.progress.value()}/{self.progress.maximum()} domains')

    def export_results(self):
        file_path, _ = QFileDialog.getSaveFileName(self, 'Save File', '', 'Excel Files (*.xlsx)')
        if file_path:
            self.save_to_excel(file_path, self.results)
            QMessageBox.information(self, 'Export', f'Results exported to {file_path}')

    def save_to_excel(self, file_path, results):
        wb = Workbook()
        
        # Create Results sheet
        ws_results = wb.active
        ws_results.title = "Email Results"
        
        # Headers
        headers = ['Domain', 'Email', 'Source URL']
        for col, header in enumerate(headers, 1):
            cell = ws_results.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
        
        # Data
        for row, result in enumerate(results, 2):
            ws_results.cell(row=row, column=1, value=result['domain'])
            ws_results.cell(row=row, column=2, value=result['email'])
            ws_results.cell(row=row, column=3, value=result['source_url'])
        
        # Auto-adjust column widths
        for column in ws_results.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws_results.column_dimensions[column_letter].width = adjusted_width
        
        # Create Summary sheet
        ws_summary = wb.create_sheet("Summary")
        
        # Summary data
        total_emails = len(results)
        unique_domains = len(set(result['domain'] for result in results))
        unique_emails = len(set(result['email'] for result in results))
        duplicate_count = total_emails - unique_emails
        
        # Domain statistics
        domain_counts = {}
        for result in results:
            domain = result['domain']
            domain_counts[domain] = domain_counts.get(domain, 0) + 1
        
        # Summary headers and data
        summary_data = [
            ['Summary Statistics', ''],
            ['Total Email Entries', total_emails],
            ['Unique Emails', unique_emails],
            ['Duplicate Entries Prevented', duplicate_count],
            ['Domains Processed', unique_domains],
            ['Processing Date', datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
            ['', ''],
            ['Top Domains by Email Count', ''],
        ]
        
        # Add top 10 domains
        top_domains = sorted(domain_counts.items(), key=lambda x: x[1], reverse=True)[:10]
        for domain, count in top_domains:
            summary_data.append([domain, count])
        
        # Write summary data
        for row, (label, value) in enumerate(summary_data, 1):
            ws_summary.cell(row=row, column=1, value=label)
            ws_summary.cell(row=row, column=2, value=value)
            if row == 1 or row == 7:  # Headers
                ws_summary.cell(row=row, column=1).font = Font(bold=True)
        
        # Auto-adjust summary column widths
        for column in ws_summary.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = max_length + 2
            ws_summary.column_dimensions[column_letter].width = adjusted_width
        
        wb.save(file_path)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = EmailExtractorApp()
    window.show()
    sys.exit(app.exec_())
