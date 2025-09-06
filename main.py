
import sys
import os
import asyncio
import aiohttp
import re
import csv
from bs4 import BeautifulSoup
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog,
    QProgressBar, QMessageBox, QTableWidget, QTableWidgetItem, QHeaderView, QHBoxLayout
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

EMAIL_REGEX = r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+'

class ExtractWorker(QThread):
    progress = pyqtSignal(int, int)
    result = pyqtSignal(list)
    status = pyqtSignal(str)

    def __init__(self, domains):
        super().__init__()
        self.domains = domains
        self.results = []

    def run(self):
        asyncio.run(self.async_run())

    async def async_run(self):
        connector = aiohttp.TCPConnector(limit=100, limit_per_host=10)
        timeout = aiohttp.ClientTimeout(total=5)
        async with aiohttp.ClientSession(connector=connector, timeout=timeout) as session:
            # Process domains in batches of 50 for better performance
            batch_size = 50
            for i in range(0, len(self.domains), batch_size):
                batch = self.domains[i:i+batch_size]
                tasks = [self.process_domain(session, domain) for domain in batch]
                await asyncio.gather(*tasks, return_exceptions=True)
                self.progress.emit(min(i + batch_size, len(self.domains)), len(self.domains))
            self.result.emit(self.results)
            self.status.emit('Extraction complete!')

    async def process_domain(self, session, domain):
        url = domain if domain.startswith('http') else f'http://{domain}'
        emails = set()
        try:
            html = await self.fetch(session, url)
            if html:
                emails.update(self.extract_emails(html))
                links = self.extract_links(html, url)
                # Limit to 10 links per domain for performance
                limited_links = list(links)[:10]
                tasks = [self.fetch_and_extract_emails(session, link) for link in limited_links]
                results = await asyncio.gather(*tasks, return_exceptions=True)
                for result in results:
                    if isinstance(result, set):
                        emails.update(result)
        except Exception:
            pass
        if emails:
            for email in emails:
                self.results.append({'domain': domain, 'email': email})

    async def fetch(self, session, url):
        try:
            async with session.get(url, timeout=5) as resp:
                if resp.status == 200:
                    return await resp.text()
        except Exception:
            return None
        return None

    async def fetch_and_extract_emails(self, session, url):
        html = await self.fetch(session, url)
        if html:
            return self.extract_emails(html)
        return set()

    def extract_emails(self, html):
        return set(re.findall(EMAIL_REGEX, html))

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

        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(['Domain', 'Email'])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.table)

        self.export_btn = QPushButton('Export Results')
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
        self.worker = ExtractWorker(self.domains)
        self.worker.progress.connect(self.update_progress)
        self.worker.result.connect(self.show_results)
        self.worker.status.connect(self.update_status)
        self.worker.start()

    def update_progress(self, value, maximum):
        self.progress.setValue(value)
        self.progress.setMaximum(maximum)
        self.status.setText(f'Processed {value}/{maximum} domains')

    def show_results(self, results):
        self.results = results
        self.table.setRowCount(len(results))
        for i, row in enumerate(results):
            self.table.setItem(i, 0, QTableWidgetItem(row['domain']))
            self.table.setItem(i, 1, QTableWidgetItem(row['email']))
        self.export_btn.setEnabled(True)

    def update_status(self, text):
        self.status.setText(text)
        self.start_btn.setEnabled(True)

    def export_results(self):
        file_path, _ = QFileDialog.getSaveFileName(self, 'Save File', '', 'CSV Files (*.csv)')
        if file_path:
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=['domain', 'email'])
                writer.writeheader()
                for row in self.results:
                    writer.writerow(row)
            QMessageBox.information(self, 'Export', f'Results exported to {file_path}')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = EmailExtractorApp()
    window.show()
    sys.exit(app.exec_())
