import sys
import os
from PySide6.QtCore import QThread, Qt, QPropertyAnimation, QEasingCurve, QTimer, QSize
from PySide6.QtWidgets import (
    QApplication, QWidget, QHBoxLayout, QVBoxLayout, QPushButton,
    QLabel, QFileDialog, QTextEdit, QMessageBox, QStackedWidget,
    QGraphicsOpacityEffect, QFrame, QScrollArea, QProgressBar
)
from PySide6.QtGui import QFont, QPalette, QColor, QIcon

from ui.pages.agency.page import Worker as WorkerAgency
from ui.pages.uyurk.page import Worker as WorkerUyruk


# -----------------------------------
# MODERN STYLED BUTTON
# -----------------------------------
class ModernButton(QPushButton):
    def __init__(self, text, primary=False, icon_text=None):
        super().__init__()
        self.primary = primary
        self.icon_text = icon_text

        if icon_text:
            self.setText(f"{icon_text}  {text}")
        else:
            self.setText(text)

        self.setMinimumHeight(48)
        self.setCursor(Qt.PointingHandCursor)
        self.update_style()

    def update_style(self):
        if self.primary:
            self.setStyleSheet("""
                QPushButton {
                    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                stop:0 #007AFF, stop:1 #0051D5);
                    color: white;
                    border: none;
                    border-radius: 12px;
                    font-size: 16px;
                    font-weight: 600;
                    padding: 14px 28px;
                }
                QPushButton:hover {
                    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                stop:0 #0051D5, stop:1 #004DB8);
                }
                QPushButton:pressed {
                    background: #004DB8;
                }
                QPushButton:disabled {
                    background: #C7C7CC;
                }
            """)
        else:
            self.setStyleSheet("""
                QPushButton {
                    background-color: white;
                    color: #000000;
                    border: 1.5px solid #E5E5EA;
                    border-radius: 12px;
                    font-size: 15px;
                    font-weight: 500;
                    padding: 12px 24px;
                }
                QPushButton:hover {
                    background-color: #F9F9F9;
                    border-color: #D1D1D6;
                }
                QPushButton:pressed {
                    background-color: #F2F2F7;
                }
            """)


# -----------------------------------
# SUCCESS BUTTON (Green Download)
# -----------------------------------
class SuccessButton(QPushButton):
    def __init__(self, text, icon_text="📥"):
        super().__init__(f"{icon_text}  {text}")
        self.setMinimumHeight(48)
        self.setCursor(Qt.PointingHandCursor)
        self.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                            stop:0 #34C759, stop:1 #28A745);
                color: white;
                border: none;
                border-radius: 12px;
                font-size: 16px;
                font-weight: 600;
                padding: 14px 28px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                            stop:0 #28A745, stop:1 #1E7E34);
            }
            QPushButton:pressed {
                background: #1E7E34;
            }
        """)


# -----------------------------------
# MODERN FILE SELECTOR CARD
# -----------------------------------
class FileCard(QWidget):
    def __init__(self, title, description, button_text="Dosya Seç", icon="📄"):
        super().__init__()
        self.file_path = None

        layout = QVBoxLayout()
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        # Header with icon and title
        header_layout = QHBoxLayout()
        icon_label = QLabel(icon)
        icon_label.setStyleSheet("font-size: 24px;")

        title_label = QLabel(title)
        title_label.setStyleSheet("""
            font-size: 17px;
            color: #000000;
            font-weight: 600;
        """)

        header_layout.addWidget(icon_label)
        header_layout.addWidget(title_label)
        header_layout.addStretch()

        # Description
        desc_label = QLabel(description)
        desc_label.setStyleSheet("""
            font-size: 13px;
            color: #8E8E93;
        """)
        desc_label.setWordWrap(True)

        # File path display with box
        path_container = QWidget()
        path_container.setStyleSheet("""
            QWidget {
                background-color: #F2F2F7;
                border-radius: 8px;
                padding: 12px;
            }
        """)
        path_layout = QVBoxLayout()
        path_layout.setContentsMargins(0, 0, 0, 0)

        self.path_label = QLabel("Henüz dosya seçilmedi")
        self.path_label.setStyleSheet("""
            font-size: 13px;
            color: #8E8E93;
            font-style: italic;
        """)
        self.path_label.setWordWrap(True)
        path_layout.addWidget(self.path_label)
        path_container.setLayout(path_layout)

        # Button
        self.btn = ModernButton(button_text, icon_text="📁")

        layout.addLayout(header_layout)
        layout.addWidget(desc_label)
        layout.addWidget(path_container)
        layout.addWidget(self.btn)

        self.setLayout(layout)
        self.setStyleSheet("""
            FileCard {
                background-color: white;
                border-radius: 16px;
                border: 1px solid #E5E5EA;
            }
        """)

    def set_file(self, path):
        self.file_path = path
        filename = path.split('/')[-1]
        self.path_label.setText(f"✓ {filename}")
        self.path_label.setStyleSheet("""
            font-size: 14px;
            color: #34C759;
            font-weight: 500;
            font-style: normal;
        """)


# -----------------------------------
# PROGRESS CARD
# -----------------------------------
class ProgressCard(QWidget):
    def __init__(self):
        super().__init__()

        layout = QVBoxLayout()
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        # Title
        title = QLabel("İşlem Durumu")
        title.setStyleSheet("""
            font-size: 17px;
            color: #000000;
            font-weight: 600;
        """)

        # Progress bar
        self.progress = QProgressBar()
        self.progress.setStyleSheet("""
            QProgressBar {
                border: none;
                border-radius: 8px;
                background-color: #E5E5EA;
                height: 8px;
                text-align: center;
            }
            QProgressBar::chunk {
                border-radius: 8px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                            stop:0 #007AFF, stop:1 #5AC8FA);
            }
        """)
        self.progress.setTextVisible(False)
        self.progress.setMaximum(0)  # Indeterminate
        self.progress.setMinimumHeight(8)
        self.progress.hide()

        # Status label
        self.status_label = QLabel("Hazır")
        self.status_label.setStyleSheet("""
            font-size: 14px;
            color: #8E8E93;
        """)

        layout.addWidget(title)
        layout.addWidget(self.progress)
        layout.addWidget(self.status_label)

        self.setLayout(layout)
        self.setStyleSheet("""
            ProgressCard {
                background-color: white;
                border-radius: 16px;
                border: 1px solid #E5E5EA;
            }
        """)

    def set_status(self, text, show_progress=False):
        self.status_label.setText(text)
        if show_progress:
            self.progress.show()
        else:
            self.progress.hide()


# -----------------------------------
# PAGE: Agency Report
# -----------------------------------
class AgencyPage(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.parent_window = parent
        self.input_file = None
        self.output_file = None
        self.output_path = "agency_rapor.xlsx"

        # Main scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea { border: none; background-color: transparent; }")

        content = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(50, 40, 50, 40)
        layout.setSpacing(24)

        # Header
        header_layout = QHBoxLayout()
        back_btn = QPushButton("← Geri")
        back_btn.clicked.connect(lambda: parent.show_page(0))
        back_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                color: #007AFF;
                font-size: 17px;
                text-align: left;
                padding: 8px 0px;
            }
            QPushButton:hover {
                color: #0051D5;
            }
        """)
        back_btn.setCursor(Qt.PointingHandCursor)
        header_layout.addWidget(back_btn)
        header_layout.addStretch()

        # Title section
        title = QLabel("📊 Agency Raporu")
        title.setStyleSheet("""
            font-size: 36px;
            font-weight: 700;
            color: #000000;
            margin-bottom: 8px;
        """)

        subtitle = QLabel("Agency verilerinizi analiz edin ve detaylı rapor oluşturun")
        subtitle.setStyleSheet("""
            font-size: 17px;
            color: #8E8E93;
            margin-bottom: 20px;
        """)

        # File cards
        self.input_card = FileCard(
            "Giriş Dosyası",
            "Agency verilerini içeren Excel dosyasını seçin",
            "Dosya Seç",
            "📊"
        )
        self.input_card.btn.clicked.connect(self.select_input)

        self.output_card = FileCard(
            "Çıktı Konumu",
            "Oluşturulan raporun kaydedileceği konumu belirleyin",
            "Konum Seç",
            "💾"
        )
        self.output_card.set_file(self.output_path)
        self.output_card.btn.clicked.connect(self.select_output)

        # Progress card
        self.progress_card = ProgressCard()

        # Log box
        log_label = QLabel("İşlem Günlüğü")
        log_label.setStyleSheet("""
            font-size: 17px;
            color: #000000;
            font-weight: 600;
            margin-top: 10px;
        """)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setStyleSheet("""
            QTextEdit {
                background-color: #1C1C1E;
                border: none;
                border-radius: 12px;
                padding: 16px;
                font-size: 13px;
                font-family: 'SF Mono', 'Consolas', 'Monaco', monospace;
                color: #00D4AA;
            }
        """)
        self.log_box.setMinimumHeight(180)

        # Action buttons
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(12)

        self.btn_start = ModernButton("Rapor Oluştur", primary=True, icon_text="▶️")
        self.btn_start.clicked.connect(self.start)

        self.btn_download = SuccessButton("Raporu İndir")
        self.btn_download.clicked.connect(self.download_report)
        self.btn_download.hide()

        buttons_layout.addWidget(self.btn_start)
        buttons_layout.addWidget(self.btn_download)

        # Add all to layout
        layout.addLayout(header_layout)
        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addWidget(self.input_card)
        layout.addWidget(self.output_card)
        layout.addWidget(self.progress_card)
        layout.addWidget(log_label)
        layout.addWidget(self.log_box)
        layout.addLayout(buttons_layout)
        layout.addStretch()

        content.setLayout(layout)
        scroll.setWidget(content)

        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.addWidget(scroll)
        self.setLayout(main_layout)

    def select_input(self):
        file, _ = QFileDialog.getOpenFileName(self, "Agency Dosyası Seç", "", "Excel Dosyaları (*.xlsx)")
        if file:
            self.input_file = file
            self.input_card.set_file(file)
            self.log_box.append(f"✓ Giriş dosyası seçildi: {file.split('/')[-1]}")

    def select_output(self):
        file, _ = QFileDialog.getSaveFileName(self, "Raporu Kaydet", "agency_rapor.xlsx", "Excel Dosyaları (*.xlsx)")
        if file:
            self.output_path = file
            self.output_card.set_file(file)
            self.log_box.append(f"✓ Çıktı konumu belirlendi: {file.split('/')[-1]}")

    def start(self):
        if not self.input_file:
            QMessageBox.warning(self, "Eksik Bilgi", "Lütfen giriş dosyasını seçiniz.")
            return

        self.log_box.clear()
        self.log_box.append("🚀 Agency raporu oluşturma işlemi başlatıldı...")
        self.btn_start.setEnabled(False)
        self.btn_download.hide()
        self.progress_card.set_status("İşlem devam ediyor...", True)

        self.thread = QThread()
        self.worker = WorkerAgency(self.input_file, None, self.output_path)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.on_finish)
        self.worker.error.connect(self.on_error)
        self.worker.log.connect(self.log_box.append)
        self.worker.finished.connect(self.thread.quit)
        self.worker.error.connect(self.thread.quit)
        self.thread.start()

    def on_finish(self, outfile):
        self.output_file = outfile
        self.log_box.append(f"\n✅ İşlem başarıyla tamamlandı!")
        self.log_box.append(f"📄 Rapor oluşturuldu: {outfile}")
        self.btn_start.setEnabled(True)
        self.btn_download.show()
        self.progress_card.set_status("İşlem başarıyla tamamlandı ✓", False)

        # Success message
        msg = QMessageBox(self)
        msg.setWindowTitle("Başarılı")
        msg.setText("Agency raporu başarıyla oluşturuldu!")
        msg.setIcon(QMessageBox.Information)
        msg.exec()

    def on_error(self, err):
        self.log_box.append(f"\n❌ Hata oluştu: {err}")
        self.btn_start.setEnabled(True)
        self.progress_card.set_status("Hata oluştu ✗", False)
        QMessageBox.critical(self, "Hata", f"İşlem sırasında hata oluştu:\n{err}")

    def download_report(self):
        if self.output_file and os.path.exists(self.output_file):
            os.startfile(self.output_file)  # Windows
            self.log_box.append(f"📥 Rapor açılıyor: {self.output_file}")


# -----------------------------------
# PAGE: Uyruk Report
# -----------------------------------
class UyrukPage(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.parent_window = parent
        self.input_file = None
        self.pairs_file = None
        self.output_file = None
        self.output_path = "uyruk_performans.xlsx"

        # Main scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea { border: none; background-color: transparent; }")

        content = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(50, 40, 50, 40)
        layout.setSpacing(24)

        # Header
        header_layout = QHBoxLayout()
        back_btn = QPushButton("← Geri")
        back_btn.clicked.connect(lambda: parent.show_page(0))
        back_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                color: #007AFF;
                font-size: 17px;
                text-align: left;
                padding: 8px 0px;
            }
            QPushButton:hover {
                color: #0051D5;
            }
        """)
        back_btn.setCursor(Qt.PointingHandCursor)
        header_layout.addWidget(back_btn)
        header_layout.addStretch()

        # Title section
        title = QLabel("🌍 Uyruk Performans Raporu")
        title.setStyleSheet("""
            font-size: 36px;
            font-weight: 700;
            color: #000000;
            margin-bottom: 8px;
        """)

        subtitle = QLabel("Bölge ve ülke bazlı performans analizinizi detaylı şekilde inceleyin")
        subtitle.setStyleSheet("""
            font-size: 17px;
            color: #8E8E93;
            margin-bottom: 20px;
        """)

        # File cards
        self.input_card = FileCard(
            "Uyruk Verileri",
            "Uyruk bilgilerini içeren Excel dosyasını seçin",
            "Dosya Seç",
            "🌍"
        )
        self.input_card.btn.clicked.connect(self.select_input)

        self.pairs_card = FileCard(
            "Bölge/Ülke Eşleştirme Tablosu",
            "Bölge ve ülke eşleştirmelerini içeren referans dosyasını seçin",
            "Dosya Seç",
            "🗺️"
        )
        self.pairs_card.btn.clicked.connect(self.select_pairs)

        self.output_card = FileCard(
            "Çıktı Konumu",
            "Oluşturulan raporun kaydedileceği konumu belirleyin",
            "Konum Seç",
            "💾"
        )
        self.output_card.set_file(self.output_path)
        self.output_card.btn.clicked.connect(self.select_output)

        # Progress card
        self.progress_card = ProgressCard()

        # Log box
        log_label = QLabel("İşlem Günlüğü")
        log_label.setStyleSheet("""
            font-size: 17px;
            color: #000000;
            font-weight: 600;
            margin-top: 10px;
        """)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setStyleSheet("""
            QTextEdit {
                background-color: #1C1C1E;
                border: none;
                border-radius: 12px;
                padding: 16px;
                font-size: 13px;
                font-family: 'SF Mono', 'Consolas', 'Monaco', monospace;
                color: #00D4AA;
            }
        """)
        self.log_box.setMinimumHeight(180)

        # Action buttons
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(12)

        self.btn_start = ModernButton("Rapor Oluştur", primary=True, icon_text="▶️")
        self.btn_start.clicked.connect(self.start)

        self.btn_download = SuccessButton("Raporu İndir")
        self.btn_download.clicked.connect(self.download_report)
        self.btn_download.hide()

        buttons_layout.addWidget(self.btn_start)
        buttons_layout.addWidget(self.btn_download)

        # Add all to layout
        layout.addLayout(header_layout)
        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addWidget(self.input_card)
        layout.addWidget(self.pairs_card)
        layout.addWidget(self.output_card)
        layout.addWidget(self.progress_card)
        layout.addWidget(log_label)
        layout.addWidget(self.log_box)
        layout.addLayout(buttons_layout)
        layout.addStretch()

        content.setLayout(layout)
        scroll.setWidget(content)

        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.addWidget(scroll)
        self.setLayout(main_layout)

    def select_input(self):
        file, _ = QFileDialog.getOpenFileName(self, "Uyruk Dosyası Seç", "", "Excel Dosyaları (*.xlsx)")
        if file:
            self.input_file = file
            self.input_card.set_file(file)
            self.log_box.append(f"✓ Uyruk dosyası seçildi: {file.split('/')[-1]}")

    def select_pairs(self):
        file, _ = QFileDialog.getOpenFileName(self, "Bölge/Ülke Tablosu Seç", "", "Excel Dosyaları (*.xlsx)")
        if file:
            self.pairs_file = file
            self.pairs_card.set_file(file)
            self.log_box.append(f"✓ Eşleştirme tablosu seçildi: {file.split('/')[-1]}")

    def select_output(self):
        file, _ = QFileDialog.getSaveFileName(self, "Raporu Kaydet", "uyruk_performans.xlsx",
                                              "Excel Dosyaları (*.xlsx)")
        if file:
            self.output_path = file
            self.output_card.set_file(file)
            self.log_box.append(f"✓ Çıktı konumu belirlendi: {file.split('/')[-1]}")

    def start(self):
        if not self.input_file or not self.pairs_file:
            QMessageBox.warning(self, "Eksik Bilgi", "Lütfen tüm gerekli dosyaları seçiniz.")
            return

        self.log_box.clear()
        self.log_box.append("🚀 Uyruk performans raporu oluşturma işlemi başlatıldı...")
        self.btn_start.setEnabled(False)
        self.btn_download.hide()
        self.progress_card.set_status("İşlem devam ediyor...", True)

        self.thread = QThread()
        self.worker = WorkerUyruk(self.input_file, self.pairs_file, self.output_path)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.on_finish)
        self.worker.error.connect(self.on_error)
        self.worker.log.connect(self.log_box.append)
        self.worker.finished.connect(self.thread.quit)
        self.worker.error.connect(self.thread.quit)
        self.thread.start()

    def on_finish(self, outfile):
        self.output_file = outfile
        self.log_box.append(f"\n✅ İşlem başarıyla tamamlandı!")
        self.log_box.append(f"📄 Rapor oluşturuldu: {outfile}")
        self.btn_start.setEnabled(True)
        self.btn_download.show()
        self.progress_card.set_status("İşlem başarıyla tamamlandı ✓", False)

        # Success message
        msg = QMessageBox(self)
        msg.setWindowTitle("Başarılı")
        msg.setText("Uyruk performans raporu başarıyla oluşturuldu!")
        msg.setIcon(QMessageBox.Information)
        msg.exec()

    def on_error(self, err):
        self.log_box.append(f"\n❌ Hata oluştu: {err}")
        self.btn_start.setEnabled(True)
        self.progress_card.set_status("Hata oluştu ✗", False)
        QMessageBox.critical(self, "Hata", f"İşlem sırasında hata oluştu:\n{err}")

    def download_report(self):
        if self.output_file and os.path.exists(self.output_file):
            os.startfile(self.output_file)  # Windows
            self.log_box.append(f"📥 Rapor açılıyor: {self.output_file}")


# -----------------------------------
# SELECTION PAGE
# -----------------------------------
class SelectionPage(QWidget):
    def __init__(self, parent):
        super().__init__()
        self.parent_window = parent

        layout = QVBoxLayout()
        layout.setContentsMargins(80, 60, 80, 60)
        layout.setSpacing(40)

        # Logo/Icon area
        logo_label = QLabel("📊")
        logo_label.setAlignment(Qt.AlignCenter)
        logo_label.setStyleSheet("font-size: 72px; margin-bottom: 10px;")

        # Title
        title = QLabel("Rapor Yönetim Sistemi")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("""
            font-size: 42px;
            font-weight: 700;
            color: #000000;
            letter-spacing: -1px;
        """)

        # Subtitle
        subtitle = QLabel("Verilerinizi analiz edin, detaylı raporlar oluşturun")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("""
            font-size: 18px;
            color: #8E8E93;
            margin-bottom: 30px;
        """)

        # Cards container
        cards_layout = QHBoxLayout()
        cards_layout.setSpacing(24)

        # Agency card
        agency_card = self.create_card(
            "📊",
            "Agency Raporu",
            "Agency verilerinizi analiz edin ve detaylı performans raporları oluşturun",
            ["• Otomatik rapor oluşturma", ],
            lambda: parent.show_page(1)
        )

        # Uyruk card
        uyruk_card = self.create_card(
            "🌍",
            "Uyruk Performans Raporu",
            "Bölge ve ülke bazlı performans metriklerini inceleyin ve karşılaştırın",
            ["• Karşılaştırmalı raporlama"],
            lambda: parent.show_page(2)
        )

        cards_layout.addWidget(agency_card)
        cards_layout.addWidget(uyruk_card)

        # Footer info
        footer = QLabel("Profesyonel veri analizi ve raporlama çözümü")
        footer.setAlignment(Qt.AlignCenter)
        footer.setStyleSheet("""
            font-size: 13px;
            color: #C7C7CC;
            margin-top: 20px;
        """)

        layout.addStretch()
        layout.addWidget(logo_label)
        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addLayout(cards_layout)
        layout.addWidget(footer)
        layout.addStretch()

        self.setLayout(layout)

    def create_card(self, icon, title, description, features, on_click):
        card = QWidget()
        card.setMinimumSize(380, 340)
        card.setCursor(Qt.PointingHandCursor)
        card.setStyleSheet("""
            QWidget {
                background-color: white;
                border-radius: 20px;
                border: 1.5px solid #E5E5EA;
            }
            QWidget:hover {
                background-color: #FAFAFA;
                border: 1.5px solid #007AFF;
            }
        """)

        layout = QVBoxLayout()
        layout.setContentsMargins(32, 32, 32, 32)
        layout.setSpacing(16)

        # Icon
        icon_label = QLabel(icon)
        icon_label.setAlignment(Qt.AlignLeft)
        icon_label.setStyleSheet("font-size: 56px;")

        # Title
        title_label = QLabel(title)
        title_label.setAlignment(Qt.AlignLeft)
        title_label.setStyleSheet("""
            font-size: 22px;
            font-weight: 700;
            color: #000000;
            margin-top: 8px;
        """)
        title_label.setWordWrap(True)

        # Description
        desc_label = QLabel(description)
        desc_label.setAlignment(Qt.AlignLeft)
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("""
            font-size: 14px;
            color: #8E8E93;
            line-height: 20px;
        """)

        # Features list
        features_widget = QWidget()
        features_layout = QVBoxLayout()
        features_layout.setContentsMargins(0, 12, 0, 0)
        features_layout.setSpacing(8)

        for feature in features:
            feature_label = QLabel(feature)
            feature_label.setStyleSheet("""
                font-size: 13px;
                color: #636366;
            """)
            features_layout.addWidget(feature_label)

        features_widget.setLayout(features_layout)

        # Arrow indicator
        arrow_label = QLabel("→")
        arrow_label.setAlignment(Qt.AlignRight)
        arrow_label.setStyleSheet("""
            font-size: 24px;
            color: #007AFF;
            font-weight: bold;
        """)

        layout.addWidget(icon_label)
        layout.addWidget(title_label)
        layout.addWidget(desc_label)
        layout.addWidget(features_widget)
        layout.addStretch()
        layout.addWidget(arrow_label)

        card.setLayout(layout)
        card.mousePressEvent = lambda e: on_click()

        return card


# -----------------------------------
# MAIN WINDOW
# -----------------------------------
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Rapor Yönetim Sistemi")
        self.setMinimumSize(1100, 750)

        # Set modern palette
        self.setStyleSheet("""
            QWidget {
                background-color: #F5F5F7;
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Helvetica Neue', sans-serif;
            }
            QMessageBox {
                background-color: white;
            }
            QScrollBar:vertical {
                border: none;
                background: #F5F5F7;
                width: 10px;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #C7C7CC;
                border-radius: 5px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background: #8E8E93;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
        """)

        # Stacked Widget with pages
        self.stack = QStackedWidget()
        self.selection_page = SelectionPage(self)
        self.agency_page = AgencyPage(self)
        self.uyruk_page = UyrukPage(self)

        self.stack.addWidget(self.selection_page)
        self.stack.addWidget(self.agency_page)
        self.stack.addWidget(self.uyruk_page)

        # Main layout
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.addWidget(self.stack)
        self.setLayout(main_layout)

        # Fade effect for transitions
        self.opacity_effect = QGraphicsOpacityEffect()
        self.stack.setGraphicsEffect(self.opacity_effect)

    def show_page(self, index):
        # Fade out animation
        self.fade_out = QPropertyAnimation(self.opacity_effect, b"opacity")
        self.fade_out.setDuration(200)
        self.fade_out.setStartValue(1.0)
        self.fade_out.setEndValue(0.0)
        self.fade_out.setEasingCurve(QEasingCurve.InOutCubic)
        self.fade_out.finished.connect(lambda: self._change_page(index))
        self.fade_out.start()

    def _change_page(self, index):
        self.stack.setCurrentIndex(index)

        # Fade in animation
        self.fade_in = QPropertyAnimation(self.opacity_effect, b"opacity")
        self.fade_in.setDuration(200)
        self.fade_in.setStartValue(0.0)
        self.fade_in.setEndValue(1.0)
        self.fade_in.setEasingCurve(QEasingCurve.InOutCubic)
        self.fade_in.start()


# -----------------------------------
# RUN APPLICATION
# -----------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Set application font
    font = QFont("Segoe UI", 13)
    app.setFont(font)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())
