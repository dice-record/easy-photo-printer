# -*- coding: utf-8 -*-

import sys
import os
import re
from datetime import datetime
from PIL import Image, ExifTags, ImageDraw, ImageFont, ImageOps
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QListWidget, QListWidgetItem, QLabel, QFileDialog,
    QComboBox, QSlider, QColorDialog, QLineEdit, QRadioButton,
    QGroupBox, QMessageBox, QStatusBar
)
from PyQt6.QtGui import QPixmap, QColor, QPainter, QImage, QPageSize
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize, QPoint, QRect, QPointF, QSizeF, QRectF
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog

# --- Helper functions: Date/time extraction logic ---
class DateTimeExtractor(QThread):
    """
    非同期で画像ファイルから日時情報を抽出するスレッド。
    結果をメインスレッドにシグナル経由で送信します。
    """
    # シグナル定義: (ファイルパス, 抽出された日時文字列)
    datetime_extracted = pyqtSignal(str, str)
    extraction_finished = pyqtSignal()
    status_message = pyqtSignal(str)
    
    def __init__(self, file_paths):
        super().__init__()
        self.file_paths = file_paths

    def run(self):
        for i, file_path in enumerate(self.file_paths):
            if self.isInterruptionRequested():
                return
            
            self.status_message.emit(f"画像 {i+1}/{len(self.file_paths)} の日時情報を取得中...")
            date_time_str = self._get_image_datetime(file_path)
            self.datetime_extracted.emit(file_path, date_time_str)
        self.extraction_finished.emit()

    def _get_image_datetime(self, file_path):
        """
        画像ファイルから日時情報を取得します。
        優先度: Exifデータ -> ファイル作成日時 -> ファイル名からの推測
        """
        try:
            # 1. Exifデータから作成日時を取得
            with Image.open(file_path) as img:
                exif_data = img._getexif()
                if exif_data:
                    for tag, value in exif_data.items():
                        tag_name = ExifTags.TAGS.get(tag, tag)
                        if tag_name in ['DateTimeOriginal', 'DateTimeDigitized']:
                            try:
                                return datetime.strptime(str(value), '%Y:%m:%d %H:%M:%S').strftime('%Y/%m/%d %H:%M')
                            except (ValueError, TypeError):
                                continue
        except (IOError, AttributeError, KeyError, ValueError, Exception) as e:
            print(f"[ERROR] Exifデータ読み込みエラー from {file_path}: {e}", file=sys.stderr)
            pass

        try:
            # 2. ファイル作成日時を取得
            timestamp = os.path.getctime(file_path)
            return datetime.fromtimestamp(timestamp).strftime('%Y/%m/%d %H:%M')
        except (IOError, ValueError, Exception) as e:
            print(f"[ERROR] ファイル作成日時取得エラー for {file_path}: {e}", file=sys.stderr)
            pass

        # 3. ファイル名から日時を推測
        filename = os.path.basename(file_path)
        patterns = [
            (r'(\d{8})_(\d{6})', '%Y%m%d_%H%M%S'),
            (r'IMG_(\d{8})_(\d{6})', '%Y%m%d_%H%M%S'),
            (r'DSC_(\d{14})', '%Y%m%d%H%M%S'),
            (r'(\d{8})', '%Y%m%d'),
            (r'IMG_(\d{8})', '%Y%m%d'),
        ]

        for pattern, fmt in patterns:
            match = re.search(pattern, filename)
            if match:
                try:
                    return datetime.strptime(match.group(0).replace('IMG_', '').replace('DSC_', ''), fmt).strftime('%Y/%m/%d %H:%M')
                except (ValueError, TypeError):
                    continue

        return "日時不明"

# --- Main Application Window ---
class PhotoTimestampPrinterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        try:
            self.setWindowTitle("フォトタイムスタンププリンター")
            self.setGeometry(100, 100, 1200, 800)
            self.setStyleSheet(self._get_stylesheet())

            self.selected_photos = {}
            self.datetime_extractor_thread = None
            self.current_preview_path = None
            self.text_color = QColor(Qt.GlobalColor.white)
            
            # UIウィジェットをインスタンス変数として保持
            self.photo_list_widget = QListWidget()
            self.path_label = QLabel("パス: (未選択)")
            self.date_format_combo = QComboBox()
            self.text_color_button = QPushButton("色を選択")
            self.text_color_preview = QLabel("■")
            self.text_size_slider = QSlider(Qt.Orientation.Horizontal)
            self.text_size_label = QLabel()
            self.bg_opacity_slider = QSlider(Qt.Orientation.Horizontal)
            self.bg_opacity_label = QLabel()
            self.color_radio = QRadioButton("カラー")
            self.mono_radio = QRadioButton("モノクロ")
            self.preview_area = QLabel("ここに印刷プレビューが表示されます")
            self.print_button = QPushButton("印刷実行")

            # ファイル選択ボタンもインスタンス変数として保持
            self.btn_select_files = QPushButton("写真ファイルを選択...")
            self.btn_select_folder = QPushButton("写真フォルダを選択...")
            
            self._init_ui()
            self._setup_connections()
        except Exception as e:
            QMessageBox.critical(self, "致命的なエラー", f"アプリケーションの初期化中にエラーが発生しました。\n詳細: {e}")
            sys.exit(1)

    def _get_stylesheet(self):
        """Tailwind CSS風のスタイルシートを定義します。"""
        return """
            QMainWindow {
                background-color: #f3f4f6;
                font-family: 'Inter', sans-serif;
            }
            QLabel {
                color: #374151;
            }
            QPushButton {
                background-color: #3b82f6;
                color: white;
                border-radius: 8px;
                padding: 12px 24px;
                font-weight: 600;
                border: none;
            }
            QPushButton:hover {
                background-color: #2563eb;
            }
            QPushButton#selectFolderButton {
                background-color: #4b5563;
            }
            QPushButton#selectFolderButton:hover {
                background-color: #374151;
            }
            QListWidget {
                background-color: #ffffff;
                border-radius: 8px;
                border: 1px solid #e5e7eb;
                padding: 10px;
            }
            QListWidgetItem {
                padding: 8px;
                margin-bottom: 5px;
                border-bottom: 1px solid #f3f4f6;
            }
            QListWidgetItem:selected {
                background-color: #bfdbfe;
            }
            QGroupBox {
                background-color: #eff6ff;
                border: 1px solid #bfdbfe;
                border-radius: 8px;
                margin-top: 1em;
                padding: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
                color: #1d4ed8;
                font-weight: bold;
            }
            QComboBox {
                padding: 6px;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                background-color: #ffffff;
                color: #000000;
            }
            QLineEdit, QSpinBox {
                padding: 6px;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                background-color: #ffffff;
            }
            QSlider::groove:horizontal {
                border: 1px solid #d1d5db;
                height: 8px;
                background: #e5e7eb;
                border-radius: 4px;
            }
            QSlider::handle:horizontal {
                background: #3b82f6;
                border: 1px solid #3b82f6;
                width: 18px;
                height: 18px;
                margin: -5px 0;
                border-radius: 9px;
            }
            #printButton {
                background-color: #22c55e;
                font-size: 20px;
                font-weight: bold;
                padding: 16px 32px;
            }
            #printButton:hover {
                background-color: #16a34a;
            }
            QStatusBar {
                background-color: #e5e7eb;
                color: #4b5563;
                padding: 8px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
            }
            #previewLabel {
                border: 2px dashed #d1d5db;
                border-radius: 8px;
                background-color: #f9fafb;
                min-height: 400px;
                color: #6b7280;
                font-style: italic;
                text-align: center;
            }
        """

    def _init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        left_pane = QVBoxLayout()
        left_pane.setContentsMargins(20, 20, 20, 20)
        left_pane.setSpacing(20)

        file_selection_group = QGroupBox("写真を選択")
        file_selection_layout = QVBoxLayout(file_selection_group)
        file_selection_group.setStyleSheet("QGroupBox::title { color: #1d4ed8; font-size: 18px; font-weight: bold; }")

        # インスタンス変数として保持したボタンを使用
        self.btn_select_files.setObjectName("selectFilesButton")
        self.btn_select_folder.setObjectName("selectFolderButton")
        
        file_selection_layout.addWidget(self.btn_select_files)
        file_selection_layout.addWidget(self.btn_select_folder)
        file_selection_layout.addWidget(self.path_label)
        left_pane.addWidget(file_selection_group)

        photo_list_group = QGroupBox("選択された写真")
        photo_list_layout = QVBoxLayout(photo_list_group)
        photo_list_group.setStyleSheet("QGroupBox::title { color: #1d4ed8; font-size: 18px; font-weight: bold; }")

        self.photo_list_widget.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        photo_list_layout.addWidget(self.photo_list_widget)
        left_pane.addWidget(photo_list_group)

        settings_group = QGroupBox("印刷設定")
        settings_layout = QVBoxLayout(settings_group)
        settings_group.setStyleSheet("QGroupBox::title { color: #1d4ed8; font-size: 18px; font-weight: bold; }")

        datetime_settings_group = QGroupBox("日時表示設定")
        datetime_settings_layout = QVBoxLayout(datetime_settings_group)
        settings_layout.addWidget(datetime_settings_group)

        form_layout = QHBoxLayout()
        form_layout.addWidget(QLabel("表示形式:"))
        self.date_format_combo.addItems(["YYYY/MM/DD HH:MM", "YY.MM.DD", "YYYY年MM月DD日", "カスタム..."])
        form_layout.addWidget(self.date_format_combo)
        datetime_settings_layout.addLayout(form_layout)

        form_layout = QHBoxLayout()
        form_layout.addWidget(QLabel("文字色:"))
        self.text_color_button.setObjectName("textColorButton")
        form_layout.addWidget(self.text_color_button)
        self.text_color_preview.setStyleSheet(f"color: {self.text_color.name()}; font-size: 20px;")
        form_layout.addWidget(self.text_color_preview)
        datetime_settings_layout.addLayout(form_layout)

        form_layout = QHBoxLayout()
        form_layout.addWidget(QLabel("文字サイズ:"))
        self.text_size_slider.setRange(8, 48)
        self.text_size_slider.setValue(16)
        self.text_size_label.setText(f"{self.text_size_slider.value()}px")
        form_layout.addWidget(self.text_size_slider)
        form_layout.addWidget(self.text_size_label)
        datetime_settings_layout.addLayout(form_layout)

        form_layout = QHBoxLayout()
        form_layout.addWidget(QLabel("背景透明度:"))
        self.bg_opacity_slider.setRange(0, 100)
        self.bg_opacity_slider.setValue(50)
        self.bg_opacity_label.setText(f"{self.bg_opacity_slider.value()}%")
        form_layout.addWidget(self.bg_opacity_slider)
        form_layout.addWidget(self.bg_opacity_label)
        datetime_settings_layout.addLayout(form_layout)
        
        layout_settings_group = QGroupBox("印刷レイアウト設定")
        layout_settings_layout = QVBoxLayout(layout_settings_group)
        settings_layout.addWidget(layout_settings_group)
        
        form_layout = QHBoxLayout()
        form_layout.addWidget(QLabel("カラー設定:"))
        self.color_radio.setChecked(True)
        form_layout.addWidget(self.color_radio)
        form_layout.addWidget(self.mono_radio)
        layout_settings_layout.addLayout(form_layout)

        left_pane.addWidget(settings_group)
        left_pane.addStretch(1)

        main_layout.addLayout(left_pane, 2)

        right_pane = QVBoxLayout()
        right_pane.setContentsMargins(20, 20, 20, 20)
        right_pane.setSpacing(20)

        preview_group = QGroupBox("印刷プレビュー")
        preview_layout = QVBoxLayout(preview_group)
        preview_group.setStyleSheet("QGroupBox::title { color: #1d4ed8; font-size: 18px; font-weight: bold; }")

        self.preview_area.setObjectName("previewLabel")
        self.preview_area.setAlignment(Qt.AlignmentFlag.AlignCenter)
        preview_layout.addWidget(self.preview_area)
        right_pane.addWidget(preview_group, 1)

        self.print_button.setObjectName("printButton")
        right_pane.addWidget(self.print_button)

        main_layout.addLayout(right_pane, 3)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("準備完了。")
    
    def _setup_connections(self):
        """シグナルとスロットを接続します。"""
        self.photo_list_widget.itemSelectionChanged.connect(self._update_preview)
        
        self.color_radio.toggled.connect(self._update_preview)
        self.mono_radio.toggled.connect(self._update_preview)
            
        self.text_color_button.clicked.connect(self._select_text_color)
        self.text_size_slider.valueChanged.connect(self._update_preview)
        self.bg_opacity_slider.valueChanged.connect(self._update_preview)
        self.date_format_combo.currentTextChanged.connect(self._update_preview)

        self.btn_select_files.clicked.connect(self._select_files)
        self.btn_select_folder.clicked.connect(self._select_folder)
        
        self.print_button.clicked.connect(self._print_action)
        
        self.text_size_slider.valueChanged.connect(lambda value: self.text_size_label.setText(f"{value}px"))
        self.bg_opacity_slider.valueChanged.connect(lambda value: self.bg_opacity_label.setText(f"{value}%"))

        QApplication.instance().aboutToQuit.connect(self._cleanup_threads)

    def _cleanup_threads(self):
        """アプリケーション終了時にスレッドをクリーンアップします。"""
        if self.datetime_extractor_thread and self.datetime_extractor_thread.isRunning():
            self.datetime_extractor_thread.requestInterruption()
            self.datetime_extractor_thread.wait()

    def _select_files(self):
        """画像ファイルを選択するためのファイルダイアログを開きます。"""
        file_dialog = QFileDialog()
        file_dialog.setNameFilter("画像ファイル (*.jpg *.jpeg *.png *.bmp *.gif *.tiff)")
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFiles)

        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            if selected_files:
                self.path_label.setText(f"パス: {', '.join(selected_files[:2])}{'...' if len(selected_files) > 2 else ''}")
                self._process_selected_files(selected_files)

    def _select_folder(self):
        """画像フォルダを選択するためのフォルダダイアログを開きます。"""
        folder_dialog = QFileDialog()
        folder_path = folder_dialog.getExistingDirectory(self, "画像フォルダを選択")

        if folder_path:
            self.path_label.setText(f"パス: {folder_path}")
            image_files = []
            try:
                for root, _, files in os.walk(folder_path):
                    for file in files:
                        if file.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff')):
                            image_files.append(os.path.join(root, file))
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"フォルダの読み込み中にエラーが発生しました。\n詳細: {e}")
                return

            if image_files:
                self._process_selected_files(image_files)
            else:
                QMessageBox.information(self, "情報", "選択されたフォルダに画像ファイルが見つかりませんでした。")
                self.photo_list_widget.clear()
                self.selected_photos.clear()
                self.status_bar.showMessage("画像ファイルが見つかりませんでした。")

    def _process_selected_files(self, file_paths):
        """選択されたファイルを処理し、日時情報を取得してリストに表示します。"""
        self.photo_list_widget.clear()
        self.selected_photos.clear()

        if self.datetime_extractor_thread and self.datetime_extractor_thread.isRunning():
            self.datetime_extractor_thread.requestInterruption()
            self.datetime_extractor_thread.wait()

        self.datetime_extractor_thread = DateTimeExtractor(file_paths)
        self.datetime_extractor_thread.datetime_extracted.connect(self._add_photo_to_list)
        self.datetime_extractor_thread.extraction_finished.connect(self._extraction_finished)
        self.datetime_extractor_thread.status_message.connect(self.status_bar.showMessage)
        self.datetime_extractor_thread.start()

    def _add_photo_to_list(self, file_path, date_time_str):
        """リストウィジェットに写真アイテムを追加します。"""
        item_text = f"{os.path.basename(file_path)} (日時: {date_time_str})"
        list_item = QListWidgetItem(item_text)
        list_item.setData(Qt.ItemDataRole.UserRole, file_path)
        self.photo_list_widget.addItem(list_item)
        self.selected_photos[file_path] = date_time_str
        if self.photo_list_widget.count() == 1:
            self.photo_list_widget.setCurrentItem(list_item)

    def _extraction_finished(self):
        """日時抽出スレッドの完了を処理します。"""
        self.status_bar.showMessage(f"{len(self.selected_photos)}枚の画像の日時情報取得が完了しました。")


    def _update_preview(self):
        """現在の設定で写真プレビューを更新します。"""
        selected_items = self.photo_list_widget.selectedItems()
        if not selected_items:
            self.preview_area.setText("ここに印刷プレビューが表示されます")
            self.preview_area.setPixmap(QPixmap())
            self.current_preview_path = None
            return

        first_item = selected_items[0]
        file_path = first_item.data(Qt.ItemDataRole.UserRole)
        
        try:
            pil_image = Image.open(file_path)
            
            # ここが修正箇所
            # モノクロ設定の場合、画像をグレースケールに変換し、さらにRGBに変換する
            if self.mono_radio.isChecked():
                pil_image = ImageOps.grayscale(pil_image)
                pil_image = pil_image.convert("RGB") # QImageで扱いやすいようRGBに変換
            else:
                # カラー設定の場合、RGBAに変換
                pil_image = pil_image.convert("RGBA")
            
            date_time_str = self.selected_photos.get(file_path, "日時不明")
            
            draw = ImageDraw.Draw(pil_image)
            try:
                font = ImageFont.truetype("arial.ttf", self.text_size_slider.value())
            except IOError:
                font = ImageFont.load_default()

            text_color_pil = (self.text_color.red(), self.text_color.green(), self.text_color.blue())
            draw.text((10, 10), date_time_str, font=font, fill=text_color_pil)
            
            # PILのモードに応じてQImageのフォーマットを適切に選択
            if pil_image.mode == "RGB":
                qimage = QImage(pil_image.tobytes(), pil_image.width, pil_image.height, pil_image.width * 3, QImage.Format.Format_RGB888)
            else:
                qimage = QImage(pil_image.tobytes(), pil_image.width, pil_image.height, pil_image.width * 4, QImage.Format.Format_RGBA8888)
            
            pixmap = QPixmap.fromImage(qimage)

            scaled_pixmap = pixmap.scaled(
                self.preview_area.size(),
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            )
            self.preview_area.setPixmap(scaled_pixmap)
            self.current_preview_path = file_path
        except Exception as e:
            self.preview_area.setText("プレビュー画像の読み込みに失敗しました。")
            self.preview_area.setPixmap(QPixmap())
            self.current_preview_path = None
            self.status_bar.showMessage(f"プレビュー画像の読み込みに失敗しました: {e}")
            return


    def _select_text_color(self):
        """文字色の選択ダイアログを開きます。"""
        color = QColorDialog.getColor(self.text_color, self, "文字色を選択")
        if color.isValid():
            self.text_color = color
            self.text_color_preview.setStyleSheet(f"color: {self.text_color.name()}; font-size: 20px;")
            self._update_preview()


    def _print_action(self):
        """印刷ボタンのアクションです。"""
        selected_items = self.photo_list_widget.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "警告", "印刷する写真が選択されていません。")
            return

        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        dialog = QPrintDialog(printer, self)
        # 印刷ダイアログでL判が選択されたり、デフォルト設定を使用するようにします
        # 単位をmmで設定
        printer.setPageSize(QPageSize(QSizeF(89, 127), QPageSize.Unit.Millimeter))
        printer.setResolution(300)

        print(f"[DEBUG] プリンター設定 - 解像度: {printer.resolution()}dpi")
        print(f"[DEBUG] プリンター設定 - ページサイズ (物理): {printer.pageRect(QPrinter.Unit.Millimeter).size().width()}mm x {printer.pageRect(QPrinter.Unit.Millimeter).size().height()}mm")
        print(f"[DEBUG] プリンター設定 - 印刷可能領域 (論理): {printer.pageRect(QPrinter.Unit.Point).size().width()}pt x {printer.pageRect(QPrinter.Unit.Point).size().height()}pt")

        if self.color_radio.isChecked():
            printer.setColorMode(QPrinter.ColorMode.Color)
            print("[DEBUG] カラーモード: カラー")
        else:
            printer.setColorMode(QPrinter.ColorMode.GrayScale)
            print("[DEBUG] カラーモード: グレースケール")
        
        if dialog.exec() == QPrintDialog.DialogCode.Accepted:
            self.status_bar.showMessage("印刷を開始します...")
            painter = QPainter()
            
            try:
                if not painter.begin(printer):
                    raise RuntimeError("プリンターの初期化に失敗しました。")
                print("[DEBUG] QPainter.begin() 成功")

                for i, item in enumerate(selected_items):
                    file_path = item.data(Qt.ItemDataRole.UserRole)
                    date_time_str = self.selected_photos.get(file_path, "日時不明")
                    print(f"[DEBUG] 印刷処理開始: {file_path}")

                    # PILを使って画像を処理し、タイムスタンプを描画
                    img = Image.open(file_path)
                    
                    # 印刷プレビューと一貫性を持たせるため、RGBに変換
                    if self.mono_radio.isChecked():
                        img = ImageOps.grayscale(img)
                        img = img.convert("RGB") # プレビューと同様にRGBに変換
                    else:
                        img = img.convert("RGBA")
                    
                    draw = ImageDraw.Draw(img)

                    try:
                        font = ImageFont.truetype("arial.ttf", self.text_size_slider.value())
                    except IOError:
                        font = ImageFont.load_default()

                    text_color_pil = (self.text_color.red(), self.text_color.green(), self.text_color.blue())
                    text_position = (10, 10)

                    draw.text(text_position, date_time_str, font=font, fill=text_color_pil)
                    print(f"[DEBUG] タイムスタンプ描画完了: {date_time_str}")

                    # 処理済みのPIL画像をQPixmapに変換
                    if img.mode == "RGB":
                        qimage = QImage(img.tobytes(), img.width, img.height, img.width * 3, QImage.Format.Format_RGB888)
                    else:
                        qimage = QImage(img.tobytes(), img.width, img.height, img.width * 4, QImage.Format.Format_RGBA8888)
                        
                    pixmap = QPixmap.fromImage(qimage)
                    print(f"[DEBUG] QPixmapへの変換完了. サイズ: {pixmap.size().width()}x{pixmap.size().height()}")

                    # 印刷可能領域を取得 (QRectF)
                    printable_rect_px = printer.pageRect(QPrinter.Unit.DevicePixel)
                    
                    print(f"[DEBUG] 印刷可能領域 (px): {printable_rect_px}")

                    # 描画対象のQRectFを計算
                    # プリンタの解像度に合わせて画像をスケーリングし、描画します。
                    # ここでfloatからintに変換します。
                    target_width_px = int(printable_rect_px.width())
                    target_height_px = int(printable_rect_px.height())

                    # QPainterは現在の解像度で描画するため、スケーリングしてプリンターに渡します
                    scaled_pixmap = pixmap.scaled(
                        target_width_px,
                        target_height_px,
                        Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation
                    )
                    
                    # 描画領域を計算
                    draw_width = scaled_pixmap.width()
                    draw_height = scaled_pixmap.height()
                    x = (target_width_px - draw_width) / 2
                    y = (target_height_px - draw_height) / 2

                    draw_rect = QRect(int(x), int(y), draw_width, draw_height)
                    print(f"[DEBUG] 描画領域を計算 (px): {draw_rect}")
                    
                    # QPainterの描画単位はデフォルトでポイントなので、
                    # 印刷可能領域に合わせてウィンドウ/ビューポートを設定
                    painter.setViewport(printable_rect_px.toRect()) # toRect()を追加
                    painter.setWindow(QRect(0, 0, target_width_px, target_height_px))
                    
                    print(f"[DEBUG] QPainter.drawPixmap() 実行")
                    painter.drawPixmap(draw_rect, scaled_pixmap)
                    
                    if i < len(selected_items) - 1:
                        print(f"[DEBUG] 次のページへ")
                        printer.newPage()

                painter.end()
                print("[DEBUG] QPainter.end() 成功")
                self.status_bar.showMessage(f"{len(selected_items)}枚の写真の印刷が完了しました。")

            except Exception as e:
                if painter.isActive():
                    painter.end()
                QMessageBox.critical(self, "エラー", f"印刷中にエラーが発生しました。\n詳細: {e}")
                self.status_bar.showMessage("印刷に失敗しました。")
        else:
            self.status_bar.showMessage("印刷はキャンセルされました。")


if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        window = PhotoTimestampPrinterApp()
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        QMessageBox.critical(None, "致命的なエラー", f"アプリケーションの起動中にエラーが発生しました。\n詳細: {e}")
        print(f"Error during application startup: {e}", file=sys.stderr)
        sys.exit(1)
