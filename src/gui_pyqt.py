#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
DOCX to HTML Converter - PyQt GUI
"""

import sys
import os
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QMessageBox,
    QListWidget, QDialog, QVBoxLayout, QPushButton, QLabel
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5 import uic

from converter import DocxConverter
from utils import find_docx_files, BatchResult


class ConversionWorker(QThread):
    """백그라운드 변환 작업 스레드"""

    file_started = pyqtSignal(str)  # filename
    file_done = pyqtSignal(str, bool, str)  # filename, success, error
    progress_updated = pyqtSignal(int, int)  # current, total
    finished_all = pyqtSignal(object)  # BatchResult
    log_message = pyqtSignal(str)

    def __init__(self, converter, files, output_path, options, is_batch=False, input_folder=None, output_folder=None):
        super().__init__()
        self.converter = converter
        self.files = files
        self.output_path = output_path
        self.options = options
        self.is_batch = is_batch
        self.input_folder = input_folder
        self.output_folder = output_folder
        self.cancel_requested = False

    def run(self):
        batch_result = BatchResult()

        if self.is_batch:
            # 배치 변환
            input_folder = Path(self.input_folder)
            output_folder = Path(self.output_folder)

            for i, input_file in enumerate(self.files):
                if self.cancel_requested:
                    self.log_message.emit("변환이 취소되었습니다.")
                    break

                rel_path = input_file.relative_to(input_folder)
                output_file = output_folder / rel_path.with_suffix('.html')

                self.file_started.emit(input_file.name)
                self.progress_updated.emit(i, len(self.files))

                result = self.converter.convert(str(input_file), str(output_file), self.options)
                batch_result.add(result)

                self.file_done.emit(input_file.name, result.success, result.error_message or '')

            self.progress_updated.emit(len(self.files), len(self.files))
        else:
            # 단일 파일 변환
            input_file = self.files[0]
            self.file_started.emit(Path(input_file).name)

            result = self.converter.convert(str(input_file), str(self.output_path), self.options)
            batch_result.add(result)

            self.file_done.emit(Path(input_file).name, result.success, result.error_message or '')

        self.finished_all.emit(batch_result)

    def cancel(self):
        self.cancel_requested = True


class FileListDialog(QDialog):
    """파일 목록 미리보기 대화상자"""

    def __init__(self, files, input_folder, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"파일 목록 ({len(files)}개)")
        self.setMinimumSize(500, 400)

        layout = QVBoxLayout(self)

        header = QLabel(f"총 {len(files)}개 파일")
        header.setStyleSheet("color: #6b7280; margin-bottom: 10px;")
        layout.addWidget(header)

        list_widget = QListWidget()
        list_widget.setStyleSheet("""
            QListWidget {
                background-color: white;
                border: 1px solid #e5e7eb;
                border-radius: 5px;
            }
        """)

        for f in files:
            rel_path = f.relative_to(input_folder) if f.is_relative_to(Path(input_folder)) else f
            list_widget.addItem(str(rel_path))

        layout.addWidget(list_widget)

        close_btn = QPushButton("닫기")
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn)


class ConverterApp(QMainWindow):
    """변환기 GUI 애플리케이션"""

    def __init__(self):
        super().__init__()

        # UI 로드
        ui_path = Path(__file__).parent / "ui" / "main_window.ui"
        uic.loadUi(ui_path, self)

        # 변환기 인스턴스
        self.converter = DocxConverter()

        # 상태 변수
        self.worker = None
        self.batch_result = None

        # 시그널 연결
        self._connect_signals()

        # 창 중앙 배치
        self._center_window()

    def _connect_signals(self):
        """시그널/슬롯 연결"""
        # 모드 전환
        self.singleModeRadio.toggled.connect(self._toggle_mode)

        # 파일/폴더 선택
        self.inputFileBrowseBtn.clicked.connect(self._browse_input_file)
        self.outputFileBrowseBtn.clicked.connect(self._browse_output_file)
        self.inputFolderBrowseBtn.clicked.connect(self._browse_input_folder)
        self.outputFolderBrowseBtn.clicked.connect(self._browse_output_folder)

        # 미리보기
        self.previewFilesBtn.clicked.connect(self._preview_files)

        # 변환 실행/취소
        self.convertBtn.clicked.connect(self._start_convert)
        self.cancelBtn.clicked.connect(self._cancel_convert)

        # CSV 내보내기
        self.exportCsvBtn.clicked.connect(self._export_csv)

    def _center_window(self):
        """창을 화면 중앙에 배치"""
        screen = QApplication.primaryScreen().geometry()
        x = (screen.width() - self.width()) // 2
        y = (screen.height() - self.height()) // 2
        self.move(x, y)

    def _toggle_mode(self, checked):
        """변환 모드 전환"""
        if checked:  # 단일 모드
            self.singleFrame.setVisible(True)
            self.batchFrame.setVisible(False)
        else:  # 배치 모드
            self.singleFrame.setVisible(False)
            self.batchFrame.setVisible(True)

    def _browse_input_file(self):
        """입력 파일 선택"""
        filepath, _ = QFileDialog.getOpenFileName(
            self, "Word 파일 선택", "",
            "Word 문서 (*.docx);;모든 파일 (*.*)"
        )
        if filepath:
            self.inputFileEdit.setText(filepath)
            output = Path(filepath).with_suffix('.html')
            self.outputFileEdit.setText(str(output))

    def _browse_output_file(self):
        """출력 위치 선택"""
        filepath, _ = QFileDialog.getSaveFileName(
            self, "저장 위치 선택", "",
            "HTML 파일 (*.html);;모든 파일 (*.*)"
        )
        if filepath:
            self.outputFileEdit.setText(filepath)

    def _browse_input_folder(self):
        """입력 폴더 선택"""
        folder = QFileDialog.getExistingDirectory(self, "입력 폴더 선택")
        if folder:
            self.inputFolderEdit.setText(folder)
            if not self.outputFolderEdit.text():
                self.outputFolderEdit.setText(folder)

    def _browse_output_folder(self):
        """출력 폴더 선택"""
        folder = QFileDialog.getExistingDirectory(self, "출력 폴더 선택")
        if folder:
            self.outputFolderEdit.setText(folder)

    def _preview_files(self):
        """파일 목록 미리보기"""
        input_folder = self.inputFolderEdit.text()
        if not input_folder:
            QMessageBox.warning(self, "경고", "입력 폴더를 선택하세요.")
            return

        files = find_docx_files(input_folder, self.includeSubfoldersCheck.isChecked())

        if not files:
            QMessageBox.information(self, "정보", "선택한 폴더에 .docx 파일이 없습니다.")
            return

        dialog = FileListDialog(files, input_folder, self)
        dialog.exec_()

    def _log(self, message):
        """로그 메시지 추가"""
        self.logText.append(message)

    def _clear_log(self):
        """로그 초기화"""
        self.logText.clear()

    def _start_convert(self):
        """변환 시작"""
        if self.singleModeRadio.isChecked():
            self._start_single_convert()
        else:
            self._start_batch_convert()

    def _start_single_convert(self):
        """단일 파일 변환"""
        input_file = self.inputFileEdit.text()
        output_file = self.outputFileEdit.text()

        if not input_file:
            QMessageBox.warning(self, "경고", "입력 파일을 선택하세요.")
            return

        if not output_file:
            QMessageBox.warning(self, "경고", "출력 위치를 선택하세요.")
            return

        self._prepare_conversion()
        self._log(f"변환 시작: {Path(input_file).name}")

        options = {
            'extract_images': self.extractImagesCheck.isChecked(),
            'remove_empty_paragraphs': self.removeEmptyCheck.isChecked()
        }

        self.worker = ConversionWorker(
            self.converter, [input_file], output_file, options
        )
        self._connect_worker_signals()
        self.worker.start()

    def _start_batch_convert(self):
        """배치 변환"""
        input_folder = self.inputFolderEdit.text()
        output_folder = self.outputFolderEdit.text()

        if not input_folder:
            QMessageBox.warning(self, "경고", "입력 폴더를 선택하세요.")
            return

        if not output_folder:
            QMessageBox.warning(self, "경고", "출력 폴더를 선택하세요.")
            return

        files = find_docx_files(input_folder, self.includeSubfoldersCheck.isChecked())

        if not files:
            QMessageBox.information(self, "정보", "변환할 .docx 파일이 없습니다.")
            return

        self._prepare_conversion()
        self._log(f"배치 변환 시작: {len(files)}개 파일")

        options = {
            'extract_images': self.extractImagesCheck.isChecked(),
            'remove_empty_paragraphs': self.removeEmptyCheck.isChecked()
        }

        self.worker = ConversionWorker(
            self.converter, files, None, options,
            is_batch=True, input_folder=input_folder, output_folder=output_folder
        )
        self._connect_worker_signals()
        self.worker.start()

    def _connect_worker_signals(self):
        """워커 스레드 시그널 연결"""
        self.worker.file_started.connect(self._on_file_started)
        self.worker.file_done.connect(self._on_file_done)
        self.worker.progress_updated.connect(self._on_progress_updated)
        self.worker.finished_all.connect(self._on_conversion_done)
        self.worker.log_message.connect(self._log)

    def _prepare_conversion(self):
        """변환 준비"""
        self._clear_log()
        self.convertBtn.setEnabled(False)
        self.cancelBtn.setEnabled(True)
        self.exportCsvBtn.setEnabled(False)
        self.currentProgress.setValue(0)
        self.totalProgress.setValue(0)

    def _on_file_started(self, filename):
        """파일 변환 시작"""
        self.currentFileLabel.setText(f"처리 중: {filename}")
        self.currentProgress.setMaximum(0)  # indeterminate mode

    def _on_file_done(self, filename, success, error):
        """파일 변환 완료"""
        self.currentProgress.setMaximum(100)
        status = "✓" if success else f"✗ {error}"
        self._log(f"  [{status}] {filename}")

    def _on_progress_updated(self, current, total):
        """진행률 업데이트"""
        self.totalProgressLabel.setText(f"전체: {current}/{total}")
        if total > 0:
            self.totalProgress.setValue(int(current / total * 100))

    def _on_conversion_done(self, batch_result):
        """변환 완료"""
        self.batch_result = batch_result
        self.currentProgress.setMaximum(100)
        self.currentProgress.setValue(100)
        self.currentFileLabel.setText("완료!")
        self.convertBtn.setEnabled(True)
        self.cancelBtn.setEnabled(False)
        self.exportCsvBtn.setEnabled(True)

        summary = batch_result.get_summary()

        self._log("─" * 45)
        self._log(f"  변환 완료!")
        self._log(f"  ✓ 성공: {summary['success']}개")
        if summary['failed'] > 0:
            self._log(f"  ✗ 실패: {summary['failed']}개")
        if summary['warnings'] > 0:
            self._log(f"  ⚠ 경고: {summary['warnings']}개")
        self._log("─" * 45)

        if summary['failed'] > 0:
            QMessageBox.warning(
                self, "완료",
                f"변환 완료\n\n✓ 성공: {summary['success']}개\n✗ 실패: {summary['failed']}개\n\n자세한 내용은 로그를 확인하세요."
            )
        else:
            QMessageBox.information(
                self, "완료",
                f"모든 파일 변환 완료!\n\n✓ 성공: {summary['success']}개"
            )

    def _cancel_convert(self):
        """변환 취소"""
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self._log("취소 요청됨...")

    def _export_csv(self):
        """결과 CSV 내보내기"""
        if not self.batch_result:
            return

        filepath, _ = QFileDialog.getSaveFileName(
            self, "CSV 저장", "",
            "CSV 파일 (*.csv)"
        )

        if filepath:
            try:
                self.batch_result.export_csv(filepath)
                QMessageBox.information(self, "완료", f"CSV 파일 저장 완료:\n{filepath}")
            except Exception as e:
                QMessageBox.critical(self, "오류", f"저장 실패: {e}")


def main():
    app = QApplication(sys.argv)

    # 애플리케이션 스타일 설정
    app.setStyle('Fusion')

    window = ConverterApp()
    window.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
