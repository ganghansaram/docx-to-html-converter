#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
DOCX to HTML Converter - GUI
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading
import queue

from converter import DocxConverter
from utils import find_docx_files, BatchResult


class ConverterApp:
    """변환기 GUI 애플리케이션"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("DOCX to HTML Converter")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        self.root.minsize(600, 500)

        # 변환기 인스턴스
        self.converter = DocxConverter()

        # 변수 초기화
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.extract_images = tk.BooleanVar(value=True)
        self.remove_empty = tk.BooleanVar(value=True)

        # 배치 모드 변수
        self.batch_mode = tk.BooleanVar(value=False)
        self.include_subfolders = tk.BooleanVar(value=True)
        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()

        # 진행 상태
        self.is_converting = False
        self.cancel_requested = False
        self.progress_queue = queue.Queue()

        # 배치 결과
        self.batch_result = None

        self._create_widgets()
        self._setup_progress_checker()

    def _create_widgets(self):
        """GUI 위젯 생성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 제목
        title_label = ttk.Label(
            main_frame,
            text="Word -> HTML 변환기",
            font=("맑은 고딕", 16, "bold")
        )
        title_label.pack(pady=(0, 15))

        # 모드 선택
        mode_frame = ttk.LabelFrame(main_frame, text="변환 모드", padding="10")
        mode_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Radiobutton(
            mode_frame, text="단일 파일 변환",
            variable=self.batch_mode, value=False,
            command=self._toggle_mode
        ).pack(side=tk.LEFT, padx=(0, 20))

        ttk.Radiobutton(
            mode_frame, text="배치 변환 (폴더)",
            variable=self.batch_mode, value=True,
            command=self._toggle_mode
        ).pack(side=tk.LEFT)

        # === 단일 파일 모드 프레임 ===
        self.single_frame = ttk.Frame(main_frame)
        self.single_frame.pack(fill=tk.X, pady=(0, 10))

        # 입력 파일 선택
        input_frame = ttk.LabelFrame(self.single_frame, text="입력 파일", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Entry(input_frame, textvariable=self.input_path, width=60).pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        ttk.Button(input_frame, text="찾아보기", command=self._browse_input).pack(side=tk.LEFT)

        # 출력 위치 선택
        output_frame = ttk.LabelFrame(self.single_frame, text="출력 위치", padding="10")
        output_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Entry(output_frame, textvariable=self.output_path, width=60).pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        ttk.Button(output_frame, text="찾아보기", command=self._browse_output).pack(side=tk.LEFT)

        # === 배치 모드 프레임 ===
        self.batch_frame = ttk.Frame(main_frame)
        # 초기에는 숨김

        # 입력 폴더 선택
        input_folder_frame = ttk.LabelFrame(self.batch_frame, text="입력 폴더", padding="10")
        input_folder_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Entry(input_folder_frame, textvariable=self.input_folder, width=60).pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        ttk.Button(input_folder_frame, text="찾아보기", command=self._browse_input_folder).pack(side=tk.LEFT)

        # 출력 폴더 선택
        output_folder_frame = ttk.LabelFrame(self.batch_frame, text="출력 폴더", padding="10")
        output_folder_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Entry(output_folder_frame, textvariable=self.output_folder, width=60).pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        ttk.Button(output_folder_frame, text="찾아보기", command=self._browse_output_folder).pack(side=tk.LEFT)

        # 배치 옵션
        batch_options_frame = ttk.Frame(self.batch_frame)
        batch_options_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Checkbutton(
            batch_options_frame,
            text="하위 폴더 포함",
            variable=self.include_subfolders
        ).pack(side=tk.LEFT)

        ttk.Button(
            batch_options_frame,
            text="파일 목록 미리보기",
            command=self._preview_files
        ).pack(side=tk.LEFT, padx=(20, 0))

        # === 공통 옵션 ===
        options_frame = ttk.LabelFrame(main_frame, text="옵션", padding="10")
        options_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Checkbutton(
            options_frame,
            text="이미지 추출",
            variable=self.extract_images
        ).pack(side=tk.LEFT, padx=(0, 20))

        ttk.Checkbutton(
            options_frame,
            text="빈 문단 제거",
            variable=self.remove_empty
        ).pack(side=tk.LEFT)

        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 10))

        self.convert_btn = ttk.Button(
            button_frame,
            text="변환 실행",
            command=self._start_convert
        )
        self.convert_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.cancel_btn = ttk.Button(
            button_frame,
            text="취소",
            command=self._cancel_convert,
            state=tk.DISABLED
        )
        self.cancel_btn.pack(side=tk.LEFT)

        # 진행률 프레임
        progress_frame = ttk.LabelFrame(main_frame, text="진행 상황", padding="10")
        progress_frame.pack(fill=tk.X, pady=(0, 10))

        # 현재 파일 진행률
        self.current_file_label = ttk.Label(progress_frame, text="대기 중...")
        self.current_file_label.pack(anchor=tk.W)

        self.current_progress = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.current_progress.pack(fill=tk.X, pady=(5, 10))

        # 전체 진행률 (배치 모드용)
        self.total_progress_label = ttk.Label(progress_frame, text="전체: 0/0")
        self.total_progress_label.pack(anchor=tk.W)

        self.total_progress = ttk.Progressbar(progress_frame, mode='determinate')
        self.total_progress.pack(fill=tk.X, pady=(5, 0))

        # 로그 영역
        log_frame = ttk.LabelFrame(main_frame, text="로그", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 스크롤바 추가
        log_scroll = ttk.Scrollbar(log_frame)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text = tk.Text(log_frame, height=8, state=tk.DISABLED,
                                 yscrollcommand=log_scroll.set)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        log_scroll.config(command=self.log_text.yview)

        # 결과 요약 버튼
        result_frame = ttk.Frame(main_frame)
        result_frame.pack(fill=tk.X)

        self.export_btn = ttk.Button(
            result_frame,
            text="결과 CSV 내보내기",
            command=self._export_csv,
            state=tk.DISABLED
        )
        self.export_btn.pack(side=tk.LEFT)

    def _toggle_mode(self):
        """변환 모드 전환"""
        if self.batch_mode.get():
            self.single_frame.pack_forget()
            self.batch_frame.pack(fill=tk.X, pady=(0, 10),
                                   after=self.root.nametowidget(str(self.single_frame.master.winfo_children()[1])))
        else:
            self.batch_frame.pack_forget()
            self.single_frame.pack(fill=tk.X, pady=(0, 10),
                                    after=self.root.nametowidget(str(self.batch_frame.master.winfo_children()[1])))

    def _browse_input(self):
        """입력 파일 선택 대화상자"""
        filepath = filedialog.askopenfilename(
            title="Word 파일 선택",
            filetypes=[("Word 문서", "*.docx"), ("모든 파일", "*.*")]
        )
        if filepath:
            self.input_path.set(filepath)
            output = Path(filepath).with_suffix('.html')
            self.output_path.set(str(output))

    def _browse_output(self):
        """출력 위치 선택 대화상자"""
        filepath = filedialog.asksaveasfilename(
            title="저장 위치 선택",
            defaultextension=".html",
            filetypes=[("HTML 파일", "*.html"), ("모든 파일", "*.*")]
        )
        if filepath:
            self.output_path.set(filepath)

    def _browse_input_folder(self):
        """입력 폴더 선택"""
        folder = filedialog.askdirectory(title="입력 폴더 선택")
        if folder:
            self.input_folder.set(folder)
            if not self.output_folder.get():
                self.output_folder.set(folder)

    def _browse_output_folder(self):
        """출력 폴더 선택"""
        folder = filedialog.askdirectory(title="출력 폴더 선택")
        if folder:
            self.output_folder.set(folder)

    def _preview_files(self):
        """파일 목록 미리보기"""
        input_folder = self.input_folder.get()
        if not input_folder:
            messagebox.showwarning("경고", "입력 폴더를 선택하세요.")
            return

        files = find_docx_files(input_folder, self.include_subfolders.get())

        if not files:
            messagebox.showinfo("정보", "선택한 폴더에 .docx 파일이 없습니다.")
            return

        # 미리보기 창
        preview_win = tk.Toplevel(self.root)
        preview_win.title(f"파일 목록 ({len(files)}개)")
        preview_win.geometry("500x400")

        listbox_frame = ttk.Frame(preview_win, padding="10")
        listbox_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(listbox_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        listbox = tk.Listbox(listbox_frame, yscrollcommand=scrollbar.set)
        listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)

        for f in files:
            rel_path = f.relative_to(input_folder) if f.is_relative_to(Path(input_folder)) else f
            listbox.insert(tk.END, str(rel_path))

        ttk.Button(preview_win, text="닫기", command=preview_win.destroy).pack(pady=10)

    def _log(self, message):
        """로그 메시지 추가"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _clear_log(self):
        """로그 초기화"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _setup_progress_checker(self):
        """진행 상황 업데이트 체커"""
        try:
            while True:
                msg = self.progress_queue.get_nowait()
                msg_type = msg.get('type')

                if msg_type == 'file_start':
                    self.current_file_label.config(text=f"처리 중: {msg['filename']}")
                    self.current_progress.start(10)
                elif msg_type == 'file_done':
                    self.current_progress.stop()
                    status = "성공" if msg['success'] else f"실패: {msg.get('error', '')}"
                    self._log(f"[{status}] {msg['filename']}")
                elif msg_type == 'progress':
                    self.total_progress_label.config(text=f"전체: {msg['current']}/{msg['total']}")
                    self.total_progress['value'] = (msg['current'] / msg['total']) * 100
                elif msg_type == 'done':
                    self._conversion_done(msg['result'])
                elif msg_type == 'log':
                    self._log(msg['message'])

        except queue.Empty:
            pass

        self.root.after(100, self._setup_progress_checker)

    def _start_convert(self):
        """변환 시작"""
        if self.batch_mode.get():
            self._start_batch_convert()
        else:
            self._start_single_convert()

    def _start_single_convert(self):
        """단일 파일 변환 시작"""
        input_file = self.input_path.get()
        output_file = self.output_path.get()

        if not input_file:
            messagebox.showwarning("경고", "입력 파일을 선택하세요.")
            return

        if not output_file:
            messagebox.showwarning("경고", "출력 위치를 선택하세요.")
            return

        self._prepare_conversion()
        self._log(f"변환 시작: {input_file}")

        # 백그라운드 스레드에서 실행
        thread = threading.Thread(target=self._convert_single, args=(input_file, output_file))
        thread.daemon = True
        thread.start()

    def _convert_single(self, input_file, output_file):
        """단일 파일 변환 (스레드)"""
        options = {
            'extract_images': self.extract_images.get(),
            'remove_empty_paragraphs': self.remove_empty.get()
        }

        self.progress_queue.put({
            'type': 'file_start',
            'filename': Path(input_file).name
        })

        result = self.converter.convert(input_file, output_file, options)

        self.progress_queue.put({
            'type': 'file_done',
            'filename': Path(input_file).name,
            'success': result.success,
            'error': result.error_message
        })

        batch_result = BatchResult()
        batch_result.add(result)

        self.progress_queue.put({
            'type': 'done',
            'result': batch_result
        })

    def _start_batch_convert(self):
        """배치 변환 시작"""
        input_folder = self.input_folder.get()
        output_folder = self.output_folder.get()

        if not input_folder:
            messagebox.showwarning("경고", "입력 폴더를 선택하세요.")
            return

        if not output_folder:
            messagebox.showwarning("경고", "출력 폴더를 선택하세요.")
            return

        files = find_docx_files(input_folder, self.include_subfolders.get())

        if not files:
            messagebox.showinfo("정보", "변환할 .docx 파일이 없습니다.")
            return

        self._prepare_conversion()
        self._log(f"배치 변환 시작: {len(files)}개 파일")

        # 백그라운드 스레드에서 실행
        thread = threading.Thread(
            target=self._convert_batch,
            args=(files, input_folder, output_folder)
        )
        thread.daemon = True
        thread.start()

    def _convert_batch(self, files, input_folder, output_folder):
        """배치 변환 (스레드)"""
        options = {
            'extract_images': self.extract_images.get(),
            'remove_empty_paragraphs': self.remove_empty.get()
        }

        batch_result = BatchResult()
        input_folder = Path(input_folder)
        output_folder = Path(output_folder)

        for i, input_file in enumerate(files):
            if self.cancel_requested:
                self.progress_queue.put({'type': 'log', 'message': '변환이 취소되었습니다.'})
                break

            # 출력 경로 계산 (폴더 구조 유지)
            rel_path = input_file.relative_to(input_folder)
            output_file = output_folder / rel_path.with_suffix('.html')

            self.progress_queue.put({
                'type': 'file_start',
                'filename': input_file.name
            })

            self.progress_queue.put({
                'type': 'progress',
                'current': i,
                'total': len(files)
            })

            result = self.converter.convert(str(input_file), str(output_file), options)
            batch_result.add(result)

            self.progress_queue.put({
                'type': 'file_done',
                'filename': input_file.name,
                'success': result.success,
                'error': result.error_message
            })

        # 완료
        self.progress_queue.put({
            'type': 'progress',
            'current': len(files),
            'total': len(files)
        })

        self.progress_queue.put({
            'type': 'done',
            'result': batch_result
        })

    def _prepare_conversion(self):
        """변환 준비"""
        self.is_converting = True
        self.cancel_requested = False
        self._clear_log()
        self.convert_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        self.export_btn.config(state=tk.DISABLED)
        self.total_progress['value'] = 0

    def _conversion_done(self, batch_result):
        """변환 완료 처리"""
        self.is_converting = False
        self.batch_result = batch_result
        self.current_progress.stop()
        self.convert_btn.config(state=tk.NORMAL)
        self.cancel_btn.config(state=tk.DISABLED)

        summary = batch_result.get_summary()

        self._log("-" * 40)
        self._log(f"변환 완료!")
        self._log(f"  성공: {summary['success']}개")
        self._log(f"  실패: {summary['failed']}개")
        self._log(f"  경고: {summary['warnings']}개")

        if summary['failed'] > 0:
            self.export_btn.config(state=tk.NORMAL)
            messagebox.showwarning(
                "완료",
                f"변환 완료\n\n성공: {summary['success']}개\n실패: {summary['failed']}개\n\n자세한 내용은 로그를 확인하세요."
            )
        else:
            self.export_btn.config(state=tk.NORMAL)
            messagebox.showinfo("완료", f"모든 파일 변환 완료! ({summary['success']}개)")

    def _cancel_convert(self):
        """변환 취소"""
        if self.is_converting:
            self.cancel_requested = True
            self._log("취소 요청됨...")

    def _export_csv(self):
        """결과 CSV 내보내기"""
        if not self.batch_result:
            return

        filepath = filedialog.asksaveasfilename(
            title="CSV 저장",
            defaultextension=".csv",
            filetypes=[("CSV 파일", "*.csv")]
        )

        if filepath:
            try:
                self.batch_result.export_csv(filepath)
                messagebox.showinfo("완료", f"CSV 파일 저장 완료:\n{filepath}")
            except Exception as e:
                messagebox.showerror("오류", f"저장 실패: {e}")

    def run(self):
        """애플리케이션 실행"""
        self.root.mainloop()


# 테스트용
if __name__ == "__main__":
    app = ConverterApp()
    app.run()
