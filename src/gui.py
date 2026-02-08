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
from pdf_converter import PdfConverter
from utils import find_docx_files, find_convertible_files, BatchResult


class ConverterApp:
    """변환기 GUI 애플리케이션"""

    # 색상 테마
    COLORS = {
        'bg': '#f5f5f5',
        'frame_bg': '#ffffff',
        'primary': '#2563eb',
        'primary_hover': '#1d4ed8',
        'success': '#16a34a',
        'danger': '#dc2626',
        'text': '#1f2937',
        'text_secondary': '#6b7280',
        'border': '#e5e7eb',
    }

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Document to HTML Converter")
        self.root.geometry("750x700")
        self.root.resizable(True, True)
        self.root.minsize(650, 600)
        self.root.configure(bg=self.COLORS['bg'])

        # 스타일 설정
        self._setup_styles()

        # 변환기 인스턴스
        self.converter = DocxConverter()
        self.pdf_converter = PdfConverter()

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

    def _setup_styles(self):
        """ttk 스타일 설정"""
        style = ttk.Style()

        # 테마 설정 (clam이 가장 커스터마이징 하기 좋음)
        style.theme_use('clam')

        # 기본 프레임
        style.configure(
            'TFrame',
            background=self.COLORS['bg']
        )

        # 카드 스타일 프레임
        style.configure(
            'Card.TFrame',
            background=self.COLORS['frame_bg'],
            relief='flat'
        )

        # 레이블
        style.configure(
            'TLabel',
            background=self.COLORS['bg'],
            foreground=self.COLORS['text'],
            font=('맑은 고딕', 10)
        )

        # 제목 레이블
        style.configure(
            'Title.TLabel',
            background=self.COLORS['bg'],
            foreground=self.COLORS['primary'],
            font=('맑은 고딕', 18, 'bold')
        )

        # 부제목 레이블
        style.configure(
            'Subtitle.TLabel',
            background=self.COLORS['bg'],
            foreground=self.COLORS['text_secondary'],
            font=('맑은 고딕', 9)
        )

        # LabelFrame
        style.configure(
            'TLabelframe',
            background=self.COLORS['frame_bg'],
            relief='solid',
            borderwidth=1,
            bordercolor=self.COLORS['border']
        )
        style.configure(
            'TLabelframe.Label',
            background=self.COLORS['frame_bg'],
            foreground=self.COLORS['text'],
            font=('맑은 고딕', 10, 'bold')
        )

        # 기본 버튼
        style.configure(
            'TButton',
            font=('맑은 고딕', 10),
            padding=(15, 8),
            relief='flat'
        )
        style.map('TButton',
            background=[('active', self.COLORS['border']), ('!active', self.COLORS['frame_bg'])],
            relief=[('pressed', 'flat'), ('!pressed', 'flat')]
        )

        # 주요 버튼 (변환 실행)
        style.configure(
            'Primary.TButton',
            font=('맑은 고딕', 11, 'bold'),
            padding=(30, 12),
            background=self.COLORS['primary'],
            foreground='white'
        )
        style.map('Primary.TButton',
            background=[('active', self.COLORS['primary_hover']), ('!active', self.COLORS['primary'])],
            foreground=[('active', 'white'), ('!active', 'white')]
        )

        # 위험 버튼 (취소)
        style.configure(
            'Danger.TButton',
            font=('맑은 고딕', 11),
            padding=(30, 12),
            background=self.COLORS['danger'],
            foreground='white'
        )
        style.map('Danger.TButton',
            background=[('active', '#b91c1c'), ('!active', self.COLORS['danger'])],
            foreground=[('active', 'white'), ('!active', 'white')]
        )

        # 입력 필드
        style.configure(
            'TEntry',
            font=('맑은 고딕', 10),
            padding=8,
            relief='solid',
            borderwidth=1
        )

        # 체크버튼
        style.configure(
            'TCheckbutton',
            background=self.COLORS['frame_bg'],
            foreground=self.COLORS['text'],
            font=('맑은 고딕', 10)
        )

        # 라디오버튼
        style.configure(
            'TRadiobutton',
            background=self.COLORS['frame_bg'],
            foreground=self.COLORS['text'],
            font=('맑은 고딕', 10)
        )

        # 프로그레스바
        style.configure(
            'TProgressbar',
            background=self.COLORS['primary'],
            troughcolor=self.COLORS['border'],
            borderwidth=0,
            lightcolor=self.COLORS['primary'],
            darkcolor=self.COLORS['primary']
        )

        # 성공 프로그레스바
        style.configure(
            'green.Horizontal.TProgressbar',
            background=self.COLORS['success'],
            troughcolor=self.COLORS['border']
        )

    def _create_widgets(self):
        """GUI 위젯 생성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 헤더 영역
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))

        title_label = ttk.Label(
            header_frame,
            text="DOCX / PDF → HTML Converter",
            style='Title.TLabel'
        )
        title_label.pack(anchor=tk.W)

        subtitle_label = ttk.Label(
            header_frame,
            text="Word/PDF 문서를 웹북용 HTML로 변환합니다",
            style='Subtitle.TLabel'
        )
        subtitle_label.pack(anchor=tk.W, pady=(5, 0))

        # 모드 선택
        mode_frame = ttk.LabelFrame(main_frame, text=" 변환 모드 ", padding="15")
        mode_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Radiobutton(
            mode_frame, text="단일 파일 변환",
            variable=self.batch_mode, value=False,
            command=self._toggle_mode
        ).pack(side=tk.LEFT, padx=(0, 30))

        ttk.Radiobutton(
            mode_frame, text="배치 변환 (폴더 전체)",
            variable=self.batch_mode, value=True,
            command=self._toggle_mode
        ).pack(side=tk.LEFT)

        # === 단일 파일 모드 프레임 ===
        self.single_frame = ttk.Frame(main_frame)
        self.single_frame.pack(fill=tk.X, pady=(0, 15))

        # 입력 파일 선택
        input_frame = ttk.LabelFrame(self.single_frame, text=" 입력 파일 ", padding="15")
        input_frame.pack(fill=tk.X, pady=(0, 10))

        input_entry = ttk.Entry(input_frame, textvariable=self.input_path, font=('맑은 고딕', 10))
        input_entry.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        ttk.Button(input_frame, text="파일 선택", command=self._browse_input).pack(side=tk.LEFT)

        # 출력 위치 선택
        output_frame = ttk.LabelFrame(self.single_frame, text=" 출력 위치 ", padding="15")
        output_frame.pack(fill=tk.X)

        output_entry = ttk.Entry(output_frame, textvariable=self.output_path, font=('맑은 고딕', 10))
        output_entry.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        ttk.Button(output_frame, text="위치 선택", command=self._browse_output).pack(side=tk.LEFT)

        # === 배치 모드 프레임 ===
        self.batch_frame = ttk.Frame(main_frame)
        # 초기에는 숨김

        # 입력 폴더 선택
        input_folder_frame = ttk.LabelFrame(self.batch_frame, text=" 입력 폴더 ", padding="15")
        input_folder_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Entry(input_folder_frame, textvariable=self.input_folder, font=('맑은 고딕', 10)).pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        ttk.Button(input_folder_frame, text="폴더 선택", command=self._browse_input_folder).pack(side=tk.LEFT)

        # 출력 폴더 선택
        output_folder_frame = ttk.LabelFrame(self.batch_frame, text=" 출력 폴더 ", padding="15")
        output_folder_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Entry(output_folder_frame, textvariable=self.output_folder, font=('맑은 고딕', 10)).pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        ttk.Button(output_folder_frame, text="폴더 선택", command=self._browse_output_folder).pack(side=tk.LEFT)

        # 배치 옵션
        batch_options_frame = ttk.LabelFrame(self.batch_frame, text=" 배치 옵션 ", padding="15")
        batch_options_frame.pack(fill=tk.X)

        ttk.Checkbutton(
            batch_options_frame,
            text="하위 폴더 포함",
            variable=self.include_subfolders
        ).pack(side=tk.LEFT)

        ttk.Button(
            batch_options_frame,
            text="파일 목록 미리보기",
            command=self._preview_files
        ).pack(side=tk.RIGHT)

        # === 공통 옵션 ===
        options_frame = ttk.LabelFrame(main_frame, text=" 변환 옵션 ", padding="15")
        options_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Checkbutton(
            options_frame,
            text="이미지 추출",
            variable=self.extract_images
        ).pack(side=tk.LEFT, padx=(0, 30))

        ttk.Checkbutton(
            options_frame,
            text="빈 문단 제거",
            variable=self.remove_empty
        ).pack(side=tk.LEFT)

        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 20))

        # 버튼들을 담을 내부 프레임 (중앙 정렬용)
        button_inner = ttk.Frame(button_frame)
        button_inner.pack(anchor=tk.CENTER)

        self.convert_btn = ttk.Button(
            button_inner,
            text="▶  변환 실행",
            style='Primary.TButton',
            command=self._start_convert,
            width=15
        )
        self.convert_btn.pack(side=tk.LEFT, padx=(0, 15))

        self.cancel_btn = ttk.Button(
            button_inner,
            text="■  취소",
            style='Danger.TButton',
            command=self._cancel_convert,
            state=tk.DISABLED,
            width=15
        )
        self.cancel_btn.pack(side=tk.LEFT)

        # 진행률 프레임
        progress_frame = ttk.LabelFrame(main_frame, text=" 진행 상황 ", padding="15")
        progress_frame.pack(fill=tk.X, pady=(0, 15))

        # 현재 파일 진행률
        self.current_file_label = ttk.Label(progress_frame, text="대기 중...", foreground=self.COLORS['text_secondary'])
        self.current_file_label.pack(anchor=tk.W)

        self.current_progress = ttk.Progressbar(progress_frame, mode='indeterminate', length=400)
        self.current_progress.pack(fill=tk.X, pady=(8, 12))

        # 전체 진행률 (배치 모드용)
        self.total_progress_label = ttk.Label(progress_frame, text="전체: 0/0", foreground=self.COLORS['text_secondary'])
        self.total_progress_label.pack(anchor=tk.W)

        self.total_progress = ttk.Progressbar(progress_frame, mode='determinate', length=400, style='green.Horizontal.TProgressbar')
        self.total_progress.pack(fill=tk.X, pady=(8, 0))

        # 로그 영역
        log_frame = ttk.LabelFrame(main_frame, text=" 로그 ", padding="15")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # 스크롤바 추가
        log_scroll = ttk.Scrollbar(log_frame)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text = tk.Text(
            log_frame,
            height=8,
            state=tk.DISABLED,
            yscrollcommand=log_scroll.set,
            font=('Consolas', 9),
            bg='#1f2937',
            fg='#e5e7eb',
            insertbackground='white',
            relief='flat',
            padx=10,
            pady=10
        )
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
            self.batch_frame.pack(fill=tk.X, pady=(0, 15),
                                   after=self.root.nametowidget(str(self.single_frame.master.winfo_children()[1])))
        else:
            self.batch_frame.pack_forget()
            self.single_frame.pack(fill=tk.X, pady=(0, 15),
                                    after=self.root.nametowidget(str(self.batch_frame.master.winfo_children()[1])))

    def _browse_input(self):
        """입력 파일 선택 대화상자"""
        filepath = filedialog.askopenfilename(
            title="문서 파일 선택",
            filetypes=[("지원 문서", "*.docx *.pdf"), ("Word 문서", "*.docx"), ("PDF 문서", "*.pdf"), ("모든 파일", "*.*")]
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

        files = find_convertible_files(input_folder, self.include_subfolders.get())

        if not files:
            messagebox.showinfo("정보", "선택한 폴더에 변환 가능한 파일이 없습니다. (.docx, .pdf)")
            return

        # 미리보기 창
        preview_win = tk.Toplevel(self.root)
        preview_win.title(f"파일 목록 ({len(files)}개)")
        preview_win.geometry("550x450")
        preview_win.configure(bg=self.COLORS['bg'])

        listbox_frame = ttk.Frame(preview_win, padding="15")
        listbox_frame.pack(fill=tk.BOTH, expand=True)

        header_label = ttk.Label(listbox_frame, text=f"총 {len(files)}개 파일", style='Subtitle.TLabel')
        header_label.pack(anchor=tk.W, pady=(0, 10))

        scrollbar = ttk.Scrollbar(listbox_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        listbox = tk.Listbox(
            listbox_frame,
            yscrollcommand=scrollbar.set,
            font=('맑은 고딕', 9),
            bg=self.COLORS['frame_bg'],
            relief='solid',
            borderwidth=1,
            selectbackground=self.COLORS['primary']
        )
        listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)

        for f in files:
            rel_path = f.relative_to(input_folder) if f.is_relative_to(Path(input_folder)) else f
            listbox.insert(tk.END, str(rel_path))

        ttk.Button(preview_win, text="닫기", command=preview_win.destroy).pack(pady=15)

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
                    status = "✓" if msg['success'] else f"✗ {msg.get('error', '')}"
                    self._log(f"  [{status}] {msg['filename']}")
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
        self._log(f"변환 시작: {Path(input_file).name}")

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

        if Path(input_file).suffix.lower() == '.pdf':
            result = self.pdf_converter.convert(input_file, output_file, options)
        else:
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

        files = find_convertible_files(input_folder, self.include_subfolders.get())

        if not files:
            messagebox.showinfo("정보", "변환할 문서 파일이 없습니다. (.docx, .pdf)")
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

            if input_file.suffix.lower() == '.pdf':
                result = self.pdf_converter.convert(str(input_file), str(output_file), options)
            else:
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
        self.current_file_label.config(text="완료!")
        self.convert_btn.config(state=tk.NORMAL)
        self.cancel_btn.config(state=tk.DISABLED)

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
            self.export_btn.config(state=tk.NORMAL)
            messagebox.showwarning(
                "완료",
                f"변환 완료\n\n✓ 성공: {summary['success']}개\n✗ 실패: {summary['failed']}개\n\n자세한 내용은 로그를 확인하세요."
            )
        else:
            self.export_btn.config(state=tk.NORMAL)
            messagebox.showinfo("완료", f"모든 파일 변환 완료!\n\n✓ 성공: {summary['success']}개")

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
        # 창을 화면 중앙에 배치
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

        self.root.mainloop()


# 테스트용
if __name__ == "__main__":
    app = ConverterApp()
    app.run()
