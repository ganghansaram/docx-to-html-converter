#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
DOCX to HTML Converter - GUI
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path


class ConverterApp:
    """변환기 GUI 애플리케이션"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("DOCX to HTML Converter")
        self.root.geometry("500x400")
        self.root.resizable(False, False)

        # 변수 초기화
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.extract_images = tk.BooleanVar(value=True)
        self.remove_empty = tk.BooleanVar(value=True)

        self._create_widgets()

    def _create_widgets(self):
        """GUI 위젯 생성"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 제목
        title_label = ttk.Label(
            main_frame,
            text="Word → HTML 변환기",
            font=("맑은 고딕", 16, "bold")
        )
        title_label.pack(pady=(0, 20))

        # 입력 파일 선택
        input_frame = ttk.LabelFrame(main_frame, text="입력 파일", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Entry(input_frame, textvariable=self.input_path, width=50).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(input_frame, text="찾아보기", command=self._browse_input).pack(side=tk.LEFT)

        # 출력 위치 선택
        output_frame = ttk.LabelFrame(main_frame, text="출력 위치", padding="10")
        output_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Entry(output_frame, textvariable=self.output_path, width=50).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(output_frame, text="찾아보기", command=self._browse_output).pack(side=tk.LEFT)

        # 옵션
        options_frame = ttk.LabelFrame(main_frame, text="옵션", padding="10")
        options_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Checkbutton(
            options_frame,
            text="이미지 추출",
            variable=self.extract_images
        ).pack(anchor=tk.W)

        ttk.Checkbutton(
            options_frame,
            text="빈 문단 제거",
            variable=self.remove_empty
        ).pack(anchor=tk.W)

        # 변환 버튼
        convert_btn = ttk.Button(
            main_frame,
            text="변환 실행",
            command=self._convert
        )
        convert_btn.pack(pady=20)

        # 상태 표시
        self.status_var = tk.StringVar(value="대기 중...")
        status_label = ttk.Label(main_frame, textvariable=self.status_var)
        status_label.pack()

        # 로그 영역
        log_frame = ttk.LabelFrame(main_frame, text="로그", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        self.log_text = tk.Text(log_frame, height=6, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _browse_input(self):
        """입력 파일 선택 대화상자"""
        filepath = filedialog.askopenfilename(
            title="Word 파일 선택",
            filetypes=[("Word 문서", "*.docx"), ("모든 파일", "*.*")]
        )
        if filepath:
            self.input_path.set(filepath)
            # 출력 경로 자동 설정
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

    def _log(self, message):
        """로그 메시지 추가"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _convert(self):
        """변환 실행"""
        input_file = self.input_path.get()
        output_file = self.output_path.get()

        if not input_file:
            messagebox.showwarning("경고", "입력 파일을 선택하세요.")
            return

        if not output_file:
            messagebox.showwarning("경고", "출력 위치를 선택하세요.")
            return

        self.status_var.set("변환 중...")
        self._log(f"변환 시작: {input_file}")

        # TODO: 실제 변환 로직 호출
        # from converter import DocxConverter
        # converter = DocxConverter()
        # result = converter.convert(input_file, output_file)

        self._log("변환 완료!")
        self.status_var.set("완료!")
        messagebox.showinfo("완료", "변환이 완료되었습니다.")

    def run(self):
        """애플리케이션 실행"""
        self.root.mainloop()


# 테스트용
if __name__ == "__main__":
    app = ConverterApp()
    app.run()
