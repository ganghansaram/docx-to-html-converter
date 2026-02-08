#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
DOCX to HTML Converter - 메인 엔트리 포인트
"""

import sys
from pathlib import Path

# src 폴더를 모듈 경로에 추가
src_dir = Path(__file__).parent
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))


def main():
    """GUI 애플리케이션 실행"""
    try:
        from gui import ConverterApp
        app = ConverterApp()
        app.run()
    except ImportError as e:
        print(f"오류: 필요한 패키지가 설치되지 않았습니다.")
        print(f"상세: {e}")
        print(f"\n다음 명령어로 패키지를 설치하세요:")
        print(f"  pip install python-docx PyMuPDF")
        sys.exit(1)


def cli():
    """CLI 모드 실행 (향후 확장용)"""
    import argparse

    parser = argparse.ArgumentParser(
        description='Document to HTML Converter - Word/PDF 문서를 HTML로 변환'
    )
    parser.add_argument('input', nargs='?', help='입력 파일 또는 폴더')
    parser.add_argument('-o', '--output', help='출력 파일 또는 폴더')
    parser.add_argument('-r', '--recursive', action='store_true',
                        help='하위 폴더 포함 (배치 모드)')
    parser.add_argument('--no-images', action='store_true',
                        help='이미지 추출 안 함')
    parser.add_argument('--analyze', action='store_true',
                        help='문서 구조 분석만 수행')
    parser.add_argument('--gui', action='store_true',
                        help='GUI 모드 실행')

    args = parser.parse_args()

    # GUI 모드
    if args.gui or not args.input:
        main()
        return

    # CLI 모드
    try:
        from converter import DocxConverter
        from pdf_converter import PdfConverter
        from utils import find_convertible_files, BatchResult
        import json

        docx_converter = DocxConverter()
        pdf_converter = PdfConverter()
        input_path = Path(args.input)

        options = {
            'extract_images': not args.no_images
        }

        def get_converter(filepath):
            """확장자에 따라 적절한 변환기 반환"""
            if Path(filepath).suffix.lower() == '.pdf':
                return pdf_converter
            return docx_converter

        if args.analyze:
            # 분석 모드
            if input_path.is_file():
                converter = get_converter(input_path)
                result = converter.analyze(str(input_path))
                print(json.dumps(result, indent=2, ensure_ascii=False))
            else:
                print("오류: 분석 모드는 단일 파일만 지원합니다.")
                sys.exit(1)
        elif input_path.is_file():
            # 단일 파일 변환
            output_path = args.output or str(input_path.with_suffix('.html'))
            converter = get_converter(input_path)
            result = converter.convert(str(input_path), output_path, options)

            if result.success:
                print(f"변환 완료: {result.output_path}")
                if result.warnings:
                    print(f"경고: {', '.join(result.warnings)}")
            else:
                print(f"변환 실패: {result.error_message}")
                sys.exit(1)
        elif input_path.is_dir():
            # 배치 변환
            files = find_convertible_files(str(input_path), args.recursive)
            output_dir = Path(args.output) if args.output else input_path

            if not files:
                print("변환할 문서 파일이 없습니다. (.docx, .pdf)")
                sys.exit(1)

            print(f"{len(files)}개 파일 변환 시작...")
            batch_result = BatchResult()

            for f in files:
                rel_path = f.relative_to(input_path)
                output_path = output_dir / rel_path.with_suffix('.html')

                converter = get_converter(f)
                result = converter.convert(str(f), str(output_path), options)
                batch_result.add(result)

                status = "OK" if result.success else "FAIL"
                print(f"  [{status}] {f.name}")

            summary = batch_result.get_summary()
            print(f"\n완료: 성공 {summary['success']}, 실패 {summary['failed']}")
        else:
            print(f"오류: 파일 또는 폴더를 찾을 수 없습니다: {input_path}")
            sys.exit(1)

    except ImportError as e:
        print(f"오류: {e}")
        print("pip install python-docx PyMuPDF 를 실행하세요.")
        sys.exit(1)


if __name__ == "__main__":
    # 기본적으로 GUI 모드로 시작
    # CLI 인자가 있으면 CLI 모드
    if len(sys.argv) > 1:
        cli()
    else:
        main()
