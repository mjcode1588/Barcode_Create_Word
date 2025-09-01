import os
import sys
from barcode import Code128
from barcode.writer import ImageWriter
from typing import List, Tuple, Optional, Dict
from io import BytesIO
from PIL import Image
from src.services.log_service import logger
from PIL import ImageDraw, ImageFont


class BarcodeGenerator:
    """GS1-128 바코드 생성 클래스 (메모리 기반) - MM 단위 통일"""

    def __init__(self, options: Optional[Dict] = None):
        if options is None:
            options = {}

        # MM 단위로 입력받은 값들을 바코드 라이브러리에 맞게 변환
        self.writer_options = self._convert_mm_to_barcode_units(options)

    def _convert_mm_to_barcode_units(self, options: Dict) -> Dict:
        """MM 단위 입력값을 바코드 라이브러리 단위로 변환 (안전한 범위 내에서 적용)"""
        # 사용자 설정값 (MM 단위)
        module_width_mm = options.get("module_width", 0.3)
        module_height_mm = options.get("module_height", 15.0)
        quiet_zone_mm = options.get("quiet_zone", 3.0)
        text_distance_mm = options.get("text_distance", 5.0)

        # 바코드 라이브러리 안전 범위 내에서만 적용
        # module_width와 quiet_zone은 라이브러리 기본값 사용 (단위 문제 회피)
        # module_height와 text_distance는 MM 단위로 사용자 값 적용

        return {
            "module_width": 0.35,  # 더 넓은 바코드를 위해 증가 (0.2 → 0.35)
            "module_height": max(
                10.0, min(25.0, module_height_mm)
            ),  # 사용자 값 적용 (10-25mm 범위)
            "quiet_zone": 6.5,  # 라이브러리 기본값 (안정성 우선)
            "text_distance": max(
                3.0, min(8.0, text_distance_mm)
            ),  # 사용자 값 적용 (3-8mm 범위)
            "font_size": options.get("font_size", 10),
            "dpi": options.get("dpi", 300),
        }

    def generate_barcode_in_memory(self, code: str, category: str) -> BytesIO:
        """바코드 이미지를 메모리에 생성 (파일 기반으로 변경)"""
        try:
            # 바코드 데이터가 ASCII 문자만 포함하는지 확인
            try:
                code.encode("ascii")
            except UnicodeEncodeError:
                logger.warning(
                    "BarcodeGenerator",
                    f"바코드 데이터에 ASCII가 아닌 문자가 포함됨: {code}",
                )
                # ASCII가 아닌 문자를 제거하거나 변환
                code = "".join(c for c in code if ord(c) < 128)
                logger.info("BarcodeGenerator", f"ASCII 문자만 추출: {code}")

            if not code:
                logger.error("BarcodeGenerator", "유효한 바코드 데이터가 없음")
                return None

            # 임시 파일 경로 생성
            import tempfile

            temp_dir = tempfile.gettempdir()
            temp_filename = os.path.join(temp_dir, f"barcode_{code}_{os.getpid()}.png")

            logger.debug("BarcodeGenerator", f"임시 바코드 파일 생성: {temp_filename}")

            # 파일로 바코드 생성
            success = self._generate_barcode_file(code, temp_filename)

            if success and os.path.exists(temp_filename):
                # 파일을 읽어서 메모리로 로드
                try:
                    with open(temp_filename, "rb") as f:
                        img_data = f.read()

                    # 임시 파일 삭제
                    try:
                        os.remove(temp_filename)
                    except:
                        pass  # 삭제 실패해도 무시

                    # BytesIO로 변환
                    img_buffer = BytesIO(img_data)
                    img_buffer.seek(0)

                    logger.debug(
                        "BarcodeGenerator", f"바코드 생성 완료: {code} ({category})"
                    )
                    return img_buffer

                except Exception as e:
                    logger.error(
                        "BarcodeGenerator", f"바코드 파일 읽기 실패: {code} - {e}"
                    )
                    # 임시 파일 정리
                    try:
                        os.remove(temp_filename)
                    except:
                        pass
                    return None
            else:
                logger.error("BarcodeGenerator", f"바코드 파일 생성 실패: {code}")
                return None

        except Exception as e:
            logger.error("BarcodeGenerator", f"바코드 생성 실패: {code} - {e}")
            return None

    def _generate_barcode_file(self, code: str, filename: str) -> bool:
        """바코드를 파일로 생성 (여러 방법 시도)"""

        # 방법 2: 기본 옵션으로 시도 (폰트 문제 무시)
        try:
            logger.debug("BarcodeGenerator", f"기본 옵션으로 바코드 생성 시도: {code}")
            writer = ImageWriter(format="PNG")
            barcode = Code128(code, writer=writer)

            basic_options = {
                "dpi": 300,
                "module_width": 0.35,
                "module_height": 15.0,
                "quiet_zone": 6.5,
                "text_distance": 5.0,
                "font_size": 10,
                "write_text": True,
            }

            barcode.write(filename, basic_options)

            if os.path.exists(filename) and os.path.getsize(filename) > 0:
                logger.info(
                    "BarcodeGenerator", f"기본 옵션으로 바코드 생성 성공: {code}"
                )
                return True

        except Exception as e:
            logger.warning("BarcodeGenerator", f"기본 옵션 실패: {e}")

        # 방법 1: 바코드만 생성 후 PIL로 텍스트 추가 (가장 안정적)
        try:
            logger.debug(
                "BarcodeGenerator", f"바코드+텍스트 조합으로 생성 시도: {code}"
            )

            # 먼저 텍스트 없는 바코드 생성
            temp_barcode_file = filename.replace(".png", "_temp.png")
            writer = ImageWriter(format="PNG")
            barcode = Code128(code, writer=writer)

            no_text_options = {
                "dpi": 300,
                "module_width": 0.35,
                "module_height": 15.0,
                "quiet_zone": 6.5,
                "write_text": False,  # 텍스트 없이
            }

            barcode.write(temp_barcode_file, no_text_options)

            if (
                os.path.exists(temp_barcode_file)
                and os.path.getsize(temp_barcode_file) > 0
            ):
                # PIL로 텍스트 추가 (바코드 길이에 맞춰 폰트 크기 조정)
                success = self._add_text_to_barcode(temp_barcode_file, filename, code)

                # 임시 파일 삭제
                try:
                    os.remove(temp_barcode_file)
                except:
                    pass

                if success:
                    logger.info(
                        "BarcodeGenerator", f"바코드+텍스트 조합으로 생성 성공: {code}"
                    )
                    return True

        except Exception as e:
            logger.warning("BarcodeGenerator", f"바코드+텍스트 조합 실패: {e}")

        # 방법 2: 텍스트 없이 시도
        try:
            logger.debug(
                "BarcodeGenerator", f"텍스트 없는 옵션으로 바코드 생성 시도: {code}"
            )
            writer = ImageWriter(format="PNG")
            barcode = Code128(code, writer=writer)

            no_text_options = {
                "dpi": 200,
                "module_width": 0.3,
                "module_height": 12.0,
                "quiet_zone": 5.0,
                "write_text": False,  # 텍스트 없이
            }

            barcode.write(filename, no_text_options)

            if os.path.exists(filename) and os.path.getsize(filename) > 0:
                logger.info(
                    "BarcodeGenerator", f"텍스트 없는 옵션으로 바코드 생성 성공: {code}"
                )
                return True

        except Exception as e:
            logger.warning("BarcodeGenerator", f"텍스트 없는 옵션 실패: {e}")

        # 방법 3: 최소 옵션으로 시도
        try:
            logger.debug("BarcodeGenerator", f"최소 옵션으로 바코드 생성 시도: {code}")
            writer = ImageWriter(format="PNG")
            barcode = Code128(code, writer=writer)

            minimal_options = {
                "dpi": 150,
                "write_text": False,
            }

            barcode.write(filename, minimal_options)

            if os.path.exists(filename) and os.path.getsize(filename) > 0:
                logger.info(
                    "BarcodeGenerator", f"최소 옵션으로 바코드 생성 성공: {code}"
                )
                return True

        except Exception as e:
            logger.warning("BarcodeGenerator", f"최소 옵션 실패: {e}")

        # 방법 4: PIL로 텍스트 이미지 생성
        try:
            logger.debug("BarcodeGenerator", f"텍스트 이미지로 대체 생성: {code}")
            return self._create_text_image_file(code, filename)

        except Exception as e:
            logger.error("BarcodeGenerator", f"텍스트 이미지 생성 실패: {e}")

        return False

    def _add_text_to_barcode(
        self, barcode_file: str, output_file: str, code: str
    ) -> bool:
        """바코드 이미지에 텍스트 추가 (동적 폰트 크기 조정)"""
        try:
            from PIL import Image, ImageDraw, ImageFont

            # 바코드 이미지 로드
            barcode_img = Image.open(barcode_file)
            barcode_width, barcode_height = barcode_img.size

            logger.debug(
                "BarcodeGenerator",
                f"원본 바코드 크기: {barcode_width}x{barcode_height}",
            )

            # 바코드가 너무 작은 경우 최소 크기 보장
            min_width = 300
            min_height = 80

            if barcode_width < min_width or barcode_height < min_height:
                # 비율을 유지하면서 크기 조정
                scale_factor = max(
                    min_width / barcode_width, min_height / barcode_height
                )
                new_width = int(barcode_width * scale_factor)
                new_height = int(barcode_height * scale_factor)

                barcode_img = barcode_img.resize(
                    (new_width, new_height), Image.Resampling.LANCZOS
                )
                barcode_width, barcode_height = new_width, new_height

                logger.debug(
                    "BarcodeGenerator",
                    f"바코드 크기 조정: {barcode_width}x{barcode_height} (scale: {scale_factor:.2f})",
                )

            # 바코드 크기에 비례한 폰트 크기 계산 (더 큰 텍스트)
            # 바코드 너비를 기준으로 적절한 폰트 크기 결정
            base_font_size = max(
                18, int(barcode_width / 16)
            )  # 최소 18, 바코드 너비의 1/16 (더 큰 비율)
            max_font_size = int(barcode_width / 8)  # 최대 크기 제한도 더 크게
            font_size = min(base_font_size, max_font_size)

            logger.debug(
                "BarcodeGenerator",
                f"계산된 폰트 크기: {font_size} (바코드 너비: {barcode_width})",
            )

            # 텍스트 높이를 폰트 크기에 비례하여 계산
            text_height = int(font_size * 2.5)  # 폰트 크기의 2.5배
            margin = max(5, int(font_size * 0.5))  # 여백도 폰트 크기에 비례
            total_height = barcode_height + text_height + margin

            # 새 이미지 생성 (바코드 + 텍스트 공간)
            new_img = Image.new("RGB", (barcode_width, total_height), color="white")

            # 바코드 이미지 붙여넣기
            new_img.paste(barcode_img, (0, 0))

            # 텍스트 그리기
            draw = ImageDraw.Draw(new_img)

            # 폰트 로드 시도 (동적 크기 적용)
            font = None
            font_paths = [
                "C:/Windows/Fonts/arial.ttf",
                "C:/Windows/Fonts/calibri.ttf",
                "arial.ttf",
                "calibri.ttf",
            ]

            # 시스템 폰트로 시도
            for font_path in font_paths:
                try:
                    font = ImageFont.truetype(font_path, font_size)
                    logger.debug(
                        "BarcodeGenerator",
                        f"폰트 로드 성공: {font_path}, 크기: {font_size}",
                    )
                    break
                except Exception as e:
                    logger.debug(
                        "BarcodeGenerator", f"폰트 로드 실패: {font_path} - {e}"
                    )
                    continue

            # 시스템 폰트 실패 시 기본 폰트 사용
            if font is None:
                try:
                    font = ImageFont.load_default()
                    logger.debug("BarcodeGenerator", "기본 폰트 사용")
                except:
                    logger.warning("BarcodeGenerator", "모든 폰트 로드 실패")

            # 텍스트 크기 측정 및 위치 계산
            if font:
                # 정확한 텍스트 크기 측정
                bbox = draw.textbbox((0, 0), code, font=font)
                text_width = bbox[2] - bbox[0]
                text_actual_height = bbox[3] - bbox[1]

                logger.debug(
                    "BarcodeGenerator",
                    f"텍스트 크기 측정: {text_width}x{text_actual_height}",
                )

                # 텍스트가 바코드보다 넓은 경우 폰트 크기 조정
                if text_width > barcode_width * 0.9:  # 바코드 너비의 90% 이내로 제한
                    scale_down = (barcode_width * 0.9) / text_width
                    new_font_size = max(8, int(font_size * scale_down))

                    logger.debug(
                        "BarcodeGenerator",
                        f"폰트 크기 재조정: {font_size} → {new_font_size}",
                    )

                    # 폰트 다시 로드
                    for font_path in font_paths:
                        try:
                            font = ImageFont.truetype(font_path, new_font_size)
                            break
                        except:
                            continue

                    if font is None:
                        font = ImageFont.load_default()

                    # 텍스트 크기 재측정
                    bbox = draw.textbbox((0, 0), code, font=font)
                    text_width = bbox[2] - bbox[0]
                    text_actual_height = bbox[3] - bbox[1]

            else:
                # 폰트가 없는 경우 대략적인 계산
                text_width = len(code) * (font_size * 0.6)
                text_actual_height = font_size
                logger.debug(
                    "BarcodeGenerator",
                    f"폰트 없음 - 추정 텍스트 크기: {text_width}x{text_actual_height}",
                )

            # 텍스트 위치 계산 (중앙 정렬)
            text_x = max(0, (barcode_width - text_width) // 2)
            text_y = barcode_height + (margin // 2)

            logger.debug("BarcodeGenerator", f"텍스트 위치: ({text_x}, {text_y})")

            # 텍스트 그리기
            draw.text((text_x, text_y), code, fill="black", font=font)

            # 파일 저장
            new_img.save(output_file, format="PNG")

            if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
                logger.info(
                    "BarcodeGenerator",
                    f"바코드에 텍스트 추가 완료: {code} (폰트크기: {font_size if 'new_font_size' not in locals() else new_font_size})",
                )
                return True
            else:
                logger.error(
                    "BarcodeGenerator", f"바코드+텍스트 파일 저장 실패: {output_file}"
                )
                return False

        except Exception as e:
            logger.error("BarcodeGenerator", f"바코드에 텍스트 추가 실패: {code} - {e}")
            import traceback

            logger.error("BarcodeGenerator", f"상세 오류: {traceback.format_exc()}")
            return False

    def _create_text_image_file(self, code: str, filename: str) -> bool:
        """바코드 생성 실패 시 텍스트 이미지 파일로 대체"""
        try:
            from PIL import Image, ImageDraw, ImageFont

            # 이미지 크기 계산
            width = max(200, len(code) * 12)
            height = 60

            # 이미지 생성
            img = Image.new("RGB", (width, height), color="white")
            draw = ImageDraw.Draw(img)

            # 기본 폰트 사용 (시스템 폰트 문제 회피)
            try:
                font = ImageFont.load_default()
            except:
                font = None

            # 텍스트 그리기
            text_x = 10
            text_y = 20
            draw.text((text_x, text_y), code, fill="black", font=font)

            # 간단한 바코드 모양 선 그리기
            bar_y = 5
            bar_height = 15
            bar_x = 10

            for i, char in enumerate(code):
                if i % 2 == 0:  # 짝수 위치에 선 그리기
                    draw.rectangle(
                        [bar_x + i * 8, bar_y, bar_x + i * 8 + 2, bar_y + bar_height],
                        fill="black",
                    )

            # 파일로 저장
            img.save(filename, format="PNG")

            if os.path.exists(filename) and os.path.getsize(filename) > 0:
                logger.info("BarcodeGenerator", f"텍스트 이미지 파일 생성 완료: {code}")
                return True
            else:
                logger.error(
                    "BarcodeGenerator", f"텍스트 이미지 파일 생성 실패: {filename}"
                )
                return False

        except Exception as e:
            logger.error(
                "BarcodeGenerator", f"텍스트 이미지 파일 생성 실패: {code} - {e}"
            )
            return False

    def _create_text_image(self, code: str) -> BytesIO:
        """바코드 생성 실패 시 텍스트 이미지로 대체"""
        try:
            # 간단한 텍스트 이미지 생성
            from PIL import Image, ImageDraw, ImageFont

            # 이미지 크기 계산
            width = max(200, len(code) * 12)
            height = 60

            # 이미지 생성
            img = Image.new("RGB", (width, height), color="white")
            draw = ImageDraw.Draw(img)

            # 기본 폰트 사용 (시스템 폰트 문제 회피)
            try:
                font = ImageFont.load_default()
            except:
                font = None

            # 텍스트 그리기
            text_x = 10
            text_y = 20
            draw.text((text_x, text_y), code, fill="black", font=font)

            # 간단한 바코드 모양 선 그리기
            bar_y = 5
            bar_height = 15
            bar_x = 10

            for i, char in enumerate(code):
                if i % 2 == 0:  # 짝수 위치에 선 그리기
                    draw.rectangle(
                        [bar_x + i * 8, bar_y, bar_x + i * 8 + 2, bar_y + bar_height],
                        fill="black",
                    )

            # BytesIO로 변환
            img_buffer = BytesIO()
            img.save(img_buffer, format="PNG")
            img_buffer.seek(0)

            logger.info("BarcodeGenerator", f"텍스트 이미지 생성 완료: {code}")
            return img_buffer

        except Exception as e:
            logger.error("BarcodeGenerator", f"텍스트 이미지 생성 실패: {code} - {e}")
            return None

    def generate_barcodes_for_products(
        self, products: List[Tuple[str, str, str, str]]
    ) -> dict:
        """상품 목록에 대한 바코드 생성 (메모리에 저장)"""
        barcode_images = {}

        for name, price, category, code in products:
            if code not in barcode_images:
                img_buffer = self.generate_barcode_in_memory(code, category)
                if img_buffer:
                    barcode_images[code] = img_buffer

        return barcode_images


class BarcodeFileGenerator:
    """바코드 이미지 파일 생성 클래스"""

    def __init__(self, output_dir: str = "barcodes"):
        self.output_dir = self._get_output_path(output_dir)
        self._ensure_output_dir()

        # 바코드 설정 (고화질)
        self.writer_options = {
            "dpi": 300,
            "module_width": 0.35,  # 더 넓은 바코드를 위해 증가
            "module_height": 15.0,
            "quiet_zone": 6.5,
            "text_distance": 5.0,
            "font_size": 10,
        }

    def _get_output_path(self, path: str) -> str:
        """실행 파일(exe) 또는 스크립트 위치에 따른 경로 반환"""
        if getattr(sys, "frozen", False):
            # PyInstaller로 패키징된 경우 (exe)
            base_path = os.path.dirname(sys.executable)
            print(f"PyInstaller 환경 감지: {base_path}")
        else:
            # 일반 스크립트 실행
            base_path = os.path.abspath(".")
            print(f"일반 스크립트 환경: {base_path}")

        output_path = os.path.join(base_path, path)
        logger.debug("BarcodeFileGenerator", f"바코드 출력 경로: {output_path}")
        return output_path

    def _ensure_output_dir(self):
        """출력 디렉토리 생성"""
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

    def generate_barcode(self, code: str, category: str) -> Optional[str]:
        """바코드 이미지 파일 생성"""
        filename = os.path.join(self.output_dir, f"{code}.png")
        logger.debug("BarcodeFileGenerator", f"바코드 파일 생성 시도: {filename}")

        if os.path.exists(filename):
            logger.debug(
                "BarcodeFileGenerator", f"바코드 파일이 이미 존재함: {filename}"
            )
            return filename

        try:
            logger.debug("BarcodeFileGenerator", "바코드 라이브러리 임포트 확인...")
            # 바코드 데이터 검증
            if not code or not isinstance(code, str):
                logger.error("BarcodeFileGenerator", f"잘못된 바코드 데이터: {code}")
                return None

            # ASCII 문자 확인
            try:
                code.encode("ascii")
            except UnicodeEncodeError:
                logger.warning(
                    "BarcodeFileGenerator", f"ASCII가 아닌 문자 포함: {code}"
                )
                code = "".join(c for c in code if ord(c) < 128)
                logger.info("BarcodeFileGenerator", f"ASCII 문자만 추출: {code}")

            if not code:
                logger.error("BarcodeFileGenerator", "유효한 바코드 데이터가 없음")
                return None

            logger.debug("BarcodeFileGenerator", f"Code128 바코드 생성 중: {code}")
            # GS1-128 바코드 생성
            writer = ImageWriter(format="PNG")
            barcode = Code128(code, writer=writer)

            logger.debug("BarcodeFileGenerator", f"바코드 파일 저장 중: {filename}")
            # 파일로 저장
            barcode.write(filename, self.writer_options)

            # 파일 생성 확인
            if os.path.exists(filename):
                file_size = os.path.getsize(filename)
                logger.info(
                    "BarcodeFileGenerator",
                    f"바코드 생성 완료: {code} ({category}) -> {filename} ({file_size} bytes)",
                )
                return filename
            else:
                logger.error(
                    "BarcodeFileGenerator",
                    "바코드 파일 생성 실패: 파일이 생성되지 않음",
                )
                return None

        except ImportError as e:
            logger.error("BarcodeFileGenerator", f"바코드 라이브러리 임포트 실패: {e}")
            return None
        except Exception as e:
            logger.error("BarcodeFileGenerator", f"바코드 생성 실패: {code} - {e}")
            import traceback

            traceback.print_exc()
            return None

    def generate_barcodes_for_products(
        self, products: List[Tuple[str, str, str, str]]
    ) -> Dict[str, str]:
        """상품 목록에 대한 바코드 파일 생성"""
        generated_files = {}
        unique_codes = set()

        for _, _, category, code in products:
            if code not in unique_codes:
                unique_codes.add(code)
                filepath = self.generate_barcode(code, category)
                if filepath:
                    generated_files[code] = filepath

        return generated_files

    def cleanup(self):
        """바코드 파일 및 디렉토리 정리"""
        if os.path.exists(self.output_dir):
            import shutil

            shutil.rmtree(self.output_dir)
            print(f"바코드 디렉토리 정리 완료: {self.output_dir}")

    def get_barcode_path(self, code: str) -> str:
        """바코드 파일 경로 반환"""
        return os.path.join(self.output_dir, f"{code}.png")
