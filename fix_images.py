#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
from PIL import Image
from pathlib import Path

def fix_png_images():
    # image 폴더 경로
    image_dir = Path("image")
    
    # 모든 PNG 파일 찾기
    png_files = list(image_dir.glob("*.png"))
    
    print(f"총 {len(png_files)}개의 PNG 파일을 찾았습니다.")
    
    for png_file in png_files:
        try:
            print(f"\n처리 중: {png_file.name}")
            
            # 이미지 열기
            with Image.open(png_file) as img:
                # 이미지 정보 출력
                print(f"  - 모드: {img.mode}")
                print(f"  - 크기: {img.size}")
                
                # RGBA 모드로 변환 (알파 채널 유지)
                if img.mode != 'RGBA':
                    img = img.convert('RGBA')
                
                # 새로운 이미지 생성 (이렇게 하면 icc_profile이 제거됨)
                new_img = Image.new('RGBA', img.size)
                new_img.paste(img, (0, 0))
                
                # 임시 파일로 저장
                temp_file = png_file.with_suffix('.temp.png')
                new_img.save(temp_file, 'PNG', optimize=True)
                
                # 원본 파일 삭제
                os.remove(png_file)
                
                # 임시 파일을 원래 이름으로 변경
                os.rename(temp_file, png_file)
                
                print(f"  ✓ 수정 완료")
                
        except Exception as e:
            print(f"  ! 오류 발생: {str(e)}")

if __name__ == "__main__":
    fix_png_images()
    print("\n모든 이미지 처리 완료") 