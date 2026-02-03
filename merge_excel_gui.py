#!/usr/bin/env python3
"""
엑셀 파일 시트 병합 프로그램 (GUI 버전)
- 드래그앤드롭 또는 파일 선택으로 엑셀 파일 추가
- 더블클릭으로 실행 가능
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
except ImportError:
    # openpyxl 설치 안내
    import subprocess
    result = messagebox.askyesno(
        "모듈 설치 필요",
        "openpyxl 모듈이 필요합니다.\n설치하시겠습니까?"
    )
    if result:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
        from openpyxl import Workbook, load_workbook
        from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    else:
        sys.exit(1)


# ============================================
# 기본 저장 경로 설정 (여기서 변경 가능)
# ============================================
DEFAULT_SAVE_PATH = r"D:\_Users\MyHome\Downloads"
DEFAULT_FILENAME = "merged_output.xlsx"
# ============================================


class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("엑셀 시트 병합 프로그램")
        self.root.geometry("600x500")
        self.root.resizable(True, True)

        self.file_list = []

        # 기본 저장 경로 설정
        self.save_path = DEFAULT_SAVE_PATH
        if not os.path.exists(self.save_path):
            self.save_path = os.path.expanduser("~/Downloads")

        self.setup_ui()

    def setup_ui(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 설명 라벨
        desc_label = ttk.Label(
            main_frame,
            text="엑셀 파일들을 추가하면 각 파일의 첫 번째 시트를 병합합니다.\n"
                 "파일명: #####_구간명.xlsx → 시트명: 구간명\n"
                 "O열(DIA/INCH) 합산이 요약 시트에 표시됩니다.",
            justify=tk.LEFT
        )
        desc_label.pack(pady=(0, 10))

        # 파일 목록 프레임
        list_frame = ttk.LabelFrame(main_frame, text="파일 목록", padding="5")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 리스트박스 + 스크롤바
        list_scroll = ttk.Scrollbar(list_frame)
        list_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox = tk.Listbox(
            list_frame,
            height=12,
            yscrollcommand=list_scroll.set,
            selectmode=tk.EXTENDED
        )
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        list_scroll.config(command=self.listbox.yview)

        # 버튼 프레임
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(btn_frame, text="파일 추가", command=self.add_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="선택 삭제", command=self.remove_selected).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="전체 삭제", command=self.clear_all).pack(side=tk.LEFT, padx=2)

        # 저장 경로 표시 프레임
        save_frame = ttk.LabelFrame(main_frame, text="저장 위치", padding="5")
        save_frame.pack(fill=tk.X, pady=(0, 10))

        self.save_path_var = tk.StringVar(value=self.save_path)
        save_path_label = ttk.Label(save_frame, textvariable=self.save_path_var, foreground="blue")
        save_path_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Button(save_frame, text="변경", command=self.change_save_path, width=6).pack(side=tk.RIGHT)

        # 실행 버튼
        self.merge_btn = ttk.Button(
            main_frame,
            text="병합 실행",
            command=self.merge_files,
            style='Accent.TButton'
        )
        self.merge_btn.pack(fill=tk.X, pady=(0, 10))

        # 상태 표시
        self.status_var = tk.StringVar(value="파일을 추가해주세요.")
        status_label = ttk.Label(main_frame, textvariable=self.status_var)
        status_label.pack()

        # 진행바
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.pack(fill=tk.X, pady=(5, 0))

    def add_files(self):
        files = filedialog.askopenfilenames(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        for f in files:
            if f not in self.file_list:
                self.file_list.append(f)
                self.listbox.insert(tk.END, os.path.basename(f))

        self.status_var.set(f"{len(self.file_list)}개 파일 준비됨")

    def remove_selected(self):
        selected = list(self.listbox.curselection())
        selected.reverse()  # 뒤에서부터 삭제
        for idx in selected:
            self.listbox.delete(idx)
            del self.file_list[idx]
        self.status_var.set(f"{len(self.file_list)}개 파일 준비됨")

    def clear_all(self):
        self.listbox.delete(0, tk.END)
        self.file_list.clear()
        self.status_var.set("파일을 추가해주세요.")

    def change_save_path(self):
        """저장 경로 변경"""
        new_path = filedialog.askdirectory(
            title="저장 폴더 선택",
            initialdir=self.save_path
        )
        if new_path:
            self.save_path = new_path
            self.save_path_var.set(new_path)

    def extract_section_name(self, filename):
        """파일명에서 구간명 추출"""
        base_name = os.path.splitext(os.path.basename(filename))[0]
        if '_' in base_name:
            return base_name.split('_', 1)[1]
        return base_name

    def is_valid_data_sheet(self, sheet):
        """유효한 데이터 시트인지 확인"""
        try:
            a1 = sheet['A1'].value
            b1 = sheet['B1'].value
            if a1 and '등록' in str(a1):
                return True
            if b1 and 'SBNM' in str(b1).upper():
                return True
            if a1 is not None or b1 is not None:
                return True
        except:
            pass
        return False

    def get_first_valid_sheet(self, workbook):
        """첫 번째 유효한 시트 반환"""
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            if self.is_valid_data_sheet(sheet):
                return sheet
        return workbook.worksheets[0]

    def calculate_o_column_sum(self, sheet):
        """O열 합산"""
        total = 0
        for row in range(2, sheet.max_row + 1):
            cell_value = sheet.cell(row=row, column=15).value
            if cell_value is not None:
                try:
                    total += float(str(cell_value).replace(',', ''))
                except (ValueError, TypeError):
                    pass
        return total

    def copy_sheet_data(self, source_sheet, target_sheet):
        """시트 데이터 복사"""
        for row_idx, row in enumerate(source_sheet.iter_rows(), 1):
            for col_idx, cell in enumerate(row, 1):
                target_cell = target_sheet.cell(row=row_idx, column=col_idx)
                target_cell.value = cell.value
                if cell.has_style:
                    try:
                        target_cell.font = Font(
                            name=cell.font.name,
                            size=cell.font.size,
                            bold=cell.font.bold,
                            italic=cell.font.italic
                        )
                        target_cell.alignment = Alignment(
                            horizontal=cell.alignment.horizontal,
                            vertical=cell.alignment.vertical,
                            wrap_text=cell.alignment.wrap_text
                        )
                    except:
                        pass

        for col_idx, col_dim in source_sheet.column_dimensions.items():
            if col_dim.width:
                target_sheet.column_dimensions[col_idx].width = col_dim.width

    def create_summary_sheet(self, workbook, section_sums):
        """요약 시트 생성"""
        summary = workbook.create_sheet("요약", 0)

        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, size=12, color="FFFFFF")
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        center_align = Alignment(horizontal='center', vertical='center')

        summary['A1'] = "구간별 DIA/INCH 합산 요약"
        summary['A1'].font = Font(bold=True, size=14)
        summary.merge_cells('A1:C1')

        summary['A2'] = f"생성일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        summary.merge_cells('A2:C2')

        headers = ['No.', '구간명', 'DIA/INCH 합계']
        for col, header in enumerate(headers, 1):
            cell = summary.cell(row=4, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
            cell.alignment = center_align

        total_sum = 0
        for idx, (section, sum_value) in enumerate(section_sums.items(), 1):
            row = idx + 4
            summary.cell(row=row, column=1, value=idx).border = border
            summary.cell(row=row, column=2, value=section).border = border
            cell = summary.cell(row=row, column=3, value=sum_value)
            cell.border = border
            cell.number_format = '#,##0.00'
            total_sum += sum_value

        total_row = len(section_sums) + 5
        summary.cell(row=total_row, column=1, value="").border = border
        total_label = summary.cell(row=total_row, column=2, value="총 합계")
        total_label.font = Font(bold=True)
        total_label.border = border
        total_value = summary.cell(row=total_row, column=3, value=total_sum)
        total_value.font = Font(bold=True)
        total_value.border = border
        total_value.number_format = '#,##0.00'

        summary.column_dimensions['A'].width = 8
        summary.column_dimensions['B'].width = 30
        summary.column_dimensions['C'].width = 18

    def merge_files(self):
        if not self.file_list:
            messagebox.showwarning("경고", "파일을 먼저 추가해주세요.")
            return

        # 자동 저장 경로 생성 (타임스탬프로 중복 방지)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"merged_{timestamp}.xlsx"
        output_file = os.path.join(self.save_path, filename)

        # 폴더가 없으면 생성
        if not os.path.exists(self.save_path):
            os.makedirs(self.save_path)

        self.merge_btn.config(state='disabled')
        self.progress['value'] = 0
        self.progress['maximum'] = len(self.file_list)

        merged_wb = Workbook()
        merged_wb.remove(merged_wb.active)

        section_sums = {}

        for idx, file_path in enumerate(self.file_list):
            self.status_var.set(f"처리 중: {os.path.basename(file_path)}")
            self.root.update()

            try:
                source_wb = load_workbook(file_path, data_only=True)
                section_name = self.extract_section_name(file_path)

                # 중복 시트명 처리
                original_section = section_name
                counter = 1
                while section_name in [s.title for s in merged_wb.worksheets]:
                    section_name = f"{original_section}_{counter}"
                    counter += 1

                source_sheet = self.get_first_valid_sheet(source_wb)
                target_sheet = merged_wb.create_sheet(title=section_name[:31])
                self.copy_sheet_data(source_sheet, target_sheet)

                section_sums[section_name] = self.calculate_o_column_sum(source_sheet)
                source_wb.close()

            except Exception as e:
                messagebox.showerror("오류", f"파일 처리 중 오류:\n{file_path}\n{str(e)}")

            self.progress['value'] = idx + 1
            self.root.update()

        self.create_summary_sheet(merged_wb, section_sums)
        merged_wb.save(output_file)

        self.merge_btn.config(state='normal')
        self.status_var.set("완료!")

        messagebox.showinfo(
            "완료",
            f"병합이 완료되었습니다!\n\n"
            f"저장 위치: {output_file}\n"
            f"처리된 구간: {len(section_sums)}개"
        )

        # 파일 열기 여부
        if messagebox.askyesno("파일 열기", "병합된 파일을 열어볼까요?"):
            os.startfile(output_file) if os.name == 'nt' else os.system(f'open "{output_file}"')


def main():
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
