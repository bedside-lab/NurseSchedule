"""
간호사 근무표 자동생성 - 데스크톱 GUI

엑셀에 이미 신청된 근무를 최대한 반영하면서, 나머지 빈칸을 병동 규칙에 맞게
자동으로 채워주는 프로그램입니다. 매달 바뀔 수 있는 값(데이 전담 간호사,
근무조당 필요 인원 등)은 화면에서 직접 입력하고, 나머지 규칙은 지난번 합의된
기본값으로 채워져 있습니다 (필요하면 '고급 설정'에서 조정 가능).
"""

import json
import os
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from scheduler.core import run_pipeline

APP_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(APP_DIR, 'config_last.json')

DEFAULT_SETTINGS = {
    'input_file': '',
    'output_file': '',
    'n_history_days': 6,
    'n_target_days': 30,
    'dayonly_nurses': '박지현, 하다은',
    'staff_d': 5,
    'staff_e': 5,
    'staff_n': 4,
    'staff_d_min': 4,
    'staff_d_max': 6,
    'staff_e_min': 4,
    'staff_e_max': 6,
    'staff_n_min': 3,
    'staff_n_max': 5,
    'weekly_min_off': 2,
    'night_immediate_off_days': 2,
    'night_min_gap_days': 6,
    'night_min_gap_days_fallback': 5,
    'max_consec_workdays': 5,
    'n_range_max': 1,
    'e_range_max': 5,
    'solver_time_limit_sec': 280,
}


def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, encoding='utf-8') as f:
                saved = json.load(f)
            merged = dict(DEFAULT_SETTINGS)
            merged.update(saved)
            return merged
        except Exception:
            pass
    return dict(DEFAULT_SETTINGS)


def save_settings(values):
    with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(values, f, ensure_ascii=False, indent=2)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('간호사 근무표 자동생성')
        self.geometry('720x780')
        self.resizable(True, True)

        self.settings = load_settings()
        self.log_queue = queue.Queue()
        self.worker_thread = None

        self._build_ui()
        self.after(150, self._poll_log_queue)

    # ---------------------------------------------------------------- UI ---
    def _build_ui(self):
        pad = dict(padx=8, pady=4)
        row = 0

        file_frame = ttk.LabelFrame(self, text='1. 파일 선택')
        file_frame.pack(fill='x', **pad)

        self.input_var = tk.StringVar(value=self.settings['input_file'])
        self.output_var = tk.StringVar(value=self.settings['output_file'])

        r1 = ttk.Frame(file_frame); r1.pack(fill='x', **pad)
        ttk.Label(r1, text='신청 근무 엑셀 파일:', width=18).pack(side='left')
        ttk.Entry(r1, textvariable=self.input_var).pack(side='left', fill='x', expand=True)
        ttk.Button(r1, text='찾아보기', command=self._pick_input).pack(side='left', padx=4)

        r2 = ttk.Frame(file_frame); r2.pack(fill='x', **pad)
        ttk.Label(r2, text='결과 저장 위치:', width=18).pack(side='left')
        ttk.Entry(r2, textvariable=self.output_var).pack(side='left', fill='x', expand=True)
        ttk.Button(r2, text='찾아보기', command=self._pick_output).pack(side='left', padx=4)

        month_frame = ttk.LabelFrame(self, text='2. 이번 달 설정 (매달 확인)')
        month_frame.pack(fill='x', **pad)

        self.hist_days_var = tk.IntVar(value=self.settings['n_history_days'])
        self.target_days_var = tk.IntVar(value=self.settings['n_target_days'])
        self.dayonly_var = tk.StringVar(value=self.settings['dayonly_nurses'])
        self.staff_d_var = tk.IntVar(value=self.settings['staff_d'])
        self.staff_e_var = tk.IntVar(value=self.settings['staff_e'])
        self.staff_n_var = tk.IntVar(value=self.settings['staff_n'])

        r3 = ttk.Frame(month_frame); r3.pack(fill='x', **pad)
        ttk.Label(r3, text='이전 달 기록 일수:', width=18).pack(side='left')
        ttk.Spinbox(r3, from_=0, to=31, textvariable=self.hist_days_var, width=6).pack(side='left')
        ttk.Label(r3, text='  이번 달 일수:', width=14).pack(side='left')
        ttk.Spinbox(r3, from_=28, to=31, textvariable=self.target_days_var, width=6).pack(side='left')

        r4 = ttk.Frame(month_frame); r4.pack(fill='x', **pad)
        ttk.Label(r4, text='데이 전담 간호사:', width=18).pack(side='left')
        ttk.Entry(r4, textvariable=self.dayonly_var).pack(side='left', fill='x', expand=True)
        ttk.Label(r4, text='(쉼표로 구분)').pack(side='left', padx=4)

        r5 = ttk.Frame(month_frame); r5.pack(fill='x', **pad)
        ttk.Label(r5, text='근무조당 필요 인원:', width=18).pack(side='left')
        ttk.Label(r5, text='데이').pack(side='left')
        ttk.Spinbox(r5, from_=1, to=30, textvariable=self.staff_d_var, width=5).pack(side='left', padx=(2, 10))
        ttk.Label(r5, text='이브닝').pack(side='left')
        ttk.Spinbox(r5, from_=1, to=30, textvariable=self.staff_e_var, width=5).pack(side='left', padx=(2, 10))
        ttk.Label(r5, text='나이트').pack(side='left')
        ttk.Spinbox(r5, from_=1, to=30, textvariable=self.staff_n_var, width=5).pack(side='left', padx=2)

        # ---- advanced (collapsible) ----
        self.adv_visible = tk.BooleanVar(value=False)
        adv_toggle = ttk.Checkbutton(self, text='고급 설정 펼치기 (규칙 세부값)',
                                      variable=self.adv_visible, command=self._toggle_advanced)
        adv_toggle.pack(anchor='w', padx=12)

        self.adv_frame = ttk.LabelFrame(self, text='고급 설정 (평소엔 그대로 두어도 됩니다)')

        self.weekly_min_off_var = tk.IntVar(value=self.settings['weekly_min_off'])
        self.night_off_var = tk.IntVar(value=self.settings['night_immediate_off_days'])
        self.night_gap_var = tk.IntVar(value=self.settings['night_min_gap_days'])
        self.night_gap_fallback_var = tk.IntVar(value=self.settings['night_min_gap_days_fallback'])
        self.max_consec_var = tk.IntVar(value=self.settings['max_consec_workdays'])
        self.n_range_var = tk.IntVar(value=self.settings['n_range_max'])
        self.e_range_var = tk.IntVar(value=self.settings['e_range_max'])
        self.staff_d_min = tk.IntVar(value=self.settings['staff_d_min'])
        self.staff_d_max = tk.IntVar(value=self.settings['staff_d_max'])
        self.staff_e_min = tk.IntVar(value=self.settings['staff_e_min'])
        self.staff_e_max = tk.IntVar(value=self.settings['staff_e_max'])
        self.staff_n_min = tk.IntVar(value=self.settings['staff_n_min'])
        self.staff_n_max = tk.IntVar(value=self.settings['staff_n_max'])
        self.time_limit_var = tk.IntVar(value=self.settings['solver_time_limit_sec'])

        def adv_row(label, var, extra=''):
            r = ttk.Frame(self.adv_frame); r.pack(fill='x', padx=8, pady=3)
            ttk.Label(r, text=label, width=32).pack(side='left')
            ttk.Spinbox(r, from_=0, to=999, textvariable=var, width=8).pack(side='left')
            if extra:
                ttk.Label(r, text=extra).pack(side='left', padx=6)

        adv_row('주간 최소 오프 일수:', self.weekly_min_off_var)
        adv_row('나이트 직후 연속 오프:', self.night_off_var)
        adv_row('나이트 블록 간 최소 인터벌(우선):', self.night_gap_var, '일 - 이 값으로 먼저 시도')
        adv_row('나이트 블록 간 최소 인터벌(안되면):', self.night_gap_fallback_var,
                '일 - 위 값으로 안풀리면 이 값으로 재시도')
        adv_row('최대 연속 근무일수:', self.max_consec_var)
        adv_row('나이트 개수 허용 편차:', self.n_range_var)
        adv_row('이브닝 개수 허용 편차:', self.e_range_var)
        adv_row('데이 인원 허용 범위(최소):', self.staff_d_min)
        adv_row('데이 인원 허용 범위(최대):', self.staff_d_max)
        adv_row('이브닝 인원 허용 범위(최소):', self.staff_e_min)
        adv_row('이브닝 인원 허용 범위(최대):', self.staff_e_max)
        adv_row('나이트 인원 허용 범위(최소):', self.staff_n_min)
        adv_row('나이트 인원 허용 범위(최대):', self.staff_n_max)
        adv_row('계산 제한 시간(초):', self.time_limit_var, '길수록 더 좋은 결과, 느림')

        run_frame = ttk.Frame(self)
        run_frame.pack(fill='x', padx=8, pady=8)
        self.run_button = ttk.Button(run_frame, text='근무표 생성', command=self._on_run)
        self.run_button.pack(side='left')
        self.status_label = ttk.Label(run_frame, text='')
        self.status_label.pack(side='left', padx=10)

        log_frame = ttk.LabelFrame(self, text='진행 로그')
        log_frame.pack(fill='both', expand=True, padx=8, pady=4)
        self.log_box = scrolledtext.ScrolledText(log_frame, height=16, state='disabled')
        self.log_box.pack(fill='both', expand=True)

    def _toggle_advanced(self):
        if self.adv_visible.get():
            self.adv_frame.pack(fill='x', padx=8, pady=4)
        else:
            self.adv_frame.pack_forget()

    # ------------------------------------------------------------- pickers ---
    def _pick_input(self):
        path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx')])
        if path:
            self.input_var.set(path)
            if not self.output_var.get():
                base, ext = os.path.splitext(path)
                self.output_var.set(base + '_완성.xlsx')

    def _pick_output(self):
        path = filedialog.asksaveasfilename(defaultextension='.xlsx',
                                             filetypes=[('Excel files', '*.xlsx')])
        if path:
            self.output_var.set(path)

    # --------------------------------------------------------------- run ---
    def _collect_config(self):
        dayonly = [s.strip() for s in self.dayonly_var.get().split(',') if s.strip()]
        return dict(
            input_file=self.input_var.get(),
            output_file=self.output_var.get(),
            n_history_days=self.hist_days_var.get(),
            n_target_days=self.target_days_var.get(),
            dayonly_nurses=dayonly,
            staff_target={'D': self.staff_d_var.get(), 'E': self.staff_e_var.get(), 'N': self.staff_n_var.get()},
            staff_bounds={
                'D': (self.staff_d_min.get(), self.staff_d_max.get()),
                'E': (self.staff_e_min.get(), self.staff_e_max.get()),
                'N': (self.staff_n_min.get(), self.staff_n_max.get()),
            },
            weekly_min_off=self.weekly_min_off_var.get(),
            night_immediate_off_days=self.night_off_var.get(),
            night_min_gap_days=self.night_gap_var.get(),
            night_min_gap_days_fallback=self.night_gap_fallback_var.get(),
            max_consec_workdays=self.max_consec_var.get(),
            n_range_max=self.n_range_var.get(),
            e_range_max=self.e_range_var.get(),
            solver_time_limit_sec=self.time_limit_var.get(),
        )

    def _current_settings_dict(self):
        return dict(
            input_file=self.input_var.get(),
            output_file=self.output_var.get(),
            n_history_days=self.hist_days_var.get(),
            n_target_days=self.target_days_var.get(),
            dayonly_nurses=self.dayonly_var.get(),
            staff_d=self.staff_d_var.get(), staff_e=self.staff_e_var.get(), staff_n=self.staff_n_var.get(),
            staff_d_min=self.staff_d_min.get(), staff_d_max=self.staff_d_max.get(),
            staff_e_min=self.staff_e_min.get(), staff_e_max=self.staff_e_max.get(),
            staff_n_min=self.staff_n_min.get(), staff_n_max=self.staff_n_max.get(),
            weekly_min_off=self.weekly_min_off_var.get(),
            night_immediate_off_days=self.night_off_var.get(),
            night_min_gap_days=self.night_gap_var.get(),
            night_min_gap_days_fallback=self.night_gap_fallback_var.get(),
            max_consec_workdays=self.max_consec_var.get(),
            n_range_max=self.n_range_var.get(),
            e_range_max=self.e_range_var.get(),
            solver_time_limit_sec=self.time_limit_var.get(),
        )

    def _on_run(self):
        if not self.input_var.get() or not os.path.exists(self.input_var.get()):
            messagebox.showerror('오류', '엑셀 파일을 먼저 선택해주세요.')
            return
        if not self.output_var.get():
            messagebox.showerror('오류', '결과 저장 위치를 지정해주세요.')
            return

        save_settings(self._current_settings_dict())
        config = self._collect_config()

        self.run_button.config(state='disabled')
        self.status_label.config(text='생성 중...')
        self.log_box.config(state='normal')
        self.log_box.delete('1.0', tk.END)
        self.log_box.config(state='disabled')

        self.worker_thread = threading.Thread(target=self._run_worker, args=(config,), daemon=True)
        self.worker_thread.start()

    def _run_worker(self, config):
        def log(msg):
            self.log_queue.put(('log', str(msg)))
        try:
            solution, errors = run_pipeline(config, log=log)
            self.log_queue.put(('done', errors))
        except Exception as e:
            self.log_queue.put(('error', str(e)))

    def _poll_log_queue(self):
        try:
            while True:
                kind, payload = self.log_queue.get_nowait()
                if kind == 'log':
                    self.log_box.config(state='normal')
                    self.log_box.insert(tk.END, payload + '\n')
                    self.log_box.see(tk.END)
                    self.log_box.config(state='disabled')
                elif kind == 'done':
                    self.run_button.config(state='normal')
                    errors = payload
                    if errors:
                        self.status_label.config(text=f'완료 (검증 오류 {len(errors)}건, 로그 확인 필요)')
                        messagebox.showwarning('완료 (확인 필요)',
                                                f'근무표는 생성됐지만 검증에서 {len(errors)}건의 문제가 발견됐습니다.\n'
                                                '로그를 확인해주세요.')
                    else:
                        self.status_label.config(text='완료! (검증 통과)')
                        messagebox.showinfo('완료', '근무표 생성이 완료됐고, 모든 규칙 검증을 통과했습니다.')
                elif kind == 'error':
                    self.run_button.config(state='normal')
                    self.status_label.config(text='오류 발생')
                    messagebox.showerror('오류', f'근무표 생성 중 문제가 발생했습니다:\n{payload}')
        except queue.Empty:
            pass
        self.after(150, self._poll_log_queue)


if __name__ == '__main__':
    app = App()
    app.mainloop()
