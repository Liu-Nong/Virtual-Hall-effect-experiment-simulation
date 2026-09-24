# Virtual-Hall-effect-experiment-simulation
霍尔效应虚拟仿真代码
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
霍尔效应虚拟仿真实验 - 2.3.0
- 磁场分布模型参数解耦：平台区半宽 x_flat 与过渡区宽度 w 独立可调
- 与 DH4512D 实测数据拟合（默认 x_flat=38mm, w=3.5mm）
- 保留 2.2.0 全部功能
"""

import tkinter as tk
from tkinter import ttk, messagebox
import numpy as np
import matplotlib
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
matplotlib.rcParams['axes.unicode_minus'] = False
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from mpl_toolkits.mplot3d import Axes3D
from mpl_toolkits.mplot3d.art3d import Poly3DCollection
import json
import csv
from datetime import datetime
import os

try:
    from mpl_toolkits.mplot3d import proj3d
    HAS_PROJ3D = True
except Exception:
    HAS_PROJ3D = False

try:
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


# ---------- 动态 Tooltip ----------
class DynamicToolTip:
    def __init__(self, widget, text_func):
        self.widget = widget; self.text_func = text_func; self.tip_window = None
        self.widget.bind('<Enter>', self.enter); self.widget.bind('<Leave>', self.leave)
    def enter(self, event=None):
        try: x, y, cx, cy = self.widget.bbox("insert")
        except Exception: x = y = 0
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        text = self.text_func()
        if not text: return
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True); tw.wm_geometry(f"+{x}+{y}")
        tk.Label(tw, text=text, justify=tk.LEFT, background="#ffffe0",
                 relief=tk.SOLID, borderwidth=1, font=("Arial", 10, "normal")).pack(ipadx=1)
    def leave(self, event=None):
        if self.tip_window: self.tip_window.destroy(); self.tip_window = None


def static_tip(widget, text):
    class StaticTip:
        def __init__(self, w, t):
            self.w = w; self.t = t; self.tip = None
            w.bind('<Enter>', self.enter); w.bind('<Leave>', self.leave)
        def enter(self, e):
            try: x, y, cx, cy = self.w.bbox("insert")
            except Exception: x = y = 0
            x += self.w.winfo_rootx() + 25
            y += self.w.winfo_rooty() + 20
            self.tip = tw = tk.Toplevel(self.w)
            tw.wm_overrideredirect(True); tw.wm_geometry(f"+{x}+{y}")
            tk.Label(tw, text=self.t, justify=tk.LEFT, background="#ffffe0",
                     relief=tk.SOLID, borderwidth=1, font=("Arial", 10, "normal")).pack(ipadx=1)
        def leave(self, e):
            if self.tip: self.tip.destroy(); self.tip = None
    StaticTip(widget, text)


# ---------- 主程序 ----------
class HallEffectSimulation:
    def __init__(self, root):
        self.root = root
        self.root.title("霍尔效应实验虚拟仿真 - 2.3.0")
        self.root.geometry("1600x900")
        self.root.grid_rowconfigure(0, weight=0)
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.configure('TLabel', font=('Arial', 12))
        style.configure('TLabelframe.Label', font=('Arial', 12, 'bold'))
        style.configure('TButton', font=('Arial', 11))
        style.configure('TCombobox', font=('Arial', 11))
        style.configure('TEntry', font=('Arial', 11))
        style.configure('TScale', font=('Arial', 10))
        style.configure('TNotebook.Tab', font=('Arial', 14, 'bold'))

        self.e = 1.602e-19

        self.init_parameters()
        self.measurements = []
        self.create_toolbar()
        self.create_main_panel()
        self.update_calculations()
        self.start_electron_animation()

    def init_parameters(self):
        self.semiconductor_type = tk.StringVar(value="P型")
        self.L = tk.DoubleVar(value=3.21)
        self.b = tk.DoubleVar(value=4.0)
        self.d = tk.DoubleVar(value=0.5)
        self.KH = tk.DoubleVar(value=43.0)
        self.sigma = tk.DoubleVar(value=21.95)
        self.C = tk.DoubleVar(value=3.903)
        self.I_m = tk.DoubleVar(value=0.3)
        self.B_direction = tk.StringVar(value="正")
        self.Is = tk.DoubleVar(value=2.0)
        self.Is_direction = tk.StringVar(value="负")

        self.alpha = tk.DoubleVar(value=0.05)
        self.beta = tk.DoubleVar(value=0.427)
        self.gamma = tk.DoubleVar(value=0.0854)

        self.B = tk.DoubleVar(value=0.0)
        self.VH1 = tk.DoubleVar(value=0.0)
        self.VH2 = tk.DoubleVar(value=0.0)
        self.VH3 = tk.DoubleVar(value=0.0)
        self.VH4 = tk.DoubleVar(value=0.0)
        self.VH_avg = tk.DoubleVar(value=0.0)
        self.RH = tk.DoubleVar(value=0.0)
        self.n = tk.DoubleVar(value=0.0)
        self.v = tk.DoubleVar(value=0.0)
        self.mu = tk.DoubleVar(value=0.0)

        self.experiment_type = tk.StringVar(value="自由探索")
        self.secondary_mode = tk.StringVar(value="纯理论")
        self.last_secondary_mode = "纯理论"

        self.power_on = False

        self.E_field = 0.0
        self.E_target = 0.0
        self.step_count = 0
        self.max_steps = 5
        self.accumulation_active = True
        self.target_vx = 0.0

        self.num_electrons = tk.IntVar(value=10)
        self.animation_running = True
        self.animation_timer = None
        self.animation_interval = 100
        self.performance_mode = tk.StringVar(value="标准")
        self.draw_trails = tk.BooleanVar(value=True)
        self.use_gradients = tk.BooleanVar(value=True)
        self.trail_max_len = 8

        self.trajectory_windows = {}

        # ---- 磁场分布测量（2.3.0 新参数体系）----
        self.probe_x = tk.DoubleVar(value=0.0)              # 探头位置，-60~60 mm
        self.probe_x_flat = tk.DoubleVar(value=38.0)        # 平台区半宽，默认 38 mm
        self.probe_w_transition = tk.DoubleVar(value=3.5)   # 过渡区特征宽度，默认 3.5 mm
        self.field_distribution_data = []

        self.dragging_probe = False

        self.electrons = []
        self.view_initialized = False
        self.init_electrons()

    # ---------- 磁场分布模型（2.3.0 解耦参数版）----------
    def field_profile(self, x, B0, x_flat, w_transition):
        """
        电磁铁气隙磁场水平分布模型（双曲正切，平台区与过渡区独立可调）

        参数：
          x            : 探头位置 (mm)，标量或数组
          B0           : 中心磁感应强度 (T)
          x_flat       : 平台区半宽 (mm)，B 保持 B0 的区间
          w_transition : 过渡区特征宽度 (mm)，控制下降陡度
        返回：
          与 x 同形状的 B 值 (T)

        性质：
          - 对称性：B(-x) = B(x)
          - 单调性：|x| 增大，B 单调递减
          - 峰值位置：x = 0 处 B ≈ B0（严格最大值）
          - 半高点：|x| = x_flat + 1.5·w_transition 处 B = B0/2
        """
        x_arr = np.asarray(x, dtype=float)
        if w_transition <= 0:
            w_transition = 1.0
        # 修正位移：让 |x|=x_flat 处 B≈0.95·B0（平台区边界），半高位于 x_flat+1.5w
        x_shift = x_flat + 1.5 * w_transition
        return B0 * 0.5 * (1.0 - np.tanh((np.abs(x_arr) - x_shift) / w_transition))

    def _field_display_range(self):
        """根据当前参数返回合适的显示范围（用于坐标轴和磁场线覆盖）"""
        x_flat = self.probe_x_flat.get()
        w = self.probe_w_transition.get()
        return max(x_flat + 4 * w + 10, 62)

    # ---------- 工具栏 ----------
    def create_toolbar(self):
        toolbar = ttk.Frame(self.root, padding="5", relief=tk.RAISED)
        toolbar.grid(row=0, column=0, sticky="ew")
        ttk.Label(toolbar, text="霍尔效应实验 2.3.0", font=("Arial", 14, "bold")).pack(side=tk.LEFT, padx=10)
        self.power_button = ttk.Button(toolbar, text="电源: OFF", command=self.toggle_power, width=12)
        self.power_button.pack(side=tk.LEFT, padx=5)
        self.pause_button = ttk.Button(toolbar, text="⏸️ 暂停", command=self.toggle_animation, width=10)
        self.pause_button.pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="🔄 重置电子", command=self.reset_electrons, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="🔄 重置视角", command=self.reset_view, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Label(toolbar, text="").pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(toolbar, text="提示：暂停后点击粒子可查看其完整轨迹俯视图",
                  font=("Arial", 10), foreground="gray").pack(side=tk.RIGHT, padx=10)

    # ---------- 主布局 ----------
    def create_main_panel(self):
        main_panel = ttk.Frame(self.root)
        main_panel.grid(row=1, column=0, sticky="nsew")
        main_panel.grid_rowconfigure(0, weight=1)
        main_panel.grid_columnconfigure(0, weight=1)

        self.paned = ttk.PanedWindow(main_panel, orient=tk.HORIZONTAL)
        self.paned.grid(row=0, column=0, sticky="nsew")

        self.left_frame = ttk.Frame(self.paned, padding="5")
        self.paned.add(self.left_frame, weight=1)
        self.left_frame.grid_rowconfigure(0, weight=1)
        self.left_frame.grid_columnconfigure(0, weight=1)

        self.notebook = ttk.Notebook(self.left_frame)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        self.param_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.param_frame, text="参数设置")
        self.build_param_panel()

        self.data_display_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.data_display_frame, text="数据显示")
        self.build_data_display_panel()

        self.field_dist_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.field_dist_frame, text="磁场分布测量")
        self.build_field_dist_panel()

        self.right_frame = ttk.Frame(self.paned, padding="5")
        self.paned.add(self.right_frame, weight=3)
        self.right_frame.grid_rowconfigure(0, weight=1)
        self.right_frame.grid_columnconfigure(0, weight=1)

        # 3D 画布
        self.canvas_3d_container = ttk.Frame(self.right_frame)
        self.canvas_3d_container.grid(row=0, column=0, sticky="nsew")
        self.fig = plt.figure(figsize=(10, 8), dpi=100)
        self.ax_3d = self.fig.add_subplot(111, projection='3d')
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.canvas_3d_container)
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        NavigationToolbar2Tk(self.canvas, self.canvas_3d_container)

        # 磁场分布画布
        self.canvas_field_container = ttk.Frame(self.right_frame)
        self.canvas_field_container.grid(row=0, column=0, sticky="nsew")
        self.canvas_field_container.grid_remove()

        self.field_fig = plt.figure(figsize=(10, 8), dpi=100)
        self.field_ax_top = self.field_fig.add_subplot(111)
        self.canvas_field = FigureCanvasTkAgg(self.field_fig, master=self.canvas_field_container)
        self.canvas_field.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        NavigationToolbar2Tk(self.canvas_field, self.canvas_field_container)

        self.canvas_field.mpl_connect('button_press_event', self.on_field_mouse_press)
        self.canvas_field.mpl_connect('button_release_event', self.on_field_mouse_release)
        self.canvas_field.mpl_connect('motion_notify_event', self.on_field_mouse_move)
        self.canvas_field.mpl_connect('scroll_event', self.on_field_scroll)

        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        self.canvas.mpl_connect('button_press_event', self.on_canvas_click)

    def on_tab_changed(self, event=None):
        try: current = self.notebook.index(self.notebook.select())
        except Exception: return
        if current == 2:
            self.canvas_3d_container.grid_remove()
            self.canvas_field_container.grid()
            self.update_field_distribution_plot()
        else:
            self.canvas_field_container.grid_remove()
            self.canvas_3d_container.grid()
            self.update_3d_plot()

    # ---------- 磁场分布右侧交互 ----------
    def on_field_mouse_press(self, event):
        if event.inaxes != self.field_ax_top: return
        if event.xdata is None: return
        if abs(event.xdata - self.probe_x.get()) < 5:
            self.dragging_probe = True

    def on_field_mouse_release(self, event):
        self.dragging_probe = False

    def on_field_mouse_move(self, event):
        if not self.dragging_probe: return
        if event.inaxes != self.field_ax_top or event.xdata is None: return
        x_range = self._field_display_range()
        new_x = float(np.clip(event.xdata, -x_range, x_range))
        new_x = float(f"{new_x:.6f}")
        self.probe_x.set(new_x)
        self.update_calculations()

    def on_field_scroll(self, event):
        """滚轮调整平台区半宽 x_flat（1 mm / 格）"""
        if event.inaxes != self.field_ax_top: return
        old = self.probe_x_flat.get()
        if event.button == 'up': new = min(old + 1, 55)
        elif event.button == 'down': new = max(old - 1, 5)
        else: return
        self.probe_x_flat.set(new)
        self.update_calculations()

    # ---------- 参数设置标签页 ----------
    def build_param_panel(self):
        self.slider_frames = []
        param_frame = self.param_frame
        param_frame.grid_rowconfigure(0, weight=1)
        param_frame.grid_columnconfigure(0, weight=1)

        canvas = tk.Canvas(param_frame)
        v_scrollbar = ttk.Scrollbar(param_frame, orient="vertical", command=canvas.yview)
        h_scrollbar = ttk.Scrollbar(param_frame, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        scrollable = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=scrollable, anchor="nw")

        def on_configure(event): canvas.configure(scrollregion=canvas.bbox("all"))
        scrollable.bind("<Configure>", on_configure)

        canvas.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")

        def on_mousewheel(event): canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        def on_shift_mousewheel(event): canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<MouseWheel>", on_mousewheel)
        scrollable.bind("<MouseWheel>", on_mousewheel)
        canvas.bind("<Shift-MouseWheel>", on_shift_mousewheel)
        scrollable.bind("<Shift-MouseWheel>", on_shift_mousewheel)

        # 半导体类型
        f0 = ttk.LabelFrame(scrollable, text="半导体类型", padding="5")
        f0.pack(fill="x", pady=5)
        frame = ttk.Frame(f0); frame.pack(fill="x")
        ttk.Label(frame, text="类型：").pack(side=tk.LEFT)
        sem_combo = ttk.Combobox(frame, textvariable=self.semiconductor_type,
                                 values=["P型", "N型"], state="readonly", width=10)
        sem_combo.pack(side=tk.LEFT, padx=5)
        sem_combo.bind("<<ComboboxSelected>>", self.on_semiconductor_changed)
        self.carrier_label = ttk.Label(frame, text="(载流子: 空穴)", foreground="blue")
        self.carrier_label.pack(side=tk.LEFT, padx=10)
        def update_carrier_label(*args):
            sem_type = self.semiconductor_type.get()
            if sem_type == "P型":
                self.carrier_label.config(text="(载流子: 空穴)", foreground="blue")
            else:
                self.carrier_label.config(text="(载流子: 电子)", foreground="red")
        self.semiconductor_type.trace('w', update_carrier_label)

        # 元件参数
        f1 = ttk.LabelFrame(scrollable, text="元件参数", padding="5")
        f1.pack(fill="x", pady=5)
        self.add_slider(f1, "L (mm)", self.L, 0, 15, "mm", tooltip="改变长度", resolution=0.1)
        self.add_slider(f1, "b (mm)", self.b, 0, 12, "mm", tooltip="改变宽度", resolution=0.1)
        self.add_slider(f1, "d (mm)", self.d, 0, 1.5, "mm", tooltip="改变高度", resolution=0.1)
        self.add_slider(f1, "KH", self.KH, 0, 500, "V/(A·T)", tooltip="灵敏度系数 K_H = 1/(n·e·d)")
        self.add_slider(f1, "σ", self.sigma, 0, 50, "A/(m·V)", tooltip="电导率，可计算 μ = σ·|R_H|")

        # 磁场参数
        f2 = ttk.LabelFrame(scrollable, text="磁场参数", padding="5")
        f2.pack(fill="x", pady=5)
        self.add_slider(f2, "C (kG/sA)", self.C, 0, 12, "kG/sA", tooltip="转换系数，B = C · I_m / 10")
        self.im_label = tk.Label(f2, text="Im (A)", font=("Arial", 12))
        self.im_label.pack(anchor="w", pady=(5,0))
        self.add_slider(f2, "", self.I_m, 0.0, 1.0, "A",
                        callback=lambda *args: self.reset_accumulation(),
                        tooltip="流过电磁铁的电流", resolution=0.1)
        frame = ttk.Frame(f2); frame.pack(fill="x", pady=2)
        ttk.Label(frame, text="方向：").pack(side=tk.LEFT)
        B_dir_combo = ttk.Combobox(frame, textvariable=self.B_direction,
                                   values=["正", "负"], state="readonly", width=8)
        B_dir_combo.pack(side=tk.LEFT, padx=5)
        B_dir_combo.bind("<<ComboboxSelected>>", lambda e: (self.reset_accumulation(), self.update_calculations()))
        frame = ttk.Frame(f2); frame.pack(fill="x", pady=2)
        ttk.Label(frame, text="当前 B：").pack(side=tk.LEFT)
        B_entry = ttk.Entry(frame, textvariable=self.B, width=8, state='readonly')
        B_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(frame, text="T").pack(side=tk.LEFT)

        # 工作电流
        f3 = ttk.LabelFrame(scrollable, text="工作电流", padding="5")
        f3.pack(fill="x", pady=5)
        self.is_label = tk.Label(f3, text="Is (mA)", font=("Arial", 12))
        self.is_label.pack(anchor="w", pady=(5,0))
        self.add_slider(f3, "", self.Is, 0, 10, "mA", resolution=0.1)
        frame = ttk.Frame(f3); frame.pack(fill="x", pady=2)
        ttk.Label(frame, text="方向：").pack(side=tk.LEFT)
        Is_dir_combo = ttk.Combobox(frame, textvariable=self.Is_direction,
                                    values=["正", "负"], state="readonly", width=8)
        Is_dir_combo.pack(side=tk.LEFT, padx=5)
        Is_dir_combo.bind("<<ComboboxSelected>>", lambda e: (self.reset_accumulation(), self.update_calculations()))

        # 实验模式
        f4 = ttk.LabelFrame(scrollable, text="实验模式", padding="5")
        f4.pack(fill="x", pady=5)
        left_frame = ttk.Frame(f4); left_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Label(left_frame, text="类型：").pack(anchor="w")
        type_combo = ttk.Combobox(left_frame, textvariable=self.experiment_type,
                                  values=["自由探索", "UH-Is曲线测绘", "UH-Im曲线测绘"],
                                  state="readonly", width=18)
        type_combo.pack(fill=tk.X, pady=2)
        type_combo.bind("<<ComboboxSelected>>", self.on_mode_change)

        right_frame = ttk.Frame(f4); right_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=5)
        ttk.Label(right_frame, text="副效应：").pack(anchor="w")
        self.secondary_combo = ttk.Combobox(right_frame, textvariable=self.secondary_mode,
                                            values=["纯理论", "仿真"], state="readonly", width=18)
        self.secondary_combo.pack(fill=tk.X, pady=2)
        self.secondary_combo.bind("<<ComboboxSelected>>", self.on_mode_change)
        def get_secondary_tip():
            return "没有副效应模拟" if self.secondary_mode.get() == "纯理论" else \
                   ("添加副效应模拟：\nV₀ = α·I_s·sign(I_s)\nVₜ = β·B·sign(B)\n"
                    "Vₚ = γ·I_s·B·sign(I_s)·sign(B)")
        DynamicToolTip(self.secondary_combo, get_secondary_tip)

        # 样品特性参数
        f5 = ttk.LabelFrame(scrollable, text="样品特性参数", padding="5")
        f5.pack(fill="x", pady=5)
        self.add_slider(f5, "α (mV/mA)", self.alpha, 0, 0.2, "mV/mA",
                        tooltip="取决于霍尔元件电极的几何对称性和材料电阻率均匀性。")
        self.add_slider(f5, "β (mV/T)", self.beta, 0, 1.0, "mV/T",
                        tooltip="取决于材料的导热性、载流子迁移率以及温度梯度。")
        self.add_slider(f5, "γ (mV/(mA·T))", self.gamma, 0, 0.2, "mV/(mA·T)",
                        tooltip="反映高阶耦合效应，通常非常小。")

        # 粒子数量与动画设置
        anim_frame = ttk.LabelFrame(scrollable, text="粒子数量与动画设置", padding="5")
        anim_frame.pack(fill="x", pady=5)
        f = ttk.Frame(anim_frame); f.pack(fill="x", pady=2)
        ttk.Label(f, text="粒子数量：").pack(side=tk.LEFT)
        entry = ttk.Entry(f, textvariable=self.num_electrons, width=5)
        entry.pack(side=tk.LEFT, padx=5)
        entry.bind("<Return>", lambda e: self.update_particle_count())
        entry.bind("<FocusOut>", lambda e: self.update_particle_count())
        slider = ttk.Scale(f, from_=1, to=25, variable=self.num_electrons,
                           orient='horizontal', length=150, command=lambda v: self.update_particle_count())
        slider.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Label(f, text="(1~25)").pack(side=tk.LEFT, padx=5)

        f = ttk.Frame(anim_frame); f.pack(fill="x", pady=2)
        ttk.Label(f, text="性能模式：").pack(side=tk.LEFT)
        perf_combo = ttk.Combobox(f, textvariable=self.performance_mode,
                                  values=["标准", "省电"], state="readonly", width=10)
        perf_combo.pack(side=tk.LEFT, padx=5)
        perf_combo.bind("<<ComboboxSelected>>", self.on_performance_mode_change)
        ttk.Checkbutton(anim_frame, text="显示载流子轨迹", variable=self.draw_trails,
                        command=self.toggle_trails).pack(anchor="w", pady=2)
        ttk.Checkbutton(anim_frame, text="使用渐变效果", variable=self.use_gradients,
                        command=self.toggle_gradients).pack(anchor="w", pady=2)

    def add_slider(self, parent, label, var, minv, maxv, unit, callback=None, tooltip=None, resolution=None):
        f = ttk.Frame(parent)
        f.pack(fill="x", pady=2)
        if label:
            lbl = ttk.Label(f, text=f"{label}：")
            lbl.pack(side=tk.LEFT)
            if tooltip: static_tip(lbl, tooltip)
        entry = ttk.Entry(f, textvariable=var, width=8)
        entry.pack(side=tk.LEFT, padx=5)
        entry.bind("<Return>", lambda e: self.validate_param(var, minv, maxv, callback))
        entry.bind("<FocusOut>", lambda e: self.validate_param(var, minv, maxv, callback))

        def on_slider_move(v):
            if resolution is not None:
                try:
                    snapped = round(float(v) / resolution) * resolution
                    snapped = float(f"{snapped:.6f}")
                    snapped = max(minv, min(maxv, snapped))
                    if abs(snapped - var.get()) > 1e-9:
                        var.set(snapped)
                except Exception:
                    pass
            self.update_calculations()

        slider = ttk.Scale(f, from_=minv, to=maxv, variable=var,
                           orient='horizontal', length=150, command=on_slider_move)
        slider.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Label(f, text=unit).pack(side=tk.LEFT, padx=5)
        if callback: var.trace('w', callback)
        self.slider_frames.append((f, var, slider, entry))

    def update_secondary_sliders_state(self):
        enabled = (self.secondary_mode.get() == "仿真")
        for frame, var, slider, entry in self.slider_frames:
            if var in (self.alpha, self.beta, self.gamma):
                state = tk.NORMAL if enabled else tk.DISABLED
                slider.config(state=state); entry.config(state=state)
                if not enabled: var.set(0.0)

    # ---------- 数据显示标签页 ----------
    def build_data_display_panel(self):
        display = self.data_display_frame
        display.grid_rowconfigure(0, weight=1); display.grid_columnconfigure(0, weight=1)

        canvas = tk.Canvas(display)
        v_scrollbar = ttk.Scrollbar(display, orient="vertical", command=canvas.yview)
        h_scrollbar = ttk.Scrollbar(display, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        scrollable = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=scrollable, anchor="nw")
        def on_configure(event): canvas.configure(scrollregion=canvas.bbox("all"))
        scrollable.bind("<Configure>", on_configure)
        canvas.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns"); h_scrollbar.grid(row=1, column=0, sticky="ew")
        def on_mousewheel(event): canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<MouseWheel>", on_mousewheel); scrollable.bind("<MouseWheel>", on_mousewheel)

        v_frame = ttk.LabelFrame(scrollable, text="霍尔电压测量结果", padding="5")
        v_frame.pack(fill="x", pady=5)
        self.VH1_label = ttk.Label(v_frame, text="U1 = 0.0000 mV"); self.VH1_label.pack(anchor="w")
        DynamicToolTip(self.VH1_label, self.get_vh1_tip)
        self.VH2_label = ttk.Label(v_frame, text="U2 = 0.0000 mV"); self.VH2_label.pack(anchor="w")
        DynamicToolTip(self.VH2_label, self.get_vh2_tip)
        self.VH3_label = ttk.Label(v_frame, text="U3 = 0.0000 mV"); self.VH3_label.pack(anchor="w")
        DynamicToolTip(self.VH3_label, self.get_vh3_tip)
        self.VH4_label = ttk.Label(v_frame, text="U4 = 0.0000 mV"); self.VH4_label.pack(anchor="w")
        DynamicToolTip(self.VH4_label, self.get_vh4_tip)
        self.VH_avg_label = ttk.Label(v_frame, text="UH = 0.0000 mV",
                                      font=("Arial", 12, "bold"), foreground="blue")
        self.VH_avg_label.pack(anchor="w")

        p_frame = ttk.LabelFrame(scrollable, text="数据处理结果", padding="5")
        p_frame.pack(fill="x", pady=5)
        self.RH_label = ttk.Label(p_frame, text="RH = 0.0000 m³/C"); self.RH_label.pack(anchor="w")
        self.n_label = ttk.Label(p_frame, text="n = 0.0000 m⁻³"); self.n_label.pack(anchor="w")
        self.v_label = ttk.Label(p_frame, text="v = 0.0000 m/s"); self.v_label.pack(anchor="w")
        self.mu_label = ttk.Label(p_frame, text="μ = 0.0000 cm²/(V·s)"); self.mu_label.pack(anchor="w")

        table_frame = ttk.LabelFrame(scrollable, text="实验数据表格", padding="5")
        table_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        columns = ('序号', 'Is/mA', 'Im/A', 'B/T', 'U1/mV', 'U2/mV', 'U3/mV', 'U4/mV', 'UH/mV')
        self.data_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=10)
        column_widths = [40, 60, 60, 80, 80, 80, 80, 80, 80]
        for i, col in enumerate(columns):
            self.data_tree.heading(col, text=col)
            self.data_tree.column(col, width=column_widths[i], minwidth=50)
        scroll_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.data_tree.yview)
        scroll_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.data_tree.xview)
        self.data_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.data_tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns"); scroll_x.grid(row=1, column=0, sticky="ew")
        table_frame.grid_rowconfigure(0, weight=1); table_frame.grid_columnconfigure(0, weight=1)

        btn_frame = ttk.Frame(scrollable); btn_frame.pack(fill="x", pady=5)
        ttk.Button(btn_frame, text="📊 记录当前数据点", command=self.record_data_point, width=18).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="🗑️ 清除表格", command=self.clear_experiment_data, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="📈 绘制曲线", command=self.open_curve_window, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="💾 导出JSON", command=lambda: self.export_data("json"), width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="💾 导出CSV", command=lambda: self.export_data("csv"), width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="🖨️ 报告", command=self.print_data_report, width=10).pack(side=tk.LEFT, padx=2)

    # ---------- 磁场分布测量标签页 ----------
    def build_field_dist_panel(self):
        frame = self.field_dist_frame
        frame.grid_rowconfigure(0, weight=1); frame.grid_columnconfigure(0, weight=1)

        canvas = tk.Canvas(frame)
        v_scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        h_scrollbar = ttk.Scrollbar(frame, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        scrollable = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=scrollable, anchor="nw")
        def on_configure(event): canvas.configure(scrollregion=canvas.bbox("all"))
        scrollable.bind("<Configure>", on_configure)
        canvas.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns"); h_scrollbar.grid(row=1, column=0, sticky="ew")
        def on_mousewheel(event): canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<MouseWheel>", on_mousewheel); scrollable.bind("<MouseWheel>", on_mousewheel)

        tip = ttk.Label(scrollable,
                        text="实验说明：固定 Is 与 Im，沿水平方向移动霍尔探头，测量不同位置的 B(x)。\n"
                             "提示：右侧图中可直接拖拽红色探头；鼠标在图内滚轮可调节平台区半宽。",
                        font=("Arial", 11), foreground="gray", justify=tk.LEFT)
        tip.pack(anchor="w", pady=5)

        # 探头与磁场参数（2.3.0 新参数体系）
        f1 = ttk.LabelFrame(scrollable, text="探头与磁场分布参数", padding="5")
        f1.pack(fill="x", pady=5)
        self.add_slider(f1, "探头位置 x (mm)", self.probe_x, -60, 60, "mm",
                        tooltip="霍尔探头沿水平方向的位置。x=0 为气隙中心，"
                                "此处 B 取严格最大值。",
                        resolution=0.1)
        self.add_slider(f1, "平台区半宽 x_flat (mm)", self.probe_x_flat, 5, 55, "mm",
                        tooltip="B 保持最大值 B₀ 的区间半宽。\n"
                                "默认 38 mm，对应 DH4512D 实测电磁铁。\n"
                                "中心 x=0 处 B 恒为最大值，此参数只改变平台长度。",
                        resolution=1.0)
        self.add_slider(f1, "过渡区宽度 w (mm)", self.probe_w_transition, 1, 15, "mm",
                        tooltip="磁场从平台区衰减到零的特征宽度。\n"
                                "值越小，边缘下降越陡；值越大，过渡越平缓。\n"
                                "默认 3.5 mm。",
                        resolution=0.5)

        f2 = ttk.LabelFrame(scrollable, text="当前测量值", padding="5")
        f2.pack(fill="x", pady=5)
        self.probe_B_label = ttk.Label(f2, text="B(x) = 0.0000 T",
                                       font=("Courier New", 13, "bold"), foreground="purple")
        self.probe_B_label.pack(anchor="w", pady=2)
        self.probe_pos_label = ttk.Label(f2, text="位置 x = 0.0000 mm", font=("Courier New", 11))
        self.probe_pos_label.pack(anchor="w", pady=2)

        f3 = ttk.LabelFrame(scrollable, text="操作", padding="5")
        f3.pack(fill="x", pady=5)
        btn1 = ttk.Frame(f3); btn1.pack(fill="x", pady=2)
        ttk.Button(btn1, text="📊 记录当前点", command=self.record_field_point, width=16).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn1, text="🗑️ 清除数据", command=self.clear_field_data, width=12).pack(side=tk.LEFT, padx=2)
        btn2 = ttk.Frame(f3); btn2.pack(fill="x", pady=2)
        ttk.Button(btn2, text="📈 生成分布曲线", command=self.open_field_curve_window, width=16).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn2, text="💾 导出CSV", command=self.export_field_data, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn2, text="🖨️ 生成报告", command=self.print_field_report, width=12).pack(side=tk.LEFT, padx=2)

        f4 = ttk.LabelFrame(scrollable, text="测量数据", padding="5")
        f4.pack(fill=tk.BOTH, expand=True, pady=5)
        cols = ('序号', 'x/mm', 'B/T')
        self.field_tree = ttk.Treeview(f4, columns=cols, show='headings', height=10)
        widths = [60, 120, 150]
        for i, c in enumerate(cols):
            self.field_tree.heading(c, text=c); self.field_tree.column(c, width=widths[i], minwidth=50)
        sy = ttk.Scrollbar(f4, orient="vertical", command=self.field_tree.yview)
        self.field_tree.configure(yscrollcommand=sy.set)
        self.field_tree.grid(row=0, column=0, sticky="nsew"); sy.grid(row=0, column=1, sticky="ns")
        f4.grid_rowconfigure(0, weight=1); f4.grid_columnconfigure(0, weight=1)

    # ---------- 磁场分布右侧视图 ----------
    def update_field_distribution_plot(self):
        x_flat = self.probe_x_flat.get()
        w = self.probe_w_transition.get()
        x_probe = self.probe_x.get()
        B_center = (self.C.get() * self.I_m.get()) / 10.0
        x_disp = self._field_display_range()

        ax1 = self.field_ax_top
        ax1.clear()
        ax1.set_title("电磁铁气隙磁场分布示意图  （拖拽红色探头移动；滚轮调节平台区半宽）",
                      fontsize=12, fontweight='bold')

        # 气隙区域（横跨整个显示范围）
        ax1.fill_between([-x_disp, x_disp], 0, 3,
                         color='lightyellow', alpha=0.3, label='气隙区域')
        # 平台区高亮
        ax1.axvspan(-x_flat, x_flat, color='lightgreen', alpha=0.35, label=f'平台区 ±{x_flat:.1f}mm')
        # 上下磁极
        ax1.add_patch(plt.Rectangle((-x_disp, 3), 2*x_disp, 0.8, color='gray', alpha=0.7))
        ax1.add_patch(plt.Rectangle((-x_disp, -0.8), 2*x_disp, 0.8, color='gray', alpha=0.7))
        ax1.text(0, 3.4, 'N 极', ha='center', fontsize=11, fontweight='bold')
        ax1.text(0, -0.6, 'S 极', ha='center', fontsize=11, fontweight='bold')

        # 磁场线（密度反映 B 强度）
        n_lines = 51
        for i in range(n_lines):
            x_pos = -x_disp + (2 * x_disp) * i / (n_lines - 1)
            B_local = float(self.field_profile(x_pos, B_center, x_flat, w))
            alpha = max(0.05, B_local / B_center) if B_center > 0 else 0.05
            ax1.annotate('', xy=(x_pos, 0), xytext=(x_pos, 3),
                         arrowprops=dict(arrowstyle='->', color='red',
                                         alpha=alpha, linewidth=1.8))

        # 探头位置
        ax1.plot(x_probe, 1.5, 'rv', markersize=18,
                 label=f'探头 x={x_probe:.4f} mm', zorder=10)
        ax1.axvline(x=x_probe, color='red', linestyle='--', alpha=0.5)

        B_at_probe = float(self.field_profile(x_probe, B_center, x_flat, w))
        ax1.text(x_probe, 2.6, f"B = {B_at_probe*1000:.4f} mT",
                 color='red', fontsize=10, ha='center', fontweight='bold',
                 bbox=dict(boxstyle='round', facecolor='white', alpha=0.8))

        ax1.set_xlim(-x_disp, x_disp)
        ax1.set_ylim(-1.2, 4.2)
        ax1.set_xlabel('水平位置 x (mm)')
        ax1.set_yticks([])
        ax1.legend(loc='upper right', fontsize=9)
        ax1.grid(True, axis='x', alpha=0.2)

        self.field_fig.tight_layout()
        self.canvas_field.draw_idle()

    # ---------- 磁场分布曲线窗口 ----------
    def open_field_curve_window(self):
        if not self.field_distribution_data:
            messagebox.showwarning("警告", "没有测量数据，请先记录数据点！"); return

        win = tk.Toplevel(self.root)
        win.title("电磁铁磁场水平分布曲线")
        win.geometry("900x650")

        fig, ax = plt.subplots(figsize=(8.5, 6))
        B_center = (self.C.get() * self.I_m.get()) / 10.0
        x_flat = self.probe_x_flat.get()
        w = self.probe_w_transition.get()
        x_disp = self._field_display_range()

        x_fit = np.linspace(-x_disp, x_disp, 600)
        B_fit = self.field_profile(x_fit, B_center, x_flat, w)
        ax.plot(x_fit, B_fit * 1000, 'r--', linewidth=2,
                label=f'理论分布 B(x)（中心 B₀={B_center*1000:.4f} mT）')

        # 平台区边界辅助线
        ax.axvline(x=x_flat, color='orange', linestyle=':', alpha=0.7,
                   label=f'平台区边界 ±{x_flat:.1f} mm')
        ax.axvline(x=-x_flat, color='orange', linestyle=':', alpha=0.7)

        # 半高位置标注
        x_half = x_flat + 1.5 * w
        ax.axvline(x=x_half, color='gray', linestyle='--', alpha=0.5,
                   label=f'半高位置 ±{x_half:.1f} mm')
        ax.axvline(x=-x_half, color='gray', linestyle='--', alpha=0.5)

        # 测量点
        xs = [p['x/mm'] for p in self.field_distribution_data]
        Bs = [p['B/T'] * 1000 for p in self.field_distribution_data]
        pairs = sorted(zip(xs, Bs))
        xs_sorted = [p[0] for p in pairs]; Bs_sorted = [p[1] for p in pairs]
        ax.scatter(xs_sorted, Bs_sorted, color='purple', s=70, zorder=5,
                   edgecolors='darkviolet', linewidths=1.5, label='测量点')

        ax.set_xlabel('探头位置 x (mm)')
        ax.set_ylabel('磁感应强度 B (mT)')
        ax.set_title('磁场沿水平方向分布 B(x)', fontsize=13, fontweight='bold')
        ax.grid(True, alpha=0.3)
        ax.legend(loc='upper right', fontsize=9)
        fig.tight_layout()

        canvas_curve = FigureCanvasTkAgg(fig, master=win)
        canvas_curve.draw()
        canvas_curve.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        toolbar = NavigationToolbar2Tk(canvas_curve, win)
        toolbar.update()

        def on_scroll(event):
            if event.inaxes != ax: return
            if event.xdata is None or event.ydata is None: return
            scale_factor = 1.15 if event.button == 'up' else 1 / 1.15
            xlim = ax.get_xlim(); ylim = ax.get_ylim()
            new_w = (xlim[1] - xlim[0]) / scale_factor
            new_h = (ylim[1] - ylim[0]) / scale_factor
            rx = (event.xdata - xlim[0]) / (xlim[1] - xlim[0])
            ry = (event.ydata - ylim[0]) / (ylim[1] - ylim[0])
            ax.set_xlim([event.xdata - new_w * rx, event.xdata + new_w * (1 - rx)])
            ax.set_ylim([event.ydata - new_h * ry, event.ydata + new_h * (1 - ry)])
            canvas_curve.draw_idle()

        drag_state = {'x': 0, 'y': 0, 'xlim': None, 'ylim': None}
        def on_press(event):
            if event.inaxes != ax: return
            if event.button == 1:
                drag_state['x'] = event.x; drag_state['y'] = event.y
                drag_state['xlim'] = ax.get_xlim()
                drag_state['ylim'] = ax.get_ylim()
                win.config(cursor="fleur")
        def on_move(event):
            if drag_state['xlim'] is None: return
            if event.x is None or event.y is None: return
            dx = event.x - drag_state['x']; dy = event.y - drag_state['y']
            bbox = ax.get_window_extent()
            x_per_pix = (drag_state['xlim'][1] - drag_state['xlim'][0]) / bbox.width
            y_per_pix = (drag_state['ylim'][1] - drag_state['ylim'][0]) / bbox.height
            ax.set_xlim([drag_state['xlim'][0] - dx * x_per_pix,
                         drag_state['xlim'][1] - dx * x_per_pix])
            ax.set_ylim([drag_state['ylim'][0] - dy * y_per_pix,
                         drag_state['ylim'][1] - dy * y_per_pix])
            canvas_curve.draw_idle()
        def on_release(event):
            drag_state['xlim'] = None; drag_state['ylim'] = None
            win.config(cursor="")

        canvas_curve.mpl_connect('scroll_event', on_scroll)
        canvas_curve.mpl_connect('button_press_event', on_press)
        canvas_curve.mpl_connect('motion_notify_event', on_move)
        canvas_curve.mpl_connect('button_release_event', on_release)

    # ---------- 磁场分布数据管理 ----------
    def record_field_point(self):
        if not self.power_on:
            messagebox.showwarning("提示", "请先打开电源！"); return
        if self.field_distribution_data:
            last = self.field_distribution_data[-1]
            if abs(last['x/mm'] - self.probe_x.get()) < 0.0001:
                messagebox.showwarning("提示", "探头位置没有变化！\n请调整位置后再记录。"); return
        point = {'序号': len(self.field_distribution_data) + 1,
                 'x/mm': self.probe_x.get(), 'B/T': self.B.get()}
        self.field_distribution_data.append(point)
        self.update_field_table()
        self.update_field_distribution_plot()

    def update_field_table(self):
        for item in self.field_tree.get_children(): self.field_tree.delete(item)
        for p in self.field_distribution_data:
            self.field_tree.insert('', 'end', values=(
                p['序号'], f"{p['x/mm']:.4f}", f"{p['B/T']:.4f}"))

    def clear_field_data(self):
        self.field_distribution_data = []
        self.update_field_table()
        if self.notebook.index(self.notebook.select()) == 2:
            self.update_field_distribution_plot()

    def export_field_data(self):
        if not self.field_distribution_data:
            messagebox.showwarning("警告", "没有数据可导出"); return
        desktop = self.get_desktop_path()
        filename = os.path.join(desktop, f"磁场分布_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
        try:
            with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.DictWriter(f, fieldnames=['序号', 'x/mm', 'B/T'])
                writer.writeheader()
                for p in self.field_distribution_data:
                    writer.writerow({
                        '序号': p['序号'],
                        'x/mm': f"{p['x/mm']:.4f}",
                        'B/T': f"{p['B/T']:.4f}"
                    })
            messagebox.showinfo("成功", f"数据已导出到桌面:\n{filename}")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {str(e)}")

    def print_field_report(self):
        if not self.field_distribution_data:
            messagebox.showwarning("警告", "没有数据可生成报告"); return
        if not HAS_DOCX:
            messagebox.showerror("缺少依赖", "请先安装 python-docx：\n\npip install python-docx"); return
        desktop = self.get_desktop_path()
        filename = os.path.join(desktop, f"磁场分布实验报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx")
        try:
            doc = Document()
            title = doc.add_heading('电磁铁磁场水平分布测量报告', level=0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_heading('一、实验条件', level=1)
            t = doc.add_table(rows=6, cols=2); t.style = 'Light Grid Accent 1'
            t.cell(0,0).text = '工作电流 Is/mA'; t.cell(0,1).text = f'{self.Is.get():.4f}'
            t.cell(1,0).text = '励磁电流 Im/A'; t.cell(1,1).text = f'{self.I_m.get():.4f}'
            t.cell(2,0).text = '磁铁规格 C (kG/sA)'; t.cell(2,1).text = f'{self.C.get():.4f}'
            t.cell(3,0).text = '平台区半宽 x_flat/mm'; t.cell(3,1).text = f'{self.probe_x_flat.get():.4f}'
            t.cell(4,0).text = '过渡区宽度 w/mm'; t.cell(4,1).text = f'{self.probe_w_transition.get():.4f}'
            t.cell(5,0).text = '中心磁场 B0/T'; t.cell(5,1).text = f'{(self.C.get()*self.I_m.get())/10:.4f}'
            doc.add_heading('二、测量数据', level=1)
            cols = ['序号', 'x/mm', 'B/T']
            dt = doc.add_table(rows=1, cols=len(cols)); dt.style = 'Light Grid Accent 1'
            for i, c in enumerate(cols): dt.rows[0].cells[i].text = c
            for p in self.field_distribution_data:
                row = dt.add_row().cells
                row[0].text = str(p['序号'])
                row[1].text = f"{p['x/mm']:.4f}"; row[2].text = f"{p['B/T']:.4f}"
            doc.add_heading('三、结论', level=1)
            doc.add_paragraph(
                f'测量结果表明：电磁铁气隙磁场在中心 |x| ≤ {self.probe_x_flat.get():.1f} mm '
                f'区域存在明显均匀平台区（B ≈ B₀，最大），过渡区特征宽度约 '
                f'{self.probe_w_transition.get():.1f} mm，磁场在过渡区快速衰减到零。'
                f'该分布符合双曲正切模型 '
                f'B(x) = B₀/2·[1-tanh((|x|-x_flat-1.5w)/w)]，'
                f'与 DH4512D 实测数据吻合良好，x=0 处取严格最大值。')
            doc.save(filename)
            messagebox.showinfo("成功", f"报告已生成到桌面:\n{filename}")
        except Exception as e:
            messagebox.showerror("错误", f"生成失败: {str(e)}")

    # ---------- Tooltip ----------
    def get_vh1_tip(self):
        return "虚仿：U1 = 主电压(+I,+B) + V₀ + Vₜ + Vₚ" if self.secondary_mode.get() == "仿真" else "纯理论：无副效应叠加"
    def get_vh2_tip(self):
        return "虚仿：U2 = 主电压(+I,-B) + V₀ - Vₜ - Vₚ" if self.secondary_mode.get() == "仿真" else "纯理论：无副效应叠加"
    def get_vh3_tip(self):
        return "虚仿：U3 = 主电压(-I,-B) - V₀ - Vₜ + Vₚ" if self.secondary_mode.get() == "仿真" else "纯理论：无副效应叠加"
    def get_vh4_tip(self):
        return "虚仿：U4 = 主电压(-I,+B) - V₀ + Vₜ - Vₚ" if self.secondary_mode.get() == "仿真" else "纯理论：无副效应叠加"

    # ---------- 粒子数量 ----------
    def update_particle_count(self):
        n = self.num_electrons.get()
        if n < 1: self.num_electrons.set(1); n = 1
        elif n > 25: self.num_electrons.set(25); n = 25
        self.init_electrons()
        if self.power_on: self.update_target_velocity_if_needed()
        self.update_3d_plot()

    def update_target_velocity_if_needed(self):
        if not self.power_on: return
        Is = self.Is.get(); sem_type = self.semiconductor_type.get()
        charge_sign = 1 if sem_type == "P型" else -1
        Is_dir_sign = 1 if self.Is_direction.get() == "正" else -1
        v0 = 0.8 * (Is / 2.0)
        new_target = v0 * Is_dir_sign * charge_sign
        if abs(new_target - self.target_vx) > 1e-6:
            self.target_vx = new_target
            for e in self.electrons: e['vel'][0] = self.target_vx

    # ---------- 核心控制 ----------
    def toggle_power(self):
        self.power_on = not self.power_on
        self.power_button.config(text=f"电源: {'ON' if self.power_on else 'OFF'}")
        if self.power_on:
            self.E_field = 0.0; self.step_count = 0; self.accumulation_active = True
            self.init_electrons()
            Is = self.Is.get(); sem_type = self.semiconductor_type.get()
            charge_sign = 1 if sem_type == "P型" else -1
            Is_dir_sign = 1 if self.Is_direction.get() == "正" else -1
            v0 = 0.8 * (Is / 2.0)
            self.target_vx = v0 * Is_dir_sign * charge_sign
            for e in self.electrons: e['vel'][0] = self.target_vx
            self.update_calculations()
        else:
            self.init_electrons()
            for e in self.electrons: e['vel'] = np.array([0.0, 0.0, 0.0])
            self.update_calculations()

    def toggle_animation(self):
        self.animation_running = not self.animation_running
        self.pause_button.config(text="⏸️ 暂停" if self.animation_running else "▶️ 继续")
        if self.animation_running: self.animation_loop()

    def reset_electrons(self):
        self.init_electrons()
        self.E_field = 0.0; self.step_count = 0; self.accumulation_active = True
        if self.power_on:
            self.update_target_velocity_if_needed()
            B = self.B.get(); B_dir_sign = 1 if self.B_direction.get() == "正" else -1
            self.E_target = self.target_vx * B * B_dir_sign * 8.0
        self.update_3d_plot()

    def reset_view(self):
        self.ax_3d.view_init(elev=20, azim=-60)
        self.canvas.draw_idle()

    def reset_accumulation(self):
        self.E_field = 0.0; self.step_count = 0; self.accumulation_active = True
        if self.power_on:
            B = self.B.get(); B_dir_sign = 1 if self.B_direction.get() == "正" else -1
            self.E_target = self.target_vx * B * B_dir_sign * 8.0

    def validate_param(self, var, minv, maxv, callback=None):
        try:
            v = var.get()
            v = float(f"{v:.6f}")
            if v < minv or v > maxv:
                var.set(max(minv, min(v, maxv)))
            else:
                var.set(v)
            self.update_calculations()
            if callback: callback()
        except:
            pass

    def on_mode_change(self, event=None):
        new_secondary = self.secondary_mode.get()
        if new_secondary != self.last_secondary_mode:
            if new_secondary == "仿真" and self.last_secondary_mode == "纯理论":
                self.alpha.set(0.05); self.beta.set(0.427); self.gamma.set(0.0854)
            self.last_secondary_mode = new_secondary
        self.reset_accumulation()
        self.measurements = []; self.update_data_table()
        self.update_calculations(); self.update_mark_labels()
        self.update_secondary_sliders_state()

    def update_mark_labels(self):
        mode = self.experiment_type.get()
        self.is_label.config(foreground="black", font=("Arial", 12))
        self.im_label.config(foreground="black", font=("Arial", 12))
        if mode == "UH-IS曲线测绘":
            self.is_label.config(foreground="red", font=("Arial", 12, "underline"))
        elif mode == "UH-IM曲线测绘":
            self.im_label.config(foreground="red", font=("Arial", 12, "underline"))

    def on_semiconductor_changed(self, event=None):
        self.update_calculations(); self.init_electrons()
        if self.power_on:
            self.E_field = 0.0; self.step_count = 0; self.accumulation_active = True
            Is = self.Is.get(); sem_type = self.semiconductor_type.get()
            charge_sign = 1 if sem_type == "P型" else -1
            Is_dir_sign = 1 if self.Is_direction.get() == "正" else -1
            v0 = 0.8 * (Is / 2.0)
            self.target_vx = v0 * Is_dir_sign * charge_sign
            for e in self.electrons: e['vel'][0] = self.target_vx
            B = self.B.get(); B_dir_sign = 1 if self.B_direction.get() == "正" else -1
            self.E_target = self.target_vx * B * B_dir_sign * 8.0
        self.update_3d_plot()

    def on_performance_mode_change(self, event=None):
        mode = self.performance_mode.get()
        self.animation_interval = 100 if mode == "标准" else 150

    def toggle_trails(self): pass
    def toggle_gradients(self): pass

    # ---------- 计算 ----------
    def update_calculations(self):
        if not self.power_on:
            self.B.set(0.0)
            self.VH1.set(0.0); self.VH2.set(0.0); self.VH3.set(0.0)
            self.VH4.set(0.0); self.VH_avg.set(0.0)
            self.RH.set(0.0); self.n.set(0.0); self.v.set(0.0); self.mu.set(0.0)
            self.update_labels(); self.update_3d_plot(); return

        L = self.L.get() * 1e-3; b = self.b.get() * 1e-3; d = self.d.get() * 1e-3
        KH = self.KH.get(); sigma = self.sigma.get()
        Im = self.I_m.get(); C = self.C.get()
        B_center = (C * Im) / 10.0

        # 磁场分布计算（2.3.0 新参数）
        x_flat = self.probe_x_flat.get()
        w = self.probe_w_transition.get()
        x_probe = self.probe_x.get()
        B = float(self.field_profile(x_probe, B_center, x_flat, w))
        self.B.set(B)

        if hasattr(self, 'probe_B_label'):
            self.probe_B_label.config(text=f"B(x) = {B:.4f} T")
            self.probe_pos_label.config(text=f"位置 x = {x_probe:.4f} mm")

        Is = self.Is.get() * 1e-3; Is_mA = self.Is.get(); B_T = B
        B_dir_sign = 1 if self.B_direction.get() == "正" else -1
        Is_dir_sign = 1 if self.Is_direction.get() == "正" else -1

        Is_raw = self.Is.get(); sem_type = self.semiconductor_type.get()
        charge_sign = 1 if sem_type == "P型" else -1
        v0 = 0.8 * (Is_raw / 2.0)
        self.target_vx = v0 * Is_dir_sign * charge_sign

        V_main_pos_pos = KH * (Is_dir_sign * Is * 1000) * (B_dir_sign * B)
        V_main_pos_neg = KH * (Is_dir_sign * Is * 1000) * (-B_dir_sign * B)
        V_main_neg_neg = KH * (-Is_dir_sign * Is * 1000) * (-B_dir_sign * B)
        V_main_neg_pos = KH * (-Is_dir_sign * Is * 1000) * (B_dir_sign * B)

        use_secondary = (self.secondary_mode.get() == "仿真")
        if use_secondary:
            alpha = self.alpha.get(); beta = self.beta.get(); gamma = self.gamma.get()
            U0 = alpha * Is_mA * Is_dir_sign
            UT = beta * B_T * B_dir_sign
            UR = gamma * Is_mA * B_T * Is_dir_sign * B_dir_sign
            VH1 = V_main_pos_pos + U0 + UT + UR
            VH2 = V_main_pos_neg + U0 - UT - UR
            VH3 = V_main_neg_neg - U0 - UT + UR
            VH4 = V_main_neg_pos - U0 + UT - UR
        else:
            VH1 = V_main_pos_pos; VH2 = V_main_pos_neg
            VH3 = V_main_neg_neg; VH4 = V_main_neg_pos

        VH_avg = (abs(VH1) + abs(VH2) + abs(VH3) + abs(VH4)) / 4

        sem_type = self.semiconductor_type.get()
        sign = 1 if sem_type == "P型" else -1
        RH = sign * KH * d
        n = 1 / (abs(RH) * self.e) if abs(RH) > 1e-25 else 0
        if abs(n * self.e * b * d) > 1e-25:
            vx_base = (Is_dir_sign * Is) / (n * self.e * b * d)
            charge_sign = 1 if sem_type == "P型" else -1
            vx = vx_base * charge_sign
        else: vx = 0
        mu = sigma / (n * self.e) / 1e4 if abs(n * self.e) > 1e-25 else 0

        self.VH1.set(VH1); self.VH2.set(VH2); self.VH3.set(VH3); self.VH4.set(VH4)
        self.VH_avg.set(VH_avg); self.RH.set(RH); self.n.set(n)
        self.v.set(vx); self.mu.set(mu)
        self.update_labels()

        if self.power_on:
            B_dir_sign = 1 if self.B_direction.get() == "正" else -1
            self.E_target = self.target_vx * B * B_dir_sign * 8.0
            if self.step_count >= self.max_steps: self.E_field = self.E_target
            elif self.accumulation_active: self.E_field = self.E_target * (self.step_count / self.max_steps)

        self.update_target_velocity_if_needed()

        try: current = self.notebook.index(self.notebook.select())
        except Exception: current = 0
        if current == 2: self.update_field_distribution_plot()
        else: self.update_3d_plot()

    def update_labels(self):
        self.VH1_label.config(text=f"U1 = {self.VH1.get():.4f} mV")
        self.VH2_label.config(text=f"U2 = {self.VH2.get():.4f} mV")
        self.VH3_label.config(text=f"U3 = {self.VH3.get():.4f} mV")
        self.VH4_label.config(text=f"U4 = {self.VH4.get():.4f} mV")
        self.VH_avg_label.config(text=f"UH = {self.VH_avg.get():.4f} mV")
        self.RH_label.config(text=f"RH = {self.RH.get():.4f} m³/C")
        n_val = self.n.get()
        self.n_label.config(text=f"n = {n_val:.4e} m⁻³" if n_val > 1e10 else "n = 0.0000 m⁻³")
        vx = self.v.get()
        if abs(vx) > 1e-6:
            dir_text = "→ (向右)" if vx > 0 else "← (向左)"
            self.v_label.config(text=f"v = {vx:+.4f} m/s {dir_text}")
        else: self.v_label.config(text=f"v = {vx:+.4f} m/s 静止")
        self.mu_label.config(text=f"μ = {self.mu.get():.4f} cm²/(V·s)")

    # ---------- 粒子轨迹俯视图 ----------
    def project_3d_to_canvas(self, x, y, z):
        if not HAS_PROJ3D: return None
        try:
            x2, y2, _ = proj3d.proj_transform(x, y, z, self.ax_3d.get_proj())
            return self.ax_3d.transData.transform((x2, y2))
        except Exception: return None

    def on_canvas_click(self, event):
        if self.animation_running: return
        if event.inaxes != self.ax_3d: return
        best_dist = float('inf'); best_electron = None
        for e in self.electrons:
            px, py, pz = e['pos']
            disp = self.project_3d_to_canvas(px, py, pz)
            if disp is None: continue
            dist = np.hypot(disp[0] - event.x, disp[1] - event.y)
            if dist < best_dist: best_dist = dist; best_electron = e
        if best_electron is not None and best_dist < 30:
            self.show_full_trajectory_topview(best_electron)

    def show_full_trajectory_topview(self, electron):
        trail = electron.get('full_trail', [])
        if len(trail) < 2:
            messagebox.showinfo("提示", "该粒子运动轨迹过短，无法显示"); return
        eid = id(electron)
        if eid in self.trajectory_windows:
            try: self.trajectory_windows[eid].destroy()
            except Exception: pass
            del self.trajectory_windows[eid]
        win = tk.Toplevel(self.root)
        win.title("粒子运动轨迹（俯视图）"); win.geometry("720x580")
        self.trajectory_windows[eid] = win
        def on_close(): self.trajectory_windows.pop(eid, None); win.destroy()
        win.protocol("WM_DELETE_WINDOW", on_close)
        fig = plt.Figure(figsize=(6.6, 5.0), dpi=100); ax = fig.add_subplot(111)
        trail_arr = np.array(trail); L = self.L.get(); b = self.b.get()
        ax.add_patch(plt.Rectangle((0, 0), L, b, fill=False, edgecolor='black', linewidth=2, label='半导体'))
        n = len(trail_arr)
        for i in range(n - 1):
            alpha = 0.15 + 0.85 * (i / max(n - 1, 1))
            ax.plot(trail_arr[i:i+2, 0], trail_arr[i:i+2, 1], color='#1f77b4', alpha=alpha, linewidth=1.5)
        ax.plot(trail_arr[0, 0], trail_arr[0, 1], 'go', markersize=10, label='起始点', zorder=5)
        ax.plot(trail_arr[-1, 0], trail_arr[-1, 1], 'ro', markersize=10, label='终点', zorder=5)
        carrier_type = "P型" if self.semiconductor_type.get() == "P型" else "N型"
        ax.set_title(f'粒子运动轨迹\n当前粒子（{carrier_type}）')
        ax.set_xlim(-0.3, L + 0.3); ax.set_ylim(-0.3, b + 0.3)
        ax.set_xlabel('X (mm)'); ax.set_ylabel('Y (mm)')
        ax.set_aspect('equal'); ax.grid(True, alpha=0.3); ax.legend(loc='upper right')
        canvas = FigureCanvasTkAgg(fig, master=win); canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        btn_frame = ttk.Frame(win); btn_frame.pack(fill=tk.X, pady=5)
        ttk.Button(btn_frame, text="关闭", command=on_close, width=12).pack()

    # ---------- 3D绘图 ----------
    def update_3d_plot(self):
        self.ax_3d.clear()
        L = self.L.get(); b = self.b.get(); d = self.d.get()
        vertices = np.array([
            [0, 0, 0], [L, 0, 0], [L, b, 0], [0, b, 0],
            [0, 0, d], [L, 0, d], [L, b, d], [0, b, d]
        ])
        faces_indices = [[0,1,2,3],[4,5,6,7],[0,1,5,4],[2,3,7,6],[1,2,6,5],[4,7,3,0]]
        face_colors = ['lightblue','lightblue','lightgreen','lightgreen','lightgreen','lightgreen']
        face_names = ["2'面","2面","3'面","3面","1面","1'面"]
        for fi, (face, color, name) in enumerate(zip(faces_indices, face_colors, face_names)):
            pts = vertices[face]
            poly = Poly3DCollection([pts], alpha=0.3, facecolors=color, edgecolors='black', linewidths=2)
            self.ax_3d.add_collection3d(poly)
            center = np.mean(pts, axis=0)
            if fi == 0: center[2] -= 0.1
            elif fi == 1: center[2] += 0.1
            elif fi == 2: center[1] -= 0.1
            elif fi == 3: center[1] += 0.1
            elif fi == 4: center[0] += 0.1
            elif fi == 5: center[0] -= 0.1
            self.ax_3d.text(center[0], center[1], center[2], name, fontsize=10,
                            color='darkblue', fontweight='bold', ha='center')

        edges = [[0,1],[1,2],[2,3],[3,0],[4,5],[5,6],[6,7],[7,4],[0,4],[1,5],[2,6],[3,7]]
        for e in edges:
            self.ax_3d.plot3D([vertices[e[0]][0], vertices[e[1]][0]],
                             [vertices[e[0]][1], vertices[e[1]][1]],
                             [vertices[e[0]][2], vertices[e[1]][2]], 'k-', linewidth=2)

        if self.power_on and self.Is.get() != 0:
            Is = self.Is.get(); Is_direction = self.Is_direction.get()
            Is_normalized = abs(Is) / 10.0
            arrow_length = L * 0.6 * (0.3 + 0.7 * Is_normalized)
            arrow_width = 2 + 3 * Is_normalized
            if Is_direction == "正": start_x = L * 0.2; direction = 1
            else: start_x = L * 0.8; direction = -1
            self.ax_3d.quiver(start_x, b/2, d/2, direction*arrow_length, 0, 0,
                              color='blue', arrow_length_ratio=0.15, linewidth=arrow_width, pivot='tail')
            self.ax_3d.text(L/2, b/2-0.8, d/2+0.5,
                            f"I_s = {Is} mA {'→' if direction==1 else '←'}",
                            fontsize=12, color='blue', fontweight='bold')

        B = self.B.get()
        if self.power_on and B > 0:
            if self.B_direction.get() == "正":
                self.ax_3d.quiver(L/2, b/2, d*1.2, 0, 0, d*0.5, color='red',
                                  arrow_length_ratio=0.3, linewidth=3, pivot='tail')
                self.ax_3d.text(L/2+0.5, b/2+0.5, d*1.5, f"B = {B:.4f} T",
                                fontsize=12, color='red', fontweight='bold')
            else:
                self.ax_3d.quiver(L/2, b/2, -d*0.2, 0, 0, -d*0.5, color='red',
                                  arrow_length_ratio=0.3, linewidth=3, pivot='tail')
                self.ax_3d.text(L/2+0.5, b/2+0.5, -d*0.7, f"B = {B:.4f} T",
                                fontsize=12, color='red', fontweight='bold')

        if self.power_on:
            status = f"E_field: {self.E_field:.4f} / {self.E_target:.4f} (步数 {self.step_count}/{self.max_steps})"
            self.ax_3d.text(L/2, -0.3, d+0.5, status, fontsize=10,
                            color='darkgreen', fontweight='bold')
        else:
            self.ax_3d.text(L/2, b/2-1.0, d+0.5, "电源关闭", fontsize=12,
                            color='gray', fontweight='bold')

        if self.power_on:
            y_mid = b/2
            n_top = sum(1 for e in self.electrons if e['pos'][1] > y_mid)
            n_bottom = sum(1 for e in self.electrons if e['pos'][1] < y_mid)
            color = 'blue' if self.semiconductor_type.get() == "P型" else 'red'
            self.ax_3d.text(L/2, b+0.5, d+0.5,
                            f"上侧积累: {n_top}\n下侧积累: {n_bottom}",
                            fontsize=10, color=color, fontweight='bold',
                            ha='center', va='top')

        self.ax_3d.set_xlabel('X (mm)'); self.ax_3d.set_ylabel('Y (mm)'); self.ax_3d.set_zlabel('Z (mm)')
        self.ax_3d.set_xlim([-L*0.3, L*1.3]); self.ax_3d.set_ylim([-b*0.3, b*1.3]); self.ax_3d.set_zlim([-d*0.3, d*1.5])
        if not self.view_initialized:
            self.ax_3d.view_init(elev=20, azim=-60); self.view_initialized = True
        self.ax_3d.grid(True)
        self.draw_electrons()
        self.canvas.draw_idle()

    # ---------- 粒子绘制 ----------
    def draw_electrons(self):
        if self.performance_mode.get() == "省电": self.draw_electrons_optimized(); return
        if self.use_gradients.get() and self.performance_mode.get() == "标准":
            self.draw_electrons_high_quality()
        else: self.draw_electrons_optimized()

    def draw_electrons_high_quality(self):
        positions = []; colors = []; sizes = []
        is_p = self.semiconductor_type.get() == "P型"
        for e in self.electrons:
            positions.append(e['pos'])
            speed = np.linalg.norm(e['vel']); norm_speed = min(speed / 0.05, 1.0)
            if is_p:
                r, g, b, a = 0.2, 0.4, 1.0, 0.85
                r += 0.5*norm_speed; g += 0.3*norm_speed; b -= 0.4*norm_speed
            else:
                r, g, b, a = 1.0, 0.2, 0.2, 0.85
                r -= 0.2*norm_speed; g += 0.4*norm_speed; b += 0.4*norm_speed
            colors.append((np.clip(r,0,1), np.clip(g,0,1), np.clip(b,0,1), a))
            sizes.append(40 + 60 * norm_speed)
            if self.draw_trails.get() and len(e['trail']) > 1:
                trail = np.array(e['trail']); step = max(1, len(trail)//8)
                for i in range(0, len(trail)-1, step):
                    if i+1 < len(trail):
                        alpha = (i+1)/len(trail)*0.5
                        color = (1,0.3,0.3,alpha) if not is_p else (0.3,0.5,1.0,alpha)
                        self.ax_3d.plot(trail[i:i+2,0], trail[i:i+2,1], trail[i:i+2,2],
                                        color=color, linewidth=1.5, alpha=alpha)
        positions = np.array(positions)
        self.ax_3d.scatter(positions[:,0], positions[:,1], positions[:,2],
                           c=colors, s=sizes, edgecolors='darkred' if not is_p else 'darkblue',
                           linewidths=0.5, marker='o')

    def draw_electrons_optimized(self):
        positions = []
        is_p = self.semiconductor_type.get() == "P型"
        color = 'blue' if is_p else 'red'
        for e in self.electrons:
            positions.append(e['pos'])
            if self.draw_trails.get() and len(e['trail']) > 1:
                trail = np.array(e['trail'])
                self.ax_3d.plot(trail[:,0], trail[:,1], trail[:,2],
                                color=color, alpha=0.4, linewidth=1.0)
        positions = np.array(positions)
        self.ax_3d.scatter(positions[:,0], positions[:,1], positions[:,2],
                           c=color, s=50, alpha=0.8,
                           edgecolors='darkblue' if is_p else 'darkred',
                           linewidths=0.5, marker='o')

    # ---------- 粒子物理 ----------
    def init_electrons(self):
        self.electrons = []
        n = self.num_electrons.get()
        L = self.L.get(); b = self.b.get(); d = self.d.get()
        grid_x = int(np.ceil(np.sqrt(n))); grid_y = int(np.ceil(n / grid_x))
        for i in range(n):
            ix = i % grid_x; iy = i // grid_x
            x = (ix + 0.5)*(L/grid_x) + np.random.uniform(-L/grid_x*0.2, L/grid_x*0.2)
            y = (iy + 0.5)*(b/grid_y) + np.random.uniform(-b/grid_y*0.2, b/grid_y*0.2)
            z = np.random.uniform(0.1*d, 0.9*d)
            x = np.clip(x, 0, L); y = np.clip(y, 0, b); z = np.clip(z, 0, d)
            self.electrons.append({
                'pos': np.array([x, y, z]), 'vel': np.array([0.0, 0.0, 0.0]),
                'trail': [], 'full_trail': []
            })

    def update_electrons_positions(self):
        if not self.power_on:
            for e in self.electrons:
                e['vel'] *= 0.9
                if np.linalg.norm(e['vel']) < 1e-6: e['vel'] = np.array([0.0, 0.0, 0.0])
            return
        L = self.L.get(); b = self.b.get(); d = self.d.get()
        B = self.B.get(); Is = self.Is.get()
        dt = 0.02
        B_dir_sign = 1 if self.B_direction.get() == "正" else -1
        sem_type = self.semiconductor_type.get()
        charge_sign = 1 if sem_type == "P型" else -1
        target_vx = self.target_vx

        to_remove_y = []
        for i, e in enumerate(self.electrons):
            if e['pos'][1] <= 0 or e['pos'][1] >= b:
                if self.accumulation_active:
                    to_remove_y.append(i)
                    if self.step_count < self.max_steps:
                        self.step_count += 1
                        self.E_field = self.E_target * (self.step_count / self.max_steps)
                        if self.step_count >= self.max_steps:
                            self.accumulation_active = False
                else:
                    e['pos'][1] = 0.01 if e['pos'][1] <= 0 else b - 0.01
                    e['vel'][1] *= -0.5

        radius = min(b, d) / 2.0
        for idx in sorted(to_remove_y, reverse=True):
            del self.electrons[idx]
            side_x = np.random.choice([0, L])
            angle = np.random.uniform(0, 2*np.pi)
            r = np.random.uniform(0, radius)
            y_new = np.clip(b/2 + r*np.cos(angle), 0, b)
            z_new = np.clip(d/2 + r*np.sin(angle), 0, d)
            self.electrons.append({
                'pos': np.array([side_x, y_new, z_new]),
                'vel': np.array([target_vx, 0.0, 0.0]),
                'trail': [], 'full_trail': []
            })

        to_remove_x = []
        for i, e in enumerate(self.electrons):
            if e['pos'][0] <= 0: to_remove_x.append((i, L))
            elif e['pos'][0] >= L: to_remove_x.append((i, 0))

        for idx, new_x in sorted(to_remove_x, key=lambda x: x[0], reverse=True):
            del self.electrons[idx]
            self.electrons.append({
                'pos': np.array([new_x, np.random.uniform(0, b), np.random.uniform(0, d)]),
                'vel': np.array([target_vx, 0.0, 0.0]),
                'trail': [], 'full_trail': []
            })

        for e in self.electrons:
            if self.draw_trails.get():
                if len(e['trail']) == 0:
                    e['trail'].append(e['pos'].copy())
                else:
                    last_pos = e['trail'][-1]
                    if np.linalg.norm(e['pos'] - last_pos) > L * 0.5: e['trail'] = []
                    else:
                        if np.linalg.norm(e['pos'] - last_pos) > 0.01:
                            e['trail'].append(e['pos'].copy())
                            if len(e['trail']) > self.trail_max_len: e['trail'].pop(0)
            if len(e['full_trail']) == 0:
                e['full_trail'].append(e['pos'].copy())
            else:
                if np.linalg.norm(e['pos'] - e['full_trail'][-1]) > 0.02:
                    e['full_trail'].append(e['pos'].copy())
                    if len(e['full_trail']) > 2000: e['full_trail'].pop(0)

            e['vel'][0] = target_vx
            ay = 0.0
            if abs(B) > 0.001 and abs(Is) > 0.1:
                vx = e['vel'][0]; cross_y = -vx * B * B_dir_sign
                ay += charge_sign * cross_y * 8.0
            if abs(self.E_field) > 0.001:
                ay += charge_sign * self.E_field * 1.0
            ay += np.random.uniform(-0.0005, 0.0005)

            e['vel'][1] += ay * dt
            if not self.accumulation_active:
                e['vel'][1] *= 0.9
                if abs(e['vel'][1]) < 0.001: e['vel'][1] = 0.0
            e['vel'][2] *= 0.99
            e['pos'] += e['vel'] * dt

            if e['pos'][2] < 0: e['pos'][2] = 0; e['vel'][2] *= -0.6
            elif e['pos'][2] > d: e['pos'][2] = d; e['vel'][2] *= -0.6

    # ---------- 动画循环 ----------
    def animation_loop(self):
        if not self.animation_running: return
        self.update_electrons_positions()
        try: current = self.notebook.index(self.notebook.select())
        except Exception: current = 0
        if current == 2: self.update_field_distribution_plot()
        else: self.update_3d_plot()
        self.animation_timer = self.root.after(self.animation_interval, self.animation_loop)

    def start_electron_animation(self):
        if self.animation_timer is not None: self.root.after_cancel(self.animation_timer)
        if not self.electrons: self.init_electrons()
        self.animation_running = True
        self.animation_loop()

    # ---------- 数据管理 ----------
    def record_data_point(self):
        if not self.power_on:
            messagebox.showwarning("提示", "请先打开电源！"); return
        if self.measurements:
            last = self.measurements[-1]; mode = self.experiment_type.get()
            if mode == "UH-IS曲线测绘":
                if abs(last['Is/mA'] - self.Is.get()) < 0.0001:
                    messagebox.showwarning("提示", "工作电流Is没有变化！"); return
            elif mode == "UH-IM曲线测绘":
                if abs(last['Im/A'] - self.I_m.get()) < 0.0001:
                    messagebox.showwarning("提示", "励磁电流Im没有变化！"); return
            elif mode == "自由探索":
                if (abs(last['Is/mA'] - self.Is.get()) < 0.0001 and
                    abs(last['Im/A'] - self.I_m.get()) < 0.0001):
                    messagebox.showwarning("提示", "参数没有变化！"); return
        data_point = {
            '序号': len(self.measurements) + 1,
            'Is/mA': self.Is.get(), 'Im/A': self.I_m.get(), 'B/T': self.B.get(),
            'U1/mV': self.VH1.get(), 'U2/mV': self.VH2.get(),
            'U3/mV': self.VH3.get(), 'U4/mV': self.VH4.get(),
            'UH/mV': self.VH_avg.get(), 'mode': self.experiment_type.get()
        }
        self.measurements.append(data_point); self.update_data_table()

    def update_data_table(self):
        for item in self.data_tree.get_children(): self.data_tree.delete(item)
        for m in self.measurements:
            self.data_tree.insert('', 'end', values=(
                m['序号'],
                f"{m['Is/mA']:.4f}", f"{m['Im/A']:.4f}", f"{m['B/T']:.4f}",
                f"{m['U1/mV']:.4f}", f"{m['U2/mV']:.4f}",
                f"{m['U3/mV']:.4f}", f"{m['U4/mV']:.4f}",
                f"{m['UH/mV']:.4f}"
            ))

    def clear_experiment_data(self):
        self.measurements = []; self.update_data_table()

    def open_curve_window(self):
        if not self.measurements:
            messagebox.showwarning("警告", "没有实验数据，请先记录数据点！"); return
        curve_window = tk.Toplevel(self.root)
        curve_window.title("实验曲线"); curve_window.geometry("800x600")
        fig, ax = plt.subplots(figsize=(8, 6))
        mode = self.measurements[0]['mode']
        if mode == "UH-IS曲线测绘":
            Is_vals = [m['Is/mA'] for m in self.measurements]
            VH_vals = [m['UH/mV'] for m in self.measurements]
            ax.scatter(Is_vals, VH_vals, color='blue', s=50, label='测量点')
            if len(Is_vals) > 1:
                coeffs = np.polyfit(Is_vals, VH_vals, 1); fit_line = np.poly1d(coeffs)
                Is_fit = np.linspace(min(Is_vals), max(Is_vals), 100)
                ax.plot(Is_fit, fit_line(Is_fit), 'r-', label=f'拟合: y={coeffs[0]:.4f}x+{coeffs[1]:.4f}')
            ax.set_xlabel('工作电流 Is/mA'); ax.set_ylabel('霍尔电压 UH/mV')
            ax.set_title('UH-Is 关系曲线'); ax.grid(True, alpha=0.3); ax.legend()
        elif mode == "UH-IM曲线测绘":
            Im_vals = [m['Im/A'] for m in self.measurements]
            VH_vals = [m['UH/mV'] for m in self.measurements]
            ax.scatter(Im_vals, VH_vals, color='green', s=50, label='测量点')
            if len(Im_vals) > 1:
                coeffs = np.polyfit(Im_vals, VH_vals, 1); fit_line = np.poly1d(coeffs)
                Im_fit = np.linspace(min(Im_vals), max(Im_vals), 100)
                ax.plot(Im_fit, fit_line(Im_fit), 'r-', label=f'拟合: y={coeffs[0]:.4f}x+{coeffs[1]:.4f}')
            ax.set_xlabel('励磁电流 Im/A'); ax.set_ylabel('霍尔电压 UH/mV')
            ax.set_title('UH-Im 关系曲线'); ax.grid(True, alpha=0.3); ax.legend()
        else:
            indices = list(range(len(self.measurements)))
            VH_vals = [m['UH/mV'] for m in self.measurements]
            ax.plot(indices, VH_vals, 'ro-', linewidth=2, markersize=6)
            ax.set_xlabel('测量序号'); ax.set_ylabel('霍尔电压 UH/mV')
            ax.set_title('霍尔电压变化趋势'); ax.grid(True, alpha=0.3)
        canvas_curve = FigureCanvasTkAgg(fig, master=curve_window); canvas_curve.draw()
        canvas_curve.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        NavigationToolbar2Tk(canvas_curve, curve_window).update()

    # ---------- 桌面路径 ----------
    def get_desktop_path(self):
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        if not os.path.exists(desktop):
            desktop = os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop")
        if not os.path.exists(desktop): desktop = os.path.expanduser("~")
        return desktop

    # ---------- 导出数据 ----------
    def export_data(self, fmt):
        if not self.measurements:
            messagebox.showwarning("警告", "没有实验数据可导出"); return
        desktop = self.get_desktop_path()
        filename = os.path.join(desktop, f"hall_effect_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{fmt}")
        try:
            if fmt == "json":
                data = {
                    'experiment_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'experiment_mode': self.experiment_type.get(),
                    'secondary_mode': self.secondary_mode.get(),
                    'parameters': {
                        'L': round(self.L.get(), 4),
                        'b': round(self.b.get(), 4),
                        'd': round(self.d.get(), 4),
                        'KH': round(self.KH.get(), 4),
                        'sigma': round(self.sigma.get(), 4),
                        'C': round(self.C.get(), 4),
                        'alpha': round(self.alpha.get(), 4),
                        'beta': round(self.beta.get(), 4),
                        'gamma': round(self.gamma.get(), 4),
                    },
                    'measurements': [
                        {
                            '序号': m['序号'],
                            'Is/mA': round(m['Is/mA'], 4),
                            'Im/A': round(m['Im/A'], 4),
                            'B/T': round(m['B/T'], 4),
                            'U1/mV': round(m['U1/mV'], 4),
                            'U2/mV': round(m['U2/mV'], 4),
                            'U3/mV': round(m['U3/mV'], 4),
                            'U4/mV': round(m['U4/mV'], 4),
                            'UH/mV': round(m['UH/mV'], 4),
                            'mode': m['mode']
                        } for m in self.measurements
                    ]
                }
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
            else:
                fieldnames = ['序号', 'Is/mA', 'Im/A', 'B/T', 'U1/mV', 'U2/mV',
                              'U3/mV', 'U4/mV', 'UH/mV']
                with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.DictWriter(f, fieldnames=fieldnames)
                    writer.writeheader()
                    for m in self.measurements:
                        writer.writerow({
                            '序号': m['序号'],
                            'Is/mA': f"{m['Is/mA']:.4f}",
                            'Im/A': f"{m['Im/A']:.4f}",
                            'B/T': f"{m['B/T']:.4f}",
                            'U1/mV': f"{m['U1/mV']:.4f}",
                            'U2/mV': f"{m['U2/mV']:.4f}",
                            'U3/mV': f"{m['U3/mV']:.4f}",
                            'U4/mV': f"{m['U4/mV']:.4f}",
                            'UH/mV': f"{m['UH/mV']:.4f}"
                        })
            messagebox.showinfo("成功", f"数据已导出到桌面:\n{filename}")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {str(e)}")

    # ---------- Word 报告 ----------
    def print_data_report(self):
        if not self.measurements:
            messagebox.showwarning("警告", "没有实验数据可打印"); return
        if not HAS_DOCX:
            messagebox.showerror("缺少依赖", "请先安装 python-docx：\n\npip install python-docx"); return
        desktop = self.get_desktop_path()
        filename = os.path.join(desktop, f"霍尔效应实验报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx")
        try:
            doc = Document()
            title = doc.add_heading('霍尔效应实验数据报告', level=0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_heading('一、基本信息', level=1)
            t = doc.add_table(rows=5, cols=2); t.style = 'Light Grid Accent 1'
            t.cell(0,0).text = '实验日期'; t.cell(0,1).text = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            t.cell(1,0).text = '实验类型'; t.cell(1,1).text = self.experiment_type.get()
            t.cell(2,0).text = '副效应模式'; t.cell(2,1).text = self.secondary_mode.get()
            t.cell(3,0).text = '半导体类型'; t.cell(3,1).text = self.semiconductor_type.get()
            t.cell(4,0).text = '数据点数'; t.cell(4,1).text = str(len(self.measurements))
            doc.add_heading('二、实验参数', level=1)
            t2 = doc.add_table(rows=8, cols=2); t2.style = 'Light Grid Accent 1'
            params = [
                ('样品长度 L/mm', f'{self.L.get():.4f}'),
                ('样品宽度 b/mm', f'{self.b.get():.4f}'),
                ('样品厚度 d/mm', f'{self.d.get():.4f}'),
                ('灵敏度 KH (V/(A·T))', f'{self.KH.get():.4f}'),
                ('电导率 σ (A/(m·V))', f'{self.sigma.get():.4f}'),
                ('磁铁规格 C (kG/sA)', f'{self.C.get():.4f}'),
                ('不等位电势系数 α (mV/mA)', f'{self.alpha.get():.4f}'),
                ('热磁系数 β (mV/T)', f'{self.beta.get():.4f}'),
            ]
            for i, (k, v) in enumerate(params):
                t2.cell(i,0).text = k; t2.cell(i,1).text = v
            doc.add_heading('三、测量数据', level=1)
            cols = ['序号','Is/mA','Im/A','B/T','U1/mV','U2/mV','U3/mV','U4/mV','UH/mV']
            dt = doc.add_table(rows=1, cols=len(cols)); dt.style = 'Light Grid Accent 1'
            for i, c in enumerate(cols): dt.rows[0].cells[i].text = c
            for m in self.measurements:
                row = dt.add_row().cells
                row[0].text = str(m['序号'])
                row[1].text = f"{m['Is/mA']:.4f}"
                row[2].text = f"{m['Im/A']:.4f}"
                row[3].text = f"{m['B/T']:.4f}"
                row[4].text = f"{m['U1/mV']:.4f}"
                row[5].text = f"{m['U2/mV']:.4f}"
                row[6].text = f"{m['U3/mV']:.4f}"
                row[7].text = f"{m['U4/mV']:.4f}"
                row[8].text = f"{m['UH/mV']:.4f}"
            if self.measurements:
                doc.add_heading('四、统计信息', level=1)
                VH_vals = [m['UH/mV'] for m in self.measurements]
                t3 = doc.add_table(rows=4, cols=2); t3.style = 'Light Grid Accent 1'
                t3.cell(0,0).text = '霍尔电压平均值 UH/mV'; t3.cell(0,1).text = f'{np.mean(VH_vals):.4f}'
                t3.cell(1,0).text = '霍尔电压标准差 UH/mV'; t3.cell(1,1).text = f'{np.std(VH_vals):.4f}'
                t3.cell(2,0).text = '霍尔电压最大值 UH/mV'; t3.cell(2,1).text = f'{np.max(VH_vals):.4f}'
                t3.cell(3,0).text = '霍尔电压最小值 UH/mV'; t3.cell(3,1).text = f'{np.min(VH_vals):.4f}'
            doc.add_heading('五、数据处理结果', level=1)
            t4 = doc.add_table(rows=4, cols=2); t4.style = 'Light Grid Accent 1'
            t4.cell(0,0).text = '霍尔系数 RH/(m³·C⁻¹)'; t4.cell(0,1).text = f'{self.RH.get():.4f}'
            t4.cell(1,0).text = '载流子浓度 n/m⁻³'; t4.cell(1,1).text = f'{self.n.get():.4e}'
            t4.cell(2,0).text = '载流子平均漂移速率 v/(m·s⁻¹)'; t4.cell(2,1).text = f'{self.v.get():.4f}'
            t4.cell(3,0).text = '载流子迁移率 μ/(cm²·V⁻¹·s⁻¹)'; t4.cell(3,1).text = f'{self.mu.get():.4f}'
            doc.add_heading('六、备注', level=1)
            note = doc.add_paragraph()
            note.add_run('1. 本报告由霍尔效应实验虚拟仿真软件自动生成。\n')
            note.add_run('2. 霍尔电压 U1~U4 采用对称测量法（四组 ±Is, ±B 组合取绝对值平均）计算，UH 为平均值。\n')
            if self.secondary_mode.get() == '仿真':
                note.add_run('3. 当前为仿真模式，测量值中叠加了不等位电势、热磁效应和交叉效应。\n')
            else:
                note.add_run('3. 当前为纯理论模式，未叠加副效应。\n')
            doc.save(filename)
            messagebox.showinfo("成功", f"实验报告已生成到桌面:\n{filename}")
        except Exception as e:
            messagebox.showerror("错误", f"生成报告失败: {str(e)}")


def main():
    root = tk.Tk()
    app = HallEffectSimulation(root)
    root.update_idletasks()
    width = root.winfo_width(); height = root.winfo_height()
    screen_width = root.winfo_screenwidth(); screen_height = root.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    root.geometry(f'{width}x{height}+{x}+{y}')
    root.mainloop()


if __name__ == "__main__":
    main()
