import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import re
import csv
import datetime
import os

# --- ドラッグ＆ドロップ用ライブラリ ---
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False
    print("【注意】tkinterdnd2 が入っていません。")

# ==========================================
# 1. 解析ロジック (G00と切削を分離)
# ==========================================
class BlockData:
    def __init__(self, name):
        self.name = name
        # 2つの辞書で管理 (rapid=早送り, cut=切削)
        self.rapid = {"X": [], "Y": [], "Z": []}
        self.cut = {"X": [], "Y": [], "Z": []}
        self.s_vals = []
        self.f_vals = []
        self.errors = []

    def add_val(self, axis, val, is_rapid):
        # モードに応じて格納先を変える
        target = self.rapid if is_rapid else self.cut
        if axis in target:
            target[axis].append(val)
    
    def get_range_str(self, axis, mode="both"):
        """指定された軸とモードの最小～最大を文字列で返す"""
        vals = []
        if mode == "rapid" or mode == "both":
            vals.extend(self.rapid[axis])
        if mode == "cut" or mode == "both":
            vals.extend(self.cut[axis])
            
        if not vals:
            return "-"
        return f"{min(vals):.3f} ~ {max(vals):.3f}"

    def get_raw_min_max(self, axis, mode="both"):
        """数値としてMin/Maxを返す（CSVやリミットチェック用）"""
        vals = []
        if mode == "rapid" or mode == "both":
            vals.extend(self.rapid[axis])
        if mode == "cut" or mode == "both":
            vals.extend(self.cut[axis])
        return (min(vals), max(vals)) if vals else (None, None)

    def get_max_s_f(self):
        max_s = max(self.s_vals) if self.s_vals else 0
        max_f = max(self.f_vals) if self.f_vals else 0
        return max_s, max_f

class NCAnalyzer:
    def __init__(self):
        self.reset()

    def reset(self):
        self.blocks = []
        self.current_block = None

    def parse_value(self, text):
        try: return float(text)
        except ValueError: return None

    def analyze(self, nc_code, machine_type):
        self.reset()
        self.current_block = BlockData("Header / Setup")
        self.blocks.append(self.current_block)

        lines = nc_code.split('\n')
        
        re_comment = re.compile(r'\((.*?)\)')
        re_coord = re.compile(r'([XYZ])\s*([-]?\d+\.?\d*)')
        re_s = re.compile(r'S\s*(\d+)')
        re_f = re.compile(r'F\s*(\d+\.?\d*)')
        re_g = re.compile(r'G(\d+)')

        has_g50_global = False
        is_rapid_mode = True # デフォルトはG00とする

        for i, line in enumerate(lines):
            line_num = i + 1
            line_content = line.split(';')[0].strip()
            if not line_content: continue

            # ブロック切り替え
            comment_match = re_comment.search(line)
            if comment_match:
                block_name = comment_match.group(1).strip()
                self.current_block = BlockData(f"Line{line_num}: {block_name}")
                self.blocks.append(self.current_block)

            # Gコードによるモード判定
            g_codes = [int(g) for g in re_g.findall(line_content.upper())]
            if 0 in g_codes:
                is_rapid_mode = True
            if any(g in [1, 2, 3] for g in g_codes):
                is_rapid_mode = False

            # 座標値の取得と振り分け
            coords = re_coord.findall(line_content.upper())
            for axis, val_str in coords:
                val = self.parse_value(val_str)
                if val is not None:
                    # ここでモード情報を渡す！
                    self.current_block.add_val(axis, val, is_rapid_mode)

            # S, F
            s_match = re_s.search(line_content.upper())
            if s_match: self.current_block.s_vals.append(float(s_match.group(1)))
            f_match = re_f.search(line_content.upper())
            if f_match: self.current_block.f_vals.append(float(f_match.group(1)))

            # 簡易チェック
            if machine_type == "FANUC_Lathe":
                if 50 in g_codes: has_g50_global = True
                if 96 in g_codes and not has_g50_global:
                    self.current_block.errors.append(f"[Line {line_num}] 危険: G50なしでG96使用")

        return self.blocks

    def get_global_stats(self):
        # 全体の最大最小（Rapid/Cut込み）
        stats = { "X": [], "Y": [], "Z": [], "S": [], "F": [] }
        for blk in self.blocks:
            # Rapid
            for ax in ["X", "Y", "Z"]:
                stats[ax].extend(blk.rapid[ax])
                stats[ax].extend(blk.cut[ax])
            stats["S"].extend(blk.s_vals)
            stats["F"].extend(blk.f_vals)
        
        result = {}
        for key in ["X", "Y", "Z"]:
            vals = stats[key]
            result[key] = (min(vals), max(vals)) if vals else (None, None)
        
        result["max_s"] = max(stats["S"]) if stats["S"] else 0
        result["max_f"] = max(stats["F"]) if stats["F"] else 0
        return result

# ==========================================
# 2. アプリ画面
# ==========================================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("NC解析 Ver.5 (G00/Cutting Separated)")
        
        try: self.root.state('zoomed')
        except: self.root.geometry("1200x900")

        self.analyzer = NCAnalyzer()
        self.current_file_name = "未選択"

        if DND_AVAILABLE:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.drop_file)

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", rowheight=25, font=("Meiryo UI", 9))
        style.configure("Treeview.Heading", font=("Meiryo UI", 9, "bold"), background="#ddd")

        # --- Top Menu ---
        frame_top = tk.Frame(root, padx=10, pady=5, bg="#f0f0f0")
        frame_top.pack(fill=tk.X)

        tk.Label(frame_top, text="制御装置:", bg="#f0f0f0").pack(side=tk.LEFT)
        self.combo_machine = ttk.Combobox(frame_top, values=["FANUC_Lathe", "OSP_Mill", "TOSNUC_Mill"], state="readonly", width=15)
        self.combo_machine.current(0)
        self.combo_machine.pack(side=tk.LEFT, padx=5)

        btn_open = tk.Button(frame_top, text="📂 ファイル読込", command=self.open_file_dialog, width=15)
        btn_open.pack(side=tk.LEFT, padx=10)

        btn_limit = tk.Button(frame_top, text="📏 範囲チェック", command=self.open_limit_checker, bg="#FF5722", fg="white", font=("Meiryo UI", 9, "bold"))
        btn_limit.pack(side=tk.LEFT, padx=10)

        btn_csv = tk.Button(frame_top, text="💾 CSV保存", command=self.save_csv, bg="#009688", fg="white", font=("Meiryo UI", 9, "bold"))
        btn_csv.pack(side=tk.LEFT, padx=10)

        # --- Dashboard ---
        self.frame_dashboard = tk.LabelFrame(root, text="📊 プログラム全体統計 (Global Stats)", padx=10, pady=5, font=("Meiryo UI", 11, "bold"), fg="#1565C0")
        self.frame_dashboard.pack(fill=tk.X, padx=10, pady=5)
        
        self.lbl_filename = tk.Label(self.frame_dashboard, text="File: 未選択", font=("Meiryo UI", 10, "bold"), fg="#555")
        self.lbl_filename.pack(anchor="w")

        frame_stats_inner = tk.Frame(self.frame_dashboard)
        frame_stats_inner.pack(fill=tk.X, pady=5)

        self.lbl_global_x = tk.Label(frame_stats_inner, text="X: ---", font=("Consolas", 12), width=25, anchor="w", fg="#333")
        self.lbl_global_y = tk.Label(frame_stats_inner, text="Y: ---", font=("Consolas", 12), width=25, anchor="w", fg="#333")
        self.lbl_global_z = tk.Label(frame_stats_inner, text="Z: ---", font=("Consolas", 12), width=25, anchor="w", fg="#333")
        self.lbl_global_sf = tk.Label(frame_stats_inner, text="S-Max: --- / F-Max: ---", font=("Consolas", 12), width=30, anchor="w", fg="#D32F2F")

        self.lbl_global_x.pack(side=tk.LEFT, padx=10)
        self.lbl_global_y.pack(side=tk.LEFT, padx=10)
        self.lbl_global_z.pack(side=tk.LEFT, padx=10)
        self.lbl_global_sf.pack(side=tk.LEFT, padx=10)

        # --- Main Layout ---
        self.paned = tk.PanedWindow(root, orient=tk.VERTICAL, sashwidth=4, bg="#ccc")
        self.paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Input
        frame_input = tk.Frame(self.paned)
        tk.Label(frame_input, text="▼ NCコード入力", font=("Meiryo UI", 9)).pack(anchor="w")
        self.txt_input = scrolledtext.ScrolledText(frame_input, height=5)
        self.txt_input.pack(fill=tk.BOTH, expand=True)
        self.paned.add(frame_input, height=100)

        # Table (Horizontal Scroll added!)
        frame_table = tk.Frame(self.paned)
        tk.Label(frame_table, text="▼ ブロック別詳細データ (G00=早送り / Cut=切削)", font=("Meiryo UI", 9, "bold"), fg="#00796B").pack(anchor="w")
        
        # 列定義：G00とCutで分ける
        columns = (
            "Block", 
            "G00 X", "Cut X", 
            "G00 Y", "Cut Y", 
            "G00 Z", "Cut Z", 
            "Max S", "Max F", "Status"
        )
        self.tree = ttk.Treeview(frame_table, columns=columns, show="headings")
        
        for col in columns:
            self.tree.heading(col, text=col)
            w = 120 # 少し広げる
            if col == "Block": w = 200
            elif col == "Status": w = 80
            self.tree.column(col, width=w, anchor="center" if col != "Block" else "w")

        # スクロールバー設定 (縦と横)
        ysb = ttk.Scrollbar(frame_table, orient=tk.VERTICAL, command=self.tree.yview)
        xsb = ttk.Scrollbar(frame_table, orient=tk.HORIZONTAL, command=self.tree.xview)
        
        self.tree.configure(yscroll=ysb.set, xscroll=xsb.set)
        
        ysb.pack(side=tk.RIGHT, fill=tk.Y)
        xsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        self.tree.tag_configure('odd', background='#E8F5E9')
        self.tree.tag_configure('even', background='white')
        self.tree.tag_configure('error', background='#FFEBEE')

        self.paned.add(frame_table, height=500, stretch="always")

        # Log
        frame_log = tk.Frame(self.paned)
        tk.Label(frame_log, text="▼ ログ", font=("Meiryo UI", 9)).pack(anchor="w")
        self.txt_log = scrolledtext.ScrolledText(frame_log, height=5, bg="#fafafa", fg="#d32f2f")
        self.txt_log.pack(fill=tk.BOTH, expand=True)
        self.paned.add(frame_log, height=100)

    # --- Limit Checker ---
    def open_limit_checker(self):
        if not hasattr(self, 'analyzed_blocks') or not self.analyzed_blocks:
            messagebox.showwarning("データなし", "解析してください")
            return
        
        stats = self.analyzer.get_global_stats()
        win = tk.Toplevel(self.root)
        win.title("📏 範囲チェック")
        win.geometry("500x400")
        
        tk.Label(win, text="リミット値を入力 (空欄は無視)").pack(pady=5)
        frame_grid = tk.Frame(win, padx=20, pady=5)
        frame_grid.pack()

        entries = {}
        row = 0
        tk.Label(frame_grid, text="Axis").grid(row=row, column=0)
        tk.Label(frame_grid, text="Min Limit").grid(row=row, column=1)
        tk.Label(frame_grid, text="Max Limit").grid(row=row, column=2)
        row += 1

        for axis in ["X", "Y", "Z"]:
            tk.Label(frame_grid, text=f"{axis}軸:", font="bold").grid(row=row, column=0)
            em = tk.Entry(frame_grid, width=15); em.grid(row=row, column=1, padx=5)
            ex = tk.Entry(frame_grid, width=15); ex.grid(row=row, column=2, padx=5)
            entries[axis] = (em, ex)
            row += 1

        lbl_result = tk.Label(win, text="待機中...", bg="#eee", width=40, height=5)
        lbl_result.pack(pady=20)

        def check():
            results = []
            safe = True
            for axis in ["X", "Y", "Z"]:
                p_min, p_max = stats[axis]
                if p_min is None: continue
                
                u_min = entries[axis][0].get()
                u_max = entries[axis][1].get()
                
                if u_min:
                    try: 
                        if p_min < float(u_min): 
                            results.append(f"❌ {axis} Min違反: {p_min} < {u_min}")
                            safe = False
                    except: pass
                if u_max:
                    try:
                        if p_max > float(u_max):
                            results.append(f"❌ {axis} Max違反: {p_max} > {u_max}")
                            safe = False
                    except: pass
            
            if safe and not results: lbl_result.config(text="✅ SAFE", bg="#C8E6C9", fg="green")
            elif safe: lbl_result.config(text="⚠️ 条件なし", bg="#FFF9C4")
            else: lbl_result.config(text="\n".join(results), bg="#FFCDD2", fg="red")

        tk.Button(win, text="判定", command=check, bg="#2196F3", fg="white", width=20).pack()

    # --- File Ops ---
    def load_and_run(self, file_path):
        self.txt_input.delete("1.0", tk.END)
        clean_path = file_path.strip('{}')
        self.current_file_name = os.path.basename(clean_path)
        self.lbl_filename.config(text=f"File: {self.current_file_name}")

        try:
            with open(clean_path, 'r', encoding='cp932') as f: content = f.read()
        except:
            try:
                with open(clean_path, 'r', encoding='utf-8') as f: content = f.read()
            except Exception as e:
                messagebox.showerror("Error", str(e)); return

        self.txt_input.insert(tk.END, content)
        self.run_analysis()

    def drop_file(self, event):
        if event.data:
            raw = event.data
            path = raw.split('}')[0] + '}' if raw.startswith('{') else raw.split()[0]
            self.load_and_run(path)

    def open_file_dialog(self):
        path = filedialog.askopenfilename()
        if path: self.load_and_run(path)

    def run_analysis(self):
        code = self.txt_input.get("1.0", tk.END).strip()
        machine = self.combo_machine.get()
        
        self.analyzed_blocks = self.analyzer.analyze(code, machine)
        global_stats = self.analyzer.get_global_stats()
        
        def fmt_mm(vals):
            min_v, max_v = vals
            return f"{min_v:.3f} ~ {max_v:.3f}" if min_v is not None else "---"

        self.lbl_global_x.config(text=f"X: {fmt_mm(global_stats['X'])}")
        self.lbl_global_y.config(text=f"Y: {fmt_mm(global_stats['Y'])}")
        self.lbl_global_z.config(text=f"Z: {fmt_mm(global_stats['Z'])}")
        self.lbl_global_sf.config(text=f"Smax: {int(global_stats['max_s'])} / Fmax: {global_stats['max_f']:.1f}")

        self.update_table(self.analyzed_blocks)
        self.update_log(self.analyzed_blocks)

    def update_table(self, blocks):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for i, blk in enumerate(blocks):
            max_s, max_f = blk.get_max_s_f()
            status = "WARN" if blk.errors else "OK"
            
            # 各モードの範囲文字列を取得 ("min ~ max" または "-")
            g0_x = blk.get_range_str("X", "rapid")
            cut_x = blk.get_range_str("X", "cut")
            g0_y = blk.get_range_str("Y", "rapid")
            cut_y = blk.get_range_str("Y", "cut")
            g0_z = blk.get_range_str("Z", "rapid")
            cut_z = blk.get_range_str("Z", "cut")
            
            values = (
                blk.name, 
                g0_x, cut_x, 
                g0_y, cut_y, 
                g0_z, cut_z, 
                int(max_s), fmt(max_f) if max_f else "-", 
                status
            )
            
            tag = 'error' if blk.errors else ('odd' if i % 2 == 0 else 'even')
            self.tree.insert("", tk.END, values=values, tags=(tag,))

    def update_log(self, blocks):
        self.txt_log.delete("1.0", tk.END)
        count = 0
        for blk in blocks:
            for err in blk.errors:
                self.txt_log.insert(tk.END, f"❌ {blk.name}: {err}\n")
                count += 1
        if count == 0: self.txt_log.insert(tk.END, "✅ エラーなし\n")

    def save_csv(self):
        if not hasattr(self, 'analyzed_blocks'): return
        filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")], initialfile=f"Report_{self.current_file_name}.csv")
        if not filename: return
        try:
            with open(filename, 'w', newline='', encoding='cp932') as f:
                writer = csv.writer(f)
                # ヘッダーも詳細に分割
                header = [
                    "Block Name", 
                    "G00 Min X", "G00 Max X", "Cut Min X", "Cut Max X",
                    "G00 Min Y", "G00 Max Y", "Cut Min Y", "Cut Max Y",
                    "G00 Min Z", "G00 Max Z", "Cut Min Z", "Cut Max Z",
                    "Max S", "Max F", "Errors"
                ]
                writer.writerow(header)
                for blk in self.analyzed_blocks:
                    # それぞれの生データを取得
                    g0_x = blk.get_raw_min_max("X", "rapid")
                    cut_x = blk.get_raw_min_max("X", "cut")
                    g0_y = blk.get_raw_min_max("Y", "rapid")
                    cut_y = blk.get_raw_min_max("Y", "cut")
                    g0_z = blk.get_raw_min_max("Z", "rapid")
                    cut_z = blk.get_raw_min_max("Z", "cut")
                    
                    s, f_val = blk.get_max_s_f()
                    err = " / ".join(blk.errors) if blk.errors else "OK"

                    # リストを展開して行を作成
                    row = [blk.name]
                    for val_pair in [g0_x, cut_x, g0_y, cut_y, g0_z, cut_z]:
                        row.extend([val_pair[0] if val_pair[0] is not None else "", 
                                    val_pair[1] if val_pair[1] is not None else ""])
                    row.extend([s, f_val, err])
                    
                    writer.writerow(row)
            messagebox.showinfo("完了", "CSV保存完了！")
        except Exception as e: messagebox.showerror("Error", str(e))

def fmt(v): return f"{v:.1f}"

if __name__ == "__main__":
    if DND_AVAILABLE: root = TkinterDnD.Tk()
    else: root = tk.Tk()
    app = App(root)
    root.mainloop()