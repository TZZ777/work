import os
import re
import sys
import shutil
import tkinter as tk
from tkinter import scrolledtext, messagebox, ttk, font
import threading
import queue
import subprocess
import ctypes
import unicodedata

dict_t = {
  "name": "租赁金融报告模版V0.2",
  "description": "租赁金融报告模版V0.2",
  "version": "1.0.0",
  "created_date": "2025-12-30",
  "author": "重命名工具",
  "rules": {
    "审计报告1": {"keywords": ["审计报告","年度审计"], "folders": ["承租人"], "tag": "#承租人审计报告#", "flag":True},
    "审计报告2": {"keywords": ["审计报告","年度审计"], "folders": ["担保人"], "tag": "#担保人审计报告#", "flag":True},
    "立项申请表": {"keywords": ["立项申请表","立项申请"], "folders": ["全套资料公司"], "tag": "#立项申请表#", "flag":True},
    "股权穿透图1": {"keywords": ["股权结构","穿透谱图","股权穿透谱图"], "folders": ["承租人"], "tag": "#承租人股权穿透图#", "flag":True},
    "股权穿透图2": {"keywords": ["股权结构","穿透谱图","股权穿透谱图"], "folders": ["担保人"], "tag": "#担保人股权穿透图#", "flag":True},
    "租赁物清单": {"keywords": ["租赁物清单","租赁物明细_模拟"], "folders": ["全套资料公司"], "tag": "#租赁物清单#", "flag":True},
    "预警通区域经济速览": {"keywords": ["经济速览"], "folders": ["全套资料公司"], "tag": "#经济速览#", "flag":True},
    "预警通财政数据": {"keywords": ["辖区经济.xlsx","财政.xlsx"], "folders": ["全套资料公司"], "tag": "#预警通财政数据#", "flag":True},
    "区域城投平台": {"keywords": ["城投平台(本级)","城投平台"], "folders": ["全套资料公司"], "tag": "#区域城投平台#", "flag":True},
    "财务报表1": {"keywords": ["合并财务报表","财务报表","会计报表","年度报表","年报"], "folders": ["承租人"], "tag": "#承租人财务报表#", "flag":False},
    "财务报表2": {"keywords": ["合并财务报表","财务报表","会计报表","年度报表","年报"], "folders": ["担保人"], "tag": "#担保人财务报表#", "flag":False},
    "评级报告1": {"keywords": ["评级报告","信用评级报告","跟踪评级报告"], "folders": ["承租人"], "tag": "#承租人评级报告#", "flag":False},
    "评级报告2": {"keywords": ["评级报告","信用评级报告","跟踪评级报告"], "folders": ["担保人"], "tag": "#担保人评级报告#", "flag":False},
    "募集说明书1": {"keywords": ["募集说明书","发行说明书"], "folders": ["承租人"], "tag": "#承租人募集说明书#", "flag":False},
    "募集说明书2": {"keywords": ["募集说明书","发行说明书"], "folders": ["担保人"], "tag": "#担保人募集说明书#", "flag":False},
    "公司简介1": {"keywords": ["公司简介","企业简介"], "folders": ["承租人"], "tag": "#承租人公司简介#", "flag":True},
    "公司简介2": {"keywords": ["公司简介","企业简介"], "folders": ["担保人"], "tag": "#担保人公司简介#", "flag":True},
    "担保明细1": {"keywords": ["担保明细","担保明细表","担保情况","担保"], "folders": ["承租人"], "tag": "#承租人担保明细#", "flag":True},
    "担保明细2": {"keywords": ["担保明细","担保明细表","担保情况","担保"], "folders": ["担保人"], "tag": "#担保人担保明细#", "flag":True},
    "有息负债明细1": {"keywords": ["融资明细","借款明细","银行贷款","融资.xlsx","融资.xls","负债","债务"], "folders": ["承租人"], "tag": "#承租人有息负债明细#", "flag":True},
    "有息负债明细2": {"keywords": ["融资明细","借款明细","银行贷款","融资.xlsx","融资.xls","负债","债务"], "folders": ["担保人"], "tag": "#担保人有息负债明细#", "flag":True},
    "法定代表人简历1": {"keywords": ["法人简历","法定代表人简历表","简历"], "folders": ["承租人"], "tag": "#承租人法定代表人简历#", "flag":True},
    "法定代表人简历2": {"keywords": ["法人简历","法定代表人简历表","简历"], "folders": ["担保人"], "tag": "#担保人法定代表人简历#", "flag":True},
    "征信报告1": {"keywords": ["征信报告","征信报告表","征信"], "folders": ["承租人"], "tag": "#承租人征信报告#", "flag":False},
    "征信报告2": {"keywords": ["征信报告","征信报告表","征信"], "folders": ["担保人"], "tag": "#担保人征信报告#", "flag":False},
    "法定代表人身份证1": {"keywords": ["法人身份证","法定代表人身份证表","身份证"], "folders": ["承租人"], "tag": "#承租人法定代表人身份证#", "flag":True},
    "法定代表人身份证2": {"keywords": ["法人身份证","法定代表人身份证表","身份证"], "folders": ["担保人"], "tag": "#担保人法定代表人身份证#", "flag":True},
    "企查查企业信用报告1": {"keywords": ["企业信用报告","企查查企业信用报告","【企查查】企业信用报告"], "folders": ["承租人"], "tag": "#承租人企查查企业信用报告#", "flag":True},
    "企查查企业信用报告2": {"keywords": ["企业信用报告","企查查企业信用报告","【企查查】企业信用报告"], "folders": ["担保人"], "tag": "#担保人企查查企业信用报告#", "flag":True},
    "企查查司法案件1": {"keywords": ["司法案件","企查查司法案件","【企查查】司法案件"], "folders": ["承租人"], "tag": "#承租人企查查司法案件#", "flag":False},
    "企查查司法案件2": {"keywords": ["司法案件","企查查司法案件","【企查查】司法案件"], "folders": ["担保人"], "tag": "#担保人企查查司法案件#", "flag":False},
    "中登网融资租赁1": {"keywords": ["中登网","中登网融资租赁","按担保人查询登记信息列表"], "folders": ["承租人"], "tag": "#承租人中登网融资租赁#", "flag":True},
    "中登网融资租赁2": {"keywords": ["中登网","中登网融资租赁","按担保人查询登记信息列表"], "folders": ["担保人"], "tag": "#担保人中登网融资租赁#", "flag":True},
    "基础材料1": {"keywords": ["基础材料","基础资料","基础情况"], "folders": ["承租人"], "tag": "#承租人基础材料#", "flag":False},
    "基础材料2": {"keywords": ["基础材料","基础资料","基础情况"], "folders": ["担保人"], "tag": "#担保人基础材料#", "flag":False},
    "项目尽调报告": {"keywords": ["项目调查报告","项目尽调报告","尽调报告"], "folders": ["全套资料公司"], "tag": "#项目尽调报告#", "flag":True},
    "地方国企类业务指引": {"keywords": ["地方国企类","业务指引"], "folders": ["全套资料公司"], "tag": "#地方国企类业务指引#", "flag":True},
    "投放数据台账": {"keywords": ["投放数据台账"], "folders": ["全套资料公司"], "tag": "#投放数据台账#", "flag":True},
    "土地成交统计_按企业类型": {"keywords": ["土地成交统计_按企业类型"], "folders": ["全套资料公司"], "tag": "#土地成交统计_按企业类型#", "flag":True},
    "土地出让明细": {"keywords": ["土地出让明细"], "folders": ["全套资料公司"], "tag": "#土地出让明细#", "flag":True},
    "风控部工作细则": {"keywords": ["风控部工作细则"], "folders": ["全套资料公司"], "tag": "#风控部工作细则#", "flag":True},
    "融资情况": {"keywords": ["融资情况"], "folders": ["全套资料公司"], "tag": "#融资情况#", "flag":True}
  }
}

def get_app_path(relative_path=""):
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
        if sys.platform == 'darwin' and '.app/Contents/MacOS' in base_path:
            base_path = os.path.abspath(os.path.join(base_path, '../../..'))
    else:
        base_path = os.path.abspath(".")
    
    full_path = os.path.join(base_path, relative_path)
    if relative_path and not os.path.exists(full_path):
        try:
            os.makedirs(full_path, exist_ok=True)
        except:
            fallback_base = os.path.join(os.path.expanduser("~"), "RenameTool")
            full_path = os.path.join(fallback_base, relative_path)
            try:
                os.makedirs(full_path, exist_ok=True)
            except:
                pass
    return os.path.abspath(full_path)

def ensure_dir(path):
    p = os.path.abspath(path)
    if not os.path.exists(p):
        os.makedirs(p, exist_ok=True)
    return p

def normalize_text(text):
    if text is None:
        return ""
    normalized = unicodedata.normalize("NFKC", str(text)).lower()
    return re.sub(r"[\s\-_()（）\[\]【】]", "", normalized)

class UniversalFileRenamer:
    def __init__(self, template_dict):
        self.current_template = template_dict.get("name", "未命名模板")
        self.file_rules = template_dict.get("rules", {})
        default_extensions = [".pdf", ".doc", ".docx", ".xlsx", ".xls", ".png", ".jpg", ".jpeg"]
        self.supported_extensions = [ext.lower() for ext in default_extensions]

    def find_target_files(self, base_folder):
        base_folder = os.path.abspath(base_folder)
        found_files = {file_type: [] for file_type in self.file_rules.keys()}
        
        for root, dirs, files in os.walk(base_folder):
            rel_path = os.path.relpath(root, base_folder)
            path_parts = [] if rel_path == '.' else rel_path.split(os.sep)
            
            for file_type, rules in self.file_rules.items():
                target_folders = rules.get("folders", [])
                folder_matched = False
                
                if not target_folders or "全套资料公司" in target_folders:
                    folder_matched = True
                else:
                    for target_folder in target_folders:
                        if any(target_folder in part for part in path_parts):
                            folder_matched = True
                            break
                
                if folder_matched:
                    folder_display = rel_path if rel_path != '.' else "根目录"
                    self._search_files_in_folder(root, folder_display, file_type, rules, found_files)
        return found_files

    def _search_files_in_folder(self, folder_path, folder_name, file_type, rules, found_files):
        try:
            for file in os.listdir(folder_path):
                file_path = os.path.abspath(os.path.join(folder_path, file))
                if file.lower().endswith(tuple(self.supported_extensions)) and os.path.isfile(file_path):
                    if '#' in file: continue
                    normalized_file = normalize_text(file)
                    if any(normalize_text(keyword) in normalized_file for keyword in rules["keywords"] if keyword):
                        if not any(f['path'] == file_path for f in found_files[file_type]):
                            found_files[file_type].append({
                                'path': file_path,
                                'folder': folder_name,
                                'filename': file,
                                'relative_path': os.path.join(folder_name, file)
                            })
        except: pass

    def generate_new_name(self, file_info, file_type):
        name_without_ext, extension = os.path.splitext(file_info['filename'])
        tag = self.file_rules[file_type]["tag"]
        return f"{name_without_ext}{tag}{extension}"

def direct_batch_process(rules_dict):
    data_folder = get_app_path("data")
    output_folder = get_app_path("output")
    missing_tags_report = {}
    
    ensure_dir(data_folder)
    ensure_dir(output_folder)

    material_packages = [os.path.join(data_folder, d) for d in os.listdir(data_folder) 
                         if os.path.isdir(os.path.join(data_folder, d)) and d.lower() != "output"]
    
    if not material_packages: return {}

    renamer = UniversalFileRenamer(rules_dict)
    for package_path in material_packages:
        package_name = os.path.basename(package_path)
        target_path = os.path.join(output_folder, package_name)
        
        try:
            found_files = renamer.find_target_files(package_path)
            
            missing_in_package = []
            keyword_issues_in_package = []
            for rule_name, files in found_files.items():
                rule_info = rules_dict["rules"].get(rule_name)
                if not files:
                    missing_in_package.append({
                        "rule": rule_name, "tag": rule_info.get("tag"),
                        "keywords": rule_info.get("keywords"), "flag": rule_info.get("flag")
                    })
                else:
                    keywords = rule_info.get("keywords", [])
                    normalized_files = [normalize_text(f["filename"]) for f in files]
                    unmatched = [
                        kw for kw in keywords
                        if normalize_text(kw) and not any(normalize_text(kw) in nf for nf in normalized_files)
                    ]
                    if unmatched:
                        keyword_issues_in_package.append({
                            "rule": rule_name,
                            "tag": rule_info.get("tag"),
                            "unmatched_keywords": unmatched
                        })
            if missing_in_package or keyword_issues_in_package:
                missing_tags_report[package_name] = {
                    "missing": missing_in_package,
                    "unmatched_keywords": keyword_issues_in_package
                }

            if any(len(f) > 0 for f in found_files.values()):
                if os.path.exists(target_path): shutil.rmtree(target_path)
                shutil.copytree(package_path, target_path)
                
                found_in_target = renamer.find_target_files(target_path)
                for file_type, files in found_in_target.items():
                    for f_info in files:
                        new_name = renamer.generate_new_name(f_info, file_type)
                        new_p = os.path.join(os.path.dirname(f_info['path']), new_name)
                        if os.path.exists(f_info['path']) and not os.path.exists(new_p):
                            os.rename(f_info['path'], new_p)
        except Exception as e:
            print(f"Error: {e}")
            
    return missing_tags_report

class RenameApp:
    def __init__(self, root):
        self.root = root
        self.root.title("通用文件重命名工具")
        self.root.geometry("600x250")
        self.root.resizable(False, False)
        self.root.configure(bg="#f8f9fa")
        
        font_family = "Microsoft YaHei" if sys.platform == "win32" else "Helvetica"
        log_font_family = "Consolas" if sys.platform == "win32" else "Menlo"
        self.title_font = font.Font(family=font_family, size=14, weight="bold")
        self.normal_font = font.Font(family=font_family, size=10)
        self.log_font = font.Font(family=log_font_family, size=10)
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        styles = {
            "Primary.TButton": ("#007bff", "#0056b3"),
            "Secondary.TButton": ("#6c757d", "#545b62"),
            "Success.TButton": ("#28a745", "#218838"),
            "Danger.TButton": ("#dc3545", "#c82333")
        }
        for name, colors in styles.items():
            self.style.configure(name, font=self.normal_font, padding=8, background=colors[0], foreground="white", borderwidth=0)
            self.style.map(name, background=[('active', colors[1]), ('disabled', '#cccccc')])

        header = tk.Frame(root, bg="white", height=60, relief="flat")
        header.pack(fill=tk.X, side=tk.TOP)
        tk.Label(header, text="通用文件重命名工具", font=self.title_font, bg="white", fg="#333333").pack(pady=15)

        main_frame = tk.Frame(root, bg="#f8f9fa")
        main_frame.pack(expand=True, fill='both', padx=30, pady=20)
        
        btn_container = tk.Frame(main_frame, bg="#f8f9fa")
        btn_container.pack(pady=10)
        
        self.btn_start = ttk.Button(btn_container, text="🚀 开始重命名", style="Success.TButton", command=self.start_process)
        self.btn_start.grid(row=0, column=0, padx=5, pady=5)
        
        self.btn_open_data = ttk.Button(btn_container, text="📁 输入(data)", style="Primary.TButton", command=self.open_data)
        self.btn_open_data.grid(row=0, column=1, padx=5, pady=5)
        
        self.btn_open_output = ttk.Button(btn_container, text="📂 输出(output)", style="Primary.TButton", command=self.open_output)
        self.btn_open_output.grid(row=0, column=2, padx=5, pady=5)
        
        self.btn_zip = ttk.Button(btn_container, text="🗜️ 压缩", style="Primary.TButton", command=self.compress_output)
        self.btn_zip.grid(row=0, column=3, padx=5, pady=5)
        
        self.queue = queue.Queue()

    def start_process(self):
        data_path = get_app_path("data")
        if not os.path.exists(data_path) or not any(os.path.isdir(os.path.join(data_path, i)) for i in os.listdir(data_path) if i != ".DS_Store"):
            messagebox.showwarning("提示", "data 文件夹内没有材料包！")
            return
        self.btn_start.config(state='disabled')
        threading.Thread(target=self.run_process, daemon=True).start()

    def run_process(self):
        try:
            report = direct_batch_process(dict_t)
            self.root.after(0, lambda: self.show_popup(report))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
        finally:
            self.root.after(0, lambda: self.btn_start.config(state='normal'))

    def show_popup(self, report):
        if not report:
            messagebox.showinfo("提示", "处理完成！未发现缺失。")
            return
        top = tk.Toplevel(self.root); top.title("分析报告"); top.geometry("700x500")
        header = tk.Frame(top, bg="#fff3cd"); header.pack(fill=tk.X)
        tk.Label(header, text="检测到部分规则未匹配到文件：", font=("Microsoft YaHei", 10, "bold"), bg="#fff3cd").pack(pady=10)
        txt = scrolledtext.ScrolledText(top, font=self.log_font, padx=10, pady=10)
        txt.pack(expand=True, fill='both')
        content = ""
        for pkg, details in report.items():
            content += f"📦 材料包: {pkg}\n{'='*60}\n"
            missing = details.get("missing", [])
            keyword_issues = details.get("unmatched_keywords", [])
            if missing:
                for i, m in enumerate(missing, 1):
                    status = "【必须】" if m['flag'] else "【可选】"
                    content += f" {i}. {status} 缺失: {m['tag']}\n    关键词: {', '.join(m['keywords'])}\n"
            if keyword_issues:
                content += "\n  关键词未匹配到文件名:\n"
                for i, item in enumerate(keyword_issues, 1):
                    content += f"  {i}. {item['tag']}\n    未匹配关键词: {', '.join(item['unmatched_keywords'])}\n"
            content += "\n"
        txt.insert(tk.END, content); txt.config(state='disabled')
        ttk.Button(top, text="确定", style="Secondary.TButton", command=top.destroy).pack(pady=10)

    def open_data(self): open_folder(get_app_path("data"))
    def open_output(self): open_folder(get_app_path("output"))

    def compress_output(self):
        output_path = get_app_path("output")
        ensure_dir(output_path)
        subdirs = [d for d in os.listdir(output_path) if os.path.isdir(os.path.join(output_path, d))]
        if not subdirs:
            messagebox.showinfo("提示", "output 文件夹内没有可压缩的子文件夹")
            return
        self.btn_zip.config(state='disabled')
        threading.Thread(target=self._zip_subfolders, args=(output_path, subdirs), daemon=True).start()

    def _zip_subfolders(self, output_path, subdirs):
        try:
            for name in subdirs:
                folder_path = os.path.join(output_path, name)
                zip_base = os.path.join(output_path, name)
                zip_file = zip_base + ".zip"
                if os.path.exists(zip_file):
                    try:
                        os.remove(zip_file)
                    except:
                        pass
                shutil.make_archive(zip_base, 'zip', root_dir=folder_path)
            self.root.after(0, lambda: messagebox.showinfo("提示", "压缩完成！"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"压缩失败: {e}"))
        finally:
            self.root.after(0, lambda: self.btn_zip.config(state='normal'))

def open_folder(path):
    path = os.path.abspath(path)
    ensure_dir(path)
    try:
        if sys.platform == 'darwin':
            subprocess.run(['open', path], check=True)
        elif sys.platform == 'win32':
            os.startfile(path)
        else:
            subprocess.run(['xdg-open', path], check=True)
    except Exception as e:
        messagebox.showerror("错误", f"无法打开文件夹: {e}")

def main():
    if sys.platform == 'win32':
        try: ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except: pass
    root = tk.Tk()
    app = RenameApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
