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
dict_t = {
  "name": "租赁金融报告模版V0.2",
  "description": "租赁金融报告模版V0.2",
  "version": "1.0.0",
  "created_date": "2025-12-30",
  "author": "重命名工具",
  "rules": {
    "审计报告1": {
      "keywords": ["审计报告","年度审计"],
      "folders": ["承租人"],
      "tag": "#承租人审计报告#"
     ,"flag":True
    },
    "审计报告2": {
      "keywords": ["审计报告","年度审计"],
      "folders": ["担保人"],
      "tag": "#担保人审计报告#"
     ,"flag":True
    },
    "立项申请表": {
      "keywords": ["立项申请表","立项申请"],
      "folders": ["全套资料公司"],
      "tag": "#立项申请表#"
     ,"flag":True
    },
    "股权穿透图1": {
      "keywords": ["股权结构","穿透谱图","股权穿透谱图"],
      "folders": ["承租人"],
      "tag": "#承租人股权穿透图#"
     ,"flag":True
    },
    "股权穿透图2": {
      "keywords": ["股权结构","穿透谱图","股权穿透谱图"],
      "folders": ["担保人"],
      "tag": "#担保人股权穿透图#"
     ,"flag":True
    },
    "租赁物清单": {
      "keywords": ["租赁物清单","租赁物明细_模拟"],
      "folders": ["全套资料公司"],
      "tag": "#租赁物清单#"
     ,"flag":True
    },
    "预警通区域经济速览": {
      "keywords": ["经济速览"],
      "folders": ["全套资料公司"],
      "tag": "#经济速览#"
     ,"flag":True
    },
    "预警通财政数据": {
      "keywords": ["辖区经济.xlsx","财政.xlsx"],
      "folders": ["全套资料公司"],
      "tag": "#预警通财政数据#"
     ,"flag":True
    },
    "区域城投平台": {
      "keywords": ["城投平台(本级)","城投平台"],
      "folders": ["全套资料公司"],
      "tag": "#区域城投平台#"
     ,"flag":True
    },
    "财务报表1": {
      "keywords": ["合并财务报表","财务报表","会计报表","年度报表"],
      "folders": ["承租人"],
      "tag": "#承租人财务报表#"
     ,"flag":False
    },
    "财务报表2": {
      "keywords": ["合并财务报表","财务报表","会计报表","年度报表"],
      "folders": ["担保人"],
      "tag": "#担保人财务报表#"
     ,"flag":False
    },
    "评级报告1": {
      "keywords": ["评级报告","信用评级报告","跟踪评级报告"],
      "folders": ["承租人"],
      "tag": "#承租人评级报告#"
     ,"flag":False
    },
    "评级报告2": {
      "keywords": ["评级报告","信用评级报告","跟踪评级报告"],
      "folders": ["担保人"],
      "tag": "#担保人评级报告#"
     ,"flag":False
    },
    "募集说明书1": {
      "keywords": ["募集说明书","发行说明书"],
      "folders": ["承租人"],
      "tag": "#承租人募集说明书#"
     ,"flag":False
    },
    "募集说明书2": {
      "keywords": ["募集说明书","发行说明书"],
      "folders": ["担保人"],
      "tag": "#担保人募集说明书#"
     ,"flag":False
    },
    "公司简介1": {
      "keywords": ["公司简介","企业简介"],
      "folders": ["承租人"],
      "tag": "#承租人公司简介#"
     ,"flag":True
    },
    "公司简介2": {
      "keywords": ["公司简介","企业简介"],
      "folders": ["担保人"],
      "tag": "#担保人公司简介#"
     ,"flag":True
    },
    "担保明细1": {
      "keywords": ["担保明细","担保明细表","担保情况"],
      "folders": ["承租人"],
      "tag": "#承租人担保明细#"
     ,"flag":True
    },
    "担保明细2": {
      "keywords": ["担保明细","担保明细表","担保情况"],
      "folders": ["担保人"],
      "tag": "#担保人担保明细#"
     ,"flag":True
    },
    "有息负债明细1": {
      "keywords": ["融资明细","借款明细","银行贷款","融资.xlsx","融资.xls"],
      "folders": ["承租人"],
      "tag": "#承租人有息负债明细#"
     ,"flag":True
    },
    "有息负债明细2": {
      "keywords": ["融资明细","借款明细","银行贷款","融资.xlsx","融资.xls"],
      "folders": ["担保人"],
      "tag": "#担保人有息负债明细#"
     ,"flag":True
    },
    "法定代表人简历1": {
      "keywords": ["法人简历","法定代表人简历表"],
      "folders": ["承租人"],
      "tag": "#承租人法定代表人简历#"
     ,"flag":True
    },
    "法定代表人简历2": {
      "keywords": ["法人简历","法定代表人简历表"],
      "folders": ["担保人"],
      "tag": "#担保人法定代表人简历#"
     ,"flag":True
    },
    "征信报告1": {
      "keywords": ["征信报告","征信报告表","征信"],
      "folders": ["承租人"],
      "tag": "#承租人征信报告#"
     ,"flag":False
    },
    "征信报告2": {
      "keywords": ["征信报告","征信报告表","征信"],
      "folders": ["担保人"],
      "tag": "#担保人征信报告#"
     ,"flag":False
    },
    "法定代表人身份证1": {
      "keywords": ["法人身份证","法定代表人身份证表"],
      "folders": ["承租人"],
      "tag": "#承租人法定代表人身份证#"
     ,"flag":True
    },
    "法定代表人身份证2": {
      "keywords": ["法人身份证","法定代表人身份证表"],
      "folders": ["担保人"],
      "tag": "#担保人法定代表人身份证#"
     ,"flag":True

    },
    "企查查企业信用报告1": {
      "keywords": ["企业信用报告","企查查企业信用报告","【企查查】企业信用报告"],
      "folders": ["承租人"],
      "tag": "#承租人企查查企业信用报告#"
     ,"flag":True

    },
    "企查查企业信用报告2": {
      "keywords": ["企业信用报告","企查查企业信用报告","【企查查】企业信用报告"],
      "folders": ["担保人"],
      "tag": "#担保人企查查企业信用报告#"
     ,"flag":True

    },
    "企查查司法案件1": {
      "keywords": ["司法案件","企查查司法案件","【企查查】司法案件"],
      "folders": ["承租人"],
      "tag": "#承租人企查查司法案件#"
     ,"flag":False

    },
    "企查查司法案件2": {
      "keywords": ["司法案件","企查查司法案件","【企查查】司法案件"],
      "folders": ["担保人"],
      "tag": "#担保人企查查司法案件#"
     ,"flag":False

    },
    "中登网融资租赁1": {
      "keywords": ["中登网","中登网融资租赁","按担保人查询登记信息列表"],
      "folders": ["承租人"],
      "tag": "#承租人中登网融资租赁#"
     ,"flag":True

    },
    "中登网融资租赁2": {  
      "keywords": ["中登网","中登网融资租赁","按担保人查询登记信息列表"],
      "folders": ["担保人"],
      "tag": "#担保人中登网融资租赁#"
     ,"flag":True

    },
    "基础材料1": {
      "keywords": ["基础材料","基础资料","基础情况"],
      "folders": ["承租人"],
      "tag": "#承租人基础材料#"
     ,"flag":False

    },
    "基础材料2": {
      "keywords": ["基础材料","基础资料","基础情况"],
      "folders": ["担保人"],
      "tag": "#担保人基础材料#"
     ,"flag":False
    },

    "项目尽调报告": {
      "keywords": ["项目调查报告","项目尽调报告","尽调报告"],
      "folders": ["全套资料公司"],
      "tag": "#项目尽调报告#"
     ,"flag":True
    },
    
    "地方国企类业务指引": {
      "keywords": ["地方国企类","业务指引"],
      "folders": ["全套资料公司"],
      "tag": "#地方国企类业务指引#"
     ,"flag":True
    },

    "投放数据台账": {
      "keywords": ["投放数据台账"],
      "folders": ["全套资料公司"],
      "tag": "#投放数据台账#"
     ,"flag":True
    },
    "土地成交统计_按企业类型": {
      "keywords": ["土地成交统计_按企业类型"],
      "folders": ["全套资料公司"],
      "tag": "#土地成交统计_按企业类型#"
     ,"flag":True
    },
    "土地出让明细": {
      "keywords": ["土地出让明细"],
      "folders": ["全套资料公司"],
      "tag": "#土地出让明细#"
     ,"flag":True
    }
  }
}

def get_app_path(relative_path=""):
    """获取应用程序运行目录的路径（用于data、output等用户数据目录）
    
    开发环境返回当前目录；打包后返回exe所在目录
    """
    if getattr(sys, 'frozen', False):
        # 如果是打包后的exe
        base_path = os.path.dirname(sys.executable)
    else:
        # 如果是开发环境
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)


def ensure_dir(path):
    """确保目录存在，不存在则创建"""
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)
    return path

class UniversalFileRenamer:
    def __init__(self, template_dict):
        """
        初始化重命名器
        :param template_dict: 模板字典 (例如 dict_t)
        """
        self.current_template = template_dict.get("name", "未命名模板")
        self.templates = {self.current_template: template_dict}
        self.file_rules = template_dict.get("rules", {})
        
        # 默认支持的后缀，增加xlsx和png，且统一转为小写以支持大小写不区分
        default_extensions = [".pdf", ".doc", ".docx", ".xlsx", ".png", ".jpg", ".jpeg"]
        raw_extensions = template_dict.get("supported_extensions", default_extensions)
        self.supported_extensions = [ext.lower() for ext in raw_extensions]

    def get_available_templates(self):
        """获取所有可用的模板列表"""
        return list(self.templates.keys())
    
    def get_template_info(self, template_name):
        """获取指定模板的详细信息"""
        if template_name in self.templates:
            return self.templates[template_name]
        return None
    
    def switch_template(self, template_name):
        """切换到指定的模板（在此脚本简化版中主要保留接口兼容性）"""
        if template_name in self.templates:
            self.current_template = template_name
            template_data = self.templates[template_name]
            self.file_rules = template_data["rules"]
            
            # 更新支持的后缀
            default_extensions = [".pdf", ".doc", ".docx", ".xlsx", ".png", ".jpg", ".jpeg"]
            raw_extensions = template_data.get("supported_extensions", default_extensions)
            self.supported_extensions = [ext.lower() for ext in raw_extensions]
            return True
        return False
    
    def display_templates(self):
        """显示当前使用的模板信息"""
        print("当前使用的文件识别模板:")
        print("=" * 80)

        template_info = self.templates[self.current_template]
        print(f"1. {template_info['name']} [当前使用]")
        print(f"   描述: {template_info['description']}")
        print(f"   文件类型数: {len(template_info['rules'])}")

        # 显示该模板支持的文件扩展名
        default_exts = ['.pdf', '.doc', '.docx', '.xlsx', '.png', '.jpg', '.jpeg']
        extensions = template_info.get('supported_extensions', default_exts)
        print(f"   支持格式: {', '.join(extensions[:5])}")
        if len(extensions) > 5:
            print(f"                等{len(extensions)}种格式")
        else:
            print()

        # 显示该模板的文件类型
        file_types = list(template_info['rules'].keys())
        print(f"   包含类型: {', '.join(file_types[:3])}")
        if len(file_types) > 3:
            print(f"                {'等' + str(len(file_types)) + '种文件类型'}")
        print()
        
        return list(self.templates.keys())

    def find_target_files(self, base_folder):
        """
        在材料包文件夹中查找目标文件
        先递归遍历找到模板指定的文件夹，再在这些文件夹中查找文件
        """
        found_files = {}
        
        for file_type, rules in self.file_rules.items():
            found_files[file_type] = []
        
        # 递归遍历所有子文件夹，找到模板指定的文件夹
        for root, dirs, files in os.walk(base_folder):
            current_folder_name = os.path.basename(root)
            # 计算相对于 base_folder 的路径，用于检查层级
            rel_path = os.path.relpath(root, base_folder)
            if rel_path == '.':
                path_parts = []
            else:
                path_parts = rel_path.split(os.sep)
            
            # 检查当前文件夹是否匹配任何规则中的指定文件夹
            for file_type, rules in self.file_rules.items():
                target_folders = rules.get("folders", [])
                
                # 检查当前文件夹名是否匹配目标文件夹
                # 如果包含 "全套资料公司" 或为空，则匹配所有文件夹
                folder_matched = False
                if not target_folders or "全套资料公司" in target_folders:
                    folder_matched = True
                else:
                    for target_folder in target_folders:
                        if not target_folder or target_folder == ".":
                            continue
                        
                        # 支持递归匹配：检查路径中的每一层目录是否包含目标名称
                        # 例如：target="承租人"，路径=".../承租人/审计报告"，应该匹配
                        for part in path_parts:
                            if target_folder in part:
                                folder_matched = True
                                break
                        if folder_matched:
                            break
                
                if folder_matched:
                    # 在匹配的文件夹中查找符合关键词的文件
                    # 传入相对路径作为 folder_name，以便在结果中显示完整层级
                    folder_display = rel_path if rel_path != '.' else "根目录"
                    self._search_files_in_folder(
                        root, folder_display, file_type, rules, found_files
                    )
        
        # 同时在根目录查找
        for file_type, rules in self.file_rules.items():
            target_folders = rules.get("folders", [])
            if not target_folders or "全套资料公司" in target_folders or "" in target_folders or "." in target_folders:
                self._search_files_in_folder(
                    base_folder, "根目录", file_type, rules, found_files
                )
        
        return found_files
    
    def _search_files_in_folder(self, folder_path, folder_name, file_type, rules, found_files):
        """
        在指定文件夹中搜索符合规则的文件
        """
        try:
            for file in os.listdir(folder_path):
                file_path = os.path.join(folder_path, file)
                
                # 处理模板中支持的文件类型
                if (file.lower().endswith(tuple(self.supported_extensions)) and
                    os.path.isfile(file_path)):
                    
                    # 检查是否已经添加了标签（任何标签）
                    if '#' in file:
                        continue  # 跳过已经有标签的文件
                    
                    # 检查文件名是否包含关键词
                    file_matched = False
                    for keyword in rules["keywords"]:
                        if keyword in file:
                            file_matched = True
                            break
                    
                    if file_matched:
                        # 检查是否已经添加过相同文件（避免重复）
                        already_added = any(
                            existing['path'] == file_path 
                            for existing in found_files[file_type]
                        )
                        
                        if not already_added:
                            relative_path = os.path.join(folder_name, file) if folder_name else file
                            found_files[file_type].append({
                                'path': file_path,
                                'folder': folder_name or "根目录",
                                'filename': file,
                                'relative_path': relative_path
                            })
        except Exception as e:
            # 忽略文件夹访问错误
            pass
    
    def generate_new_name(self, file_info, file_type):
        """
        生成新的文件名
        """
        original_name = file_info['filename']
        
        # 为所有文件添加标签
        name_without_ext, extension = os.path.splitext(original_name)
        tag = self.file_rules[file_type]["tag"]
        
        return f"{name_without_ext}{tag}{extension}"
    
    def rename_files(self, base_folder):
        """
        执行文件重命名
        """
        print(f"[Search] 正在分析文件夹: {os.path.basename(base_folder)}")
        
        # 查找目标文件
        found_files = self.find_target_files(base_folder)
        
        # 显示查找结果
        total_files = 0
        for file_type, files in found_files.items():
            if files:
                print(f"\n[Folder] {file_type}:")
                for file_info in files:
                    print(f"  [File] {file_info['relative_path']}")
                    total_files += 1
            else:
                print(f"\n[FAIL] 未找到: {file_type}")
        
        if total_files == 0:
            print("\n[WARN] 没有找到需要重命名的文件")
            return False
            
        print(f"\n[Stats] 总共找到 {total_files} 个文件需要重命名")
        
        # 确认并执行重命名
        print(f"\n[START] 开始重命名操作...")
        renamed_count = 0
        failed_count = 0
        skipped_count = 0
        
        for file_type, files in found_files.items():
            if not files:
                continue
                
            print(f"\n[Folder] 处理 {file_type}:")
            for file_info in files:
                original_path = file_info['path']
                new_filename = self.generate_new_name(file_info, file_type)
                new_path = os.path.join(os.path.dirname(original_path), new_filename)
                
                print(f"  [File] {file_info['filename']}")
                
                try:
                    # 检查目标文件名是否已存在
                    if os.path.exists(new_path):
                        print(f"     [SKIP] 目标文件已存在，跳过")
                        skipped_count += 1
                        continue
                        
                    # 执行重命名
                    os.rename(original_path, new_path)
                    print(f"     [OK] 重命名成功")
                    renamed_count += 1
                except Exception as e:
                    print(f"     [FAIL] 重命名失败: {e}")
                    failed_count += 1
        
        # 显示结果统计
        print("[Stats] 重命名操作完成！统计结果:")
        print(f"  [File] 目标文件数: {total_files}")
        print(f"  [OK] 成功重命名: {renamed_count}")
        print(f"  [FAIL] 重命名失败: {failed_count}")
        print(f"  [SKIP] 跳过文件: {skipped_count}")
        
        return renamed_count > 0

def validate_folder_path(folder_path):
    """验证文件夹路径是否有效"""
    if not os.path.exists(folder_path):
        print(f"[FAIL] 文件夹不存在: {folder_path}")
        return False
    if not os.path.isdir(folder_path):
        print(f"[FAIL] 路径不是文件夹: {folder_path}")
        return False
    return True

def scan_data_folder():
    """扫描data文件夹中的材料包文件夹"""
    data_folder = get_app_path("data")
    if not os.path.exists(data_folder):
        os.makedirs(data_folder)
        return []
    
    material_packages = []
    print("[Search] 正在扫描data文件夹中的材料包文件夹...")
    
    try:
        for item in os.listdir(data_folder):
            item_path = os.path.join(data_folder, item)
            if os.path.isdir(item_path) and item.lower() != "output":
                material_packages.append(item_path)
    except Exception as e:
        print(f"[FAIL] 扫描过程中出错: {e}")

    if material_packages:
        print(f"[Stats] 找到 {len(material_packages)} 个材料包文件夹")
    return material_packages

def direct_batch_process(rules_dict):
    """
    直接从 data 读取材料包，处理并输出到 output 文件夹
    返回 missing_tags_report 字典
    """
    data_folder = get_app_path("data")
    output_folder = get_app_path("output")
    missing_tags_report = {} # 存储缺失标签信息的字典
    
    # 确保 output 文件夹存在
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"已创建 output 文件夹: {output_folder}")
    
    if not os.path.exists(data_folder):
        print(f"[FAIL] 数据文件夹不存在: {data_folder}\n将自动创立data文件夹")
        os.makedirs(data_folder)
        print(f"已创建数据文件夹: {data_folder}")
        return {}
    # 扫描材料包
    material_packages = scan_data_folder()
    
    if not material_packages:
        print("[FAIL] 未找到符合规范的材料包")
        return {}

    print(f"找到 {len(material_packages)} 个材料包，开始直接批处理...")
    
    # 使用传入的 rules_dict 初始化重命名器
    renamer = UniversalFileRenamer(rules_dict)

    success_count = 0
    total_files_renamed = 0

    for package_path in material_packages:
        package_name = os.path.basename(package_path)
        target_path = os.path.join(output_folder, package_name)
        
        print(f"\n[START] 正在分析: {package_name}")
        
        try:
            # 1. 先在原目录查找是否有匹配文件
            found_files = renamer.find_target_files(package_path)
            
            # --- 检查缺失标签 ---
            missing_in_package = []
            for rule_name, files in found_files.items():
                if not files:
                    # 获取该规则的详细信息
                    rule_info = rules_dict["rules"].get(rule_name)
                    if rule_info:
                        missing_in_package.append({
                            "rule": rule_name,
                            "tag": rule_info.get("tag", "无标签"),
                            "keywords": rule_info.get("keywords", []),
                            "flag": rule_info.get("flag", False)
                        })
            
            if missing_in_package:
                missing_tags_report[package_name] = missing_in_package
            # -------------------

            package_file_count = sum(len(files) for files in found_files.values())
            
            if package_file_count == 0:
                print(f"  [WARN] 未找到匹配规则的文件，跳过该材料包")
                continue

            print(f"  [DONE] 找到 {package_file_count} 个匹配文件，正在输出到 output...")

            # 2. 只有找到匹配文件才复制到 output (如果已存在则覆盖)
            if os.path.exists(target_path):
                shutil.rmtree(target_path)
            shutil.copytree(package_path, target_path)
            
            # 3. 在 output 中的副本上重新查找以执行重命名
            # 因为 found_files 里的路径是原路径，我们需要目标路径下的文件信息
            found_files_in_target = renamer.find_target_files(target_path)
            
            package_renamed_count = 0
            for file_type, files in found_files_in_target.items():
                for file_info in files:
                    original_path = file_info['path']
                    new_filename = renamer.generate_new_name(file_info, file_type)
                    new_path = os.path.join(os.path.dirname(original_path), new_filename)
                    
                    if not os.path.exists(new_path):
                        os.rename(original_path, new_path)
                        package_renamed_count += 1
            
            if package_renamed_count > 0:
                print(f"  [OK] 处理完成，重命名了 {package_renamed_count} 个文件")
                success_count += 1
                total_files_renamed += package_renamed_count
            else:
                print(f"  [WARN] 重命名过程中出现异常（未找到文件）")
                
        except Exception as e:
            print(f"  [FAIL] 处理失败 {package_name}: {e}")

    print("\n" + "=" * 50)
    print(f"[DONE] 批量处理结束！")
    print(f"[Folder] 输出目录: {output_folder}")
    print(f"[OK] 成功处理材料包: {success_count}/{len(material_packages)}")
    print(f"[File] 总重命名文件数: {total_files_renamed}")
    print("=" * 50)
    return missing_tags_report

class RenameApp:
    def __init__(self, root):
        self.root = root
        self.root.title("通用文件重命名工具")
        self.root.geometry("600x250")
        self.root.resizable(False, False)  # 固定窗口大小，禁止缩放
        self.root.configure(bg="#f8f9fa")  # 浅灰色背景
        
        # 设置字体
        font_family = "Microsoft YaHei" if sys.platform == "win32" else "Helvetica"
        log_font_family = "Consolas" if sys.platform == "win32" else "Menlo"

        self.title_font = font.Font(family=font_family, size=14, weight="bold")
        self.normal_font = font.Font(family=font_family, size=10)
        self.log_font = font.Font(family=log_font_family, size=10)
        
        # 配置样式
        self.style = ttk.Style()
        self.style.theme_use('clam')  # 使用 clam 主题以便更好地自定义
        
        # 按钮样式 - 扁平化且有轻微阴影效果
        self.style.configure("Primary.TButton",
                            font=self.normal_font,
                            padding=8,
                            background="#007bff",
                            foreground="white",
                            borderwidth=0)
        self.style.map("Primary.TButton",
                      background=[('active', '#0056b3'), ('disabled', '#cccccc')])
        
        self.style.configure("Secondary.TButton",
                            font=self.normal_font,
                            padding=8,
                            background="#6c757d",
                            foreground="white",
                            borderwidth=0)
        self.style.map("Secondary.TButton",
                      background=[('active', '#545b62'), ('disabled', '#cccccc')])

        self.style.configure("Success.TButton",
                            font=self.normal_font,
                            padding=8,
                            background="#28a745",
                            foreground="white",
                            borderwidth=0)
        self.style.map("Success.TButton",
                      background=[('active', '#218838'), ('disabled', '#cccccc')])

        self.style.configure("Danger.TButton",
                            font=self.normal_font,
                            padding=8,
                            background="#dc3545",
                            foreground="white",
                            borderwidth=0)
        self.style.map("Danger.TButton",
                      background=[('active', '#c82333'), ('disabled', '#cccccc')])

        # --- 界面布局 ---
        
        # 1. 顶部标题栏
        header = tk.Frame(root, bg="white", height=60, relief="flat")
        header.pack(fill=tk.X, side=tk.TOP)
        tk.Label(header, text="通用文件重命名工具", font=self.title_font, bg="white", fg="#333333").pack(pady=15)

        # 2. 中间操作区
        main_frame = tk.Frame(root, bg="#f8f9fa")
        main_frame.pack(expand=True, fill='both', padx=30, pady=20)
        
        # 按钮栏
        btn_container = tk.Frame(main_frame, bg="#f8f9fa")
        btn_container.pack(pady=10)
        
        self.btn_start = ttk.Button(btn_container, text="🚀 开始重命名", style="Success.TButton", command=self.start_process)
        self.btn_start.grid(row=0, column=0, padx=5, pady=5)
        
        self.btn_open_data = ttk.Button(btn_container, text="📁 输入(data)", style="Primary.TButton", command=self.open_data)
        self.btn_open_data.grid(row=0, column=1, padx=5, pady=5)
        
        self.btn_open_output = ttk.Button(btn_container, text="📂 输出(output)", style="Primary.TButton", command=self.open_output)
        self.btn_open_output.grid(row=0, column=2, padx=5, pady=5)
        
        self.btn_exit = ttk.Button(btn_container, text="🚪 退出", style="Danger.TButton", command=root.quit)
        self.btn_exit.grid(row=0, column=3, padx=5, pady=5)
        
        self.queue = queue.Queue()
        self.check_queue()
        
    def clear_logs(self):
        """已移除文本框，保留空函数以防其他地方调用（可选）"""
        pass

    def check_queue(self):
        """移除对 text_area 的操作，仅清空队列"""
        while not self.queue.empty():
            self.queue.get()
        self.root.after(100, self.check_queue)
        
    def start_process(self):
        # 预先检查 data 文件夹
        data_path = get_app_path("data")
        if not os.path.exists(data_path) or not any(os.path.isdir(os.path.join(data_path, i)) for i in os.listdir(data_path)):
             messagebox.showwarning("提示", "data文件夹为空，请先放入材料包文件夹")
             return

        self.btn_start.config(state='disabled')
        thread = threading.Thread(target=self.run_process)
        thread.daemon = True
        thread.start()
        
    def run_process(self):
        try:
            missing_report = direct_batch_process(dict_t)
            # 无论是否有缺失标签，都弹出提示
            self.root.after(0, lambda: self.show_popup(missing_report))
        except Exception as e:
            err_msg = str(e)
            self.root.after(0, lambda m=err_msg: messagebox.showerror("错误", f"处理过程中出错: {m}"))
        finally:
            self.root.after(0, lambda: self.btn_start.config(state='normal'))

    def show_popup(self, report):
        if not report:
            messagebox.showinfo("提示", "处理完成！未发现缺失标签。")
            return
            
        top = tk.Toplevel(self.root)
        top.title("分析报告")
        top.geometry("700x500")
        top.configure(bg="#ffffff")
        
        # 顶部说明
        header = tk.Frame(top, bg="#fff3cd", height=50)
        header.pack(fill=tk.X)
        tk.Label(header, text="处理完成！", font=("Microsoft YaHei", 10, "bold"), bg="#fff3cd", fg="#856404").pack(pady=10)

        tk.Label(header, text="检测到部分规则未匹配到文件，请检查以下内容：", font=("Microsoft YaHei", 10, "bold"), bg="#fff3cd", fg="#856404").pack(pady=10)
        
        # 报告内容区
        container = tk.Frame(top, bg="white", padx=20, pady=20)
        container.pack(expand=True, fill='both')
        
        txt = scrolledtext.ScrolledText(
            container, 
            font=self.log_font,
            bg="#fdfdfe",
            fg="#2c3e50",
            borderwidth=0,
            highlightthickness=1,
            highlightbackground="#dee2e6",
            padx=10,
            pady=10
        )
        txt.pack(expand=True, fill='both')
        
        content = ""
        for package, missing in report.items():
            content += f"📦 材料包: {package}\n"
            content += "═" * 60 + "\n"
            if not missing:
                 content += "  ✅ 所有标签均已匹配\n"
            else:
                for i, item in enumerate(missing, 1):
                    content += "\n"
                    content += f"  {i}. [缺失标签] {item['tag']}\n"
                    content += f"     [搜索关键词] {', '.join(item['keywords'])}\n"
                    if item.get('flag', False):
                        content += "     **此项是必须项 请检查是否存在**\n"
                    else:
                        content += "     --此项非必须项--\n"
                    content += "\n"
                    content += "  " + "-" * 50 + "\n"

            content += "\n"
            
        txt.insert(tk.END, content)
        txt.configure(state='disabled') # 设置为不可编辑但可选中
        
        # 底部关闭按钮
        btn_close = ttk.Button(top, text="确定并关闭", style="Secondary.TButton", command=top.destroy)
        btn_close.pack(pady=15)

    def open_data(self):
        open_folder(get_app_path("data"))
        
    def open_output(self):
        open_folder(get_app_path("output"))

class ThreadSafeWriter:
    def __init__(self, queue):
        self.queue = queue
    def write(self, msg):
        self.queue.put(msg)
    def flush(self):
        pass

def open_folder(path):
    """跨平台打开文件夹"""
    ensure_dir(path)
    print(f"[System] 尝试打开文件夹: {path}")
    try:
        if sys.platform == 'darwin':       # macOS
            subprocess.Popen(['open', path])
        elif sys.platform == 'win32':      # Windows
            os.startfile(path)
        else:                              # Linux
            subprocess.Popen(['xdg-open', path])
    except Exception as e:
        print(f"[Error] 打开文件夹失败: {e}")
        messagebox.showerror("错误", f"无法打开文件夹: {e}")

# 在 RenameApp 类中修改按钮回调
# def open_data(self):
#     open_folder(get_app_path("data"))

# def open_output(self):
#     open_folder(get_app_path("output"))

def hide_console():
    """在 Windows 上运行 GUI 时隐藏控制台窗口"""
    # 实际上，只要打包时用了 --windowed 参数，这个函数在 Mac/Windows 都不是必须的
    if sys.platform == 'win32':
        try:
            import ctypes
            whnd = ctypes.windll.kernel32.GetConsoleWindow()
            if whnd != 0:
                ctypes.windll.user32.ShowWindow(whnd, 0)
        except:
            pass
def main():
    # Windows High DPI awareness
    if sys.platform == 'win32':
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass

    hide_console()
    root = tk.Tk()
    app = RenameApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
