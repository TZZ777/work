import fastapi
import json
import os
import re
import unicodedata
import uvicorn
from typing import List, Optional
from urllib.parse import urlparse, unquote
from pydantic import BaseModel

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
    "融资情况": {"keywords": ["融资情况"], "folders": ["全套资料公司"], "tag": "#融资情况#", "flag":True},
    "最新年份审计报告注释1": {"keywords": ["附注","注释"], "folders": ["承租人"], "tag": "#最新年份审计报告注释#", "flag":True},
    "最新年份审计报告注释2": {"keywords": ["附注","注释"], "folders": ["担保人"], "tag": "#最新年份审计报告注释#", "flag":True}
  }
}
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

def normalize_text(text):
    if text is None:
        return ""
    normalized = unicodedata.normalize("NFKC", str(text)).lower()
    return re.sub(r"[\s\-_()（）\[\]【】]", "", normalized)

def parse_local_path(url: str) -> Optional[str]:
    if not url:
        return None
    parsed = urlparse(url)
    if parsed.scheme in ("http", "https"):
        return None
    if parsed.scheme == "file":
        path = unquote(parsed.path)
        if os.name == "nt" and re.match(r"^/[a-zA-Z]:", path):
            path = path[1:]
        return os.path.abspath(path)
    if parsed.scheme == "" and parsed.netloc == "":
        return os.path.abspath(unquote(url))
    path = unquote(parsed.path) if parsed.path else url
    return os.path.abspath(path)

def resolve_existing_path(path: str) -> Optional[str]:
    candidates = []
    candidates.append(path)
    if path.startswith("\\") and "\\Users\\" in path:
        drive = os.path.splitdrive(os.getcwd())[0]
        if drive:
            candidates.append(drive + path)
    no_newlines = path.replace("\r", "").replace("\n", "").replace("\t", "")
    if no_newlines != path:
        candidates.append(no_newlines)
        if no_newlines.startswith("\\") and "\\Users\\" in no_newlines:
            drive = os.path.splitdrive(os.getcwd())[0]
            if drive:
                candidates.append(drive + no_newlines)
    no_whitespace = re.sub(r"\s+", "", path)
    if no_whitespace != path:
        candidates.append(no_whitespace)
        if no_whitespace.startswith("\\") and "\\Users\\" in no_whitespace:
            drive = os.path.splitdrive(os.getcwd())[0]
            if drive:
                candidates.append(drive + no_whitespace)
    for candidate in candidates:
        if candidate and os.path.exists(candidate) and os.path.isfile(candidate):
            return candidate
    return None

def match_rule_for_path(filename: str, path: str, rules: dict) -> Optional[str]:
    normalized_filename = normalize_text(filename)
    path_parts = os.path.dirname(path).split(os.sep) if path else []
    for file_type, rule in rules.items():
        target_folders = rule.get("folders", [])
        folder_matched = False
        if not target_folders or "全套资料公司" in target_folders:
            folder_matched = True
        else:
            for target_folder in target_folders:
                if any(target_folder in part for part in path_parts):
                    folder_matched = True
                    break
        if not folder_matched:
            continue
        keywords = rule.get("keywords", [])
        if any(normalize_text(keyword) in normalized_filename for keyword in keywords if keyword):
            return file_type
    return None

def process(url_list, map):
    renamer = UniversalFileRenamer(map)
    results = []
    for url in url_list:
        local_path = parse_local_path(url)
        if not local_path:
            results.append({
                "url": url,
                "status": "error",
                "message": "仅支持本地路径或file:// URL"
            })
            continue
        if not os.path.exists(local_path) or not os.path.isfile(local_path):
            resolved = resolve_existing_path(local_path)
            if resolved:
                local_path = resolved
            else:
                results.append({
                    "url": url,
                    "status": "error",
                    "message": "文件不存在",
                    "local_path": local_path
                })
                continue
        filename = os.path.basename(local_path)
        file_type = match_rule_for_path(filename, local_path, renamer.file_rules)
        if not file_type:
            results.append({
                "url": url,
                "status": "skipped",
                "message": "未命中规则",
                "filename": filename,
                "local_path": local_path
            })
            continue
        new_name = renamer.generate_new_name({"filename": filename}, file_type)
        new_path = os.path.join(os.path.dirname(local_path), new_name)
        if os.path.abspath(local_path) == os.path.abspath(new_path):
            results.append({
                "url": url,
                "status": "skipped",
                "message": "文件名无需变更",
                "filename": filename,
                "rule": file_type,
                "local_path": local_path
            })
            continue
        try:
            os.replace(local_path, new_path)
            results.append({
                "url": url,
                "status": "renamed",
                "old_name": filename,
                "new_name": new_name,
                "rule": file_type,
                "local_path": local_path,
                "new_path": new_path
            })
        except Exception as e:
            results.append({
                "url": url,
                "status": "error",
                "message": str(e),
                "filename": filename,
                "rule": file_type,
                "local_path": local_path
            })
    return results

def extract_processed_urls(results):
    processed = []
    for item in results:
        if not isinstance(item, dict):
            continue
        new_path = item.get("new_path")
        if new_path:
            processed.append(new_path)
            continue
        local_path = item.get("local_path")
        if local_path:
            processed.append(local_path)
            continue
        url = item.get("url")
        if url:
            processed.append(url)
    return processed

def extract_common_folder(processed_urls):
    if not processed_urls:
        return None
    dirs = [os.path.dirname(p) for p in processed_urls if p]
    if not dirs:
        return None
    common = os.path.commonpath(dirs)
    return common

app = fastapi.FastAPI()

class RenameRequest(BaseModel):
    url_list: List[str]

def parse_url_list(value):
    if value is None:
        return None
    if isinstance(value, list):
        return value
    if isinstance(value, str):
        try:
            parsed = json.loads(value)
            if isinstance(parsed, list):
                return parsed
        except Exception:
            items = re.findall(r'"((?:[^"\\]|\\.)*)"', value, flags=re.DOTALL)
            if items:
                results = []
                for item in items:
                    try:
                        decoded = json.loads(f"\"{item}\"")
                    except Exception:
                        decoded = item
                    decoded = decoded.replace("\r", "").replace("\n", "").strip()
                    if decoded:
                        results.append(decoded)
                return results if results else None
            parts = re.split(r"[\r\n,]+", value)
            cleaned = [p.strip().strip('"').strip("'") for p in parts if p.strip()]
            return cleaned if cleaned else None
    return None

@app.post("/jr/rename")
async def rename_files(request: fastapi.Request):
    url_list = None
    content_type = request.headers.get("content-type", "")
    if "application/json" in content_type:
        body = await request.json()
        if isinstance(body, dict):
            url_list = parse_url_list(body.get("url_list"))
    if url_list is None and ("multipart/form-data" in content_type or "application/x-www-form-urlencoded" in content_type):
        form = await request.form()
        url_list = parse_url_list(form.get("url_list"))
    if url_list is None:
        url_list = parse_url_list(request.query_params.get("url_list"))
    if url_list is None:
        return fastapi.responses.JSONResponse(
            status_code=400,
            content={"detail": "url_list 必须是 JSON 数组或可解析的 JSON 字符串"}
        )
    results = process(url_list, dict_t)
    processed_urls = extract_processed_urls(results)
    common_folder = extract_common_folder(processed_urls)
    return {"result": common_folder}

@app.get("/health")
def health():
    return {"status": "ok"}

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=9811)
