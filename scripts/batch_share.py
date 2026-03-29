#!/usr/bin/env python3
"""
百度网盘批量转存+分享工具
基于 BaiduPanFilesTransfers 项目逻辑实现
"""

import argparse
import os
import re
import sys
import time
from pathlib import Path
from typing import Union, List, Dict, Any, Optional

try:
    import requests
except ImportError:
    print("错误: 请先安装 requests 库: pip install requests")
    sys.exit(1)

try:
    import pandas as pd
    import openpyxl
except ImportError:
    print("错误: 请先安装依赖: pip install pandas openpyxl")
    sys.exit(1)


# ============ 常量配置 ============
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Referer': 'https://pan.baidu.com/',
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
}

BASE_URL = 'https://pan.baidu.com'

ERROR_CODES = {
    0: '成功',
    -1: '解析失败',
    2: '文件不存在',
    4: '链接已转存过',
    12: '目录已存在',
    18: '转存文件数超过限制',
    105: 'Cookie 失效',
    110: 'Cookie 失效',
    -12: '目录创建失败',
}

DELAY_SECONDS = 1


# ============ 工具函数 ============
def load_cookie_from_env() -> Optional[str]:
    """从 .env 文件加载 Cookie"""
    env_path = Path(__file__).parent.parent / ".env"
    if not env_path.exists():
        return None
    
    with open(env_path, 'r') as f:
        for line in f:
            line = line.strip()
            if line.startswith('BAIDU_COOKIE='):
                return line.split('=', 1)[1].strip()
    return None


def normalize_link(url_code: str) -> str:
    """
    预处理链接至标准格式。
    支持格式：
    - https://pan.baidu.com/s/1xxxxxx
    - https://pan.baidu.com/s/1xxxxxx?pwd=xxxx
    - https://pan.baidu.com/s/1xxxxxx 提取码：xxxx
    - https://pan.baidu.com/share/init?surl=xxxxx
    """
    normalized = url_code.replace("share/init?surl=", "s/1")
    normalized = re.sub(r'[?&]pwd=', ' ', normalized)
    normalized = re.sub(r'提取码*[:：]', ' ', normalized)
    normalized = re.sub(r'^.*?(https?://)', 'https://', normalized)
    normalized = re.sub(r'\s+', ' ', normalized)
    return normalized


def parse_url_and_code(url_code: str) -> tuple:
    """以空格分割出 URL 和提取码"""
    parts = url_code.strip().split()
    url = parts[0] if parts else ''
    code = parts[-1][-4:] if len(parts) > 1 else ''
    return url[:47] if url else '', code


def generate_code() -> str:
    """生成4位随机提取码"""
    import random
    import string
    characters = string.ascii_letters + string.digits
    return ''.join(random.choice(characters) for _ in range(4))


def update_cookie(bdclnd: str, cookie: str) -> str:
    """更新 Cookie 中的 BDCLND 值"""
    cookies_dict = dict(map(lambda item: item.split('=', 1), filter(None, cookie.split(';'))))
    cookies_dict['BDCLND'] = bdclnd
    return ';'.join([f'{key}={value}' for key, value in cookies_dict.items()])


# ============ 网络请求类 ============
class BaiduNetwork:
    """百度网盘网络请求类"""
    
    def __init__(self, cookie: str, trust_env: bool = True):
        self.session = requests.Session()
        self.headers = HEADERS.copy()
        self.headers['Cookie'] = cookie
        self.session.trust_env = trust_env
        self.bdstoken = ''
        # 忽略证书验证警告
        requests.packages.urllib3.disable_warnings()
    
    def _retry_request(self, func, *args, max_retries=3, **kwargs):
        """带重试的请求"""
        for attempt in range(max_retries):
            try:
                result = func(*args, **kwargs)
                if result is not None:
                    return result
            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(1)
                else:
                    raise e
    
    def get_bdstoken(self) -> Union[str, int]:
        """获取 bdstoken"""
        url = f'{BASE_URL}/api/gettemplatevariable'
        params = {
            'clienttype': '0',
            'app_id': '38824127',
            'web': '1',
            'fields': '["bdstoken","token","uk","isdocuser","servertime"]'
        }
        
        r = self.session.get(url=url, params=params, headers=self.headers, timeout=10, allow_redirects=False, verify=False)
        data = r.json()
        if data.get('errno') != 0:
            return data.get('errno', -1)
        
        return data.get('result', {}).get('bdstoken', '')
    
    def verify_pass_code(self, link_url: str, pass_code: str) -> Union[str, int]:
        """验证提取码"""
        url = f'{BASE_URL}/share/verify'
        surl = link_url[25:48] if len(link_url) > 48 else ''
        params = {
            'surl': surl,
            'bdstoken': self.bdstoken,
            't': str(int(time.time() * 1000)),
            'channel': 'chunlei',
            'web': '1',
            'clienttype': '0'
        }
        data = {'pwd': pass_code, 'vcode': '', 'vcode_str': ''}
        
        r = self.session.post(url=url, params=params, headers=self.headers, data=data, timeout=10, allow_redirects=False, verify=False)
        data = r.json()
        if data.get('errno') != 0:
            return data.get('errno', -1)
        
        return data.get('randsk', '')
    
    def get_transfer_params(self, url: str) -> str:
        """获取转存参数"""
        r = self.session.get(url=url, headers=self.headers, timeout=15, verify=False)
        return r.content.decode("utf-8")
    
    def transfer_file(self, params_list: List[str], folder_name: str) -> int:
        """转存文件"""
        url = f'{BASE_URL}/share/transfer'
        params = {
            'shareid': params_list[0],
            'from': params_list[1],
            'bdstoken': self.bdstoken,
            'channel': 'chunlei',
            'web': '1',
            'clienttype': '0'
        }
        data = {
            'fsidlist': f"[{','.join(params_list[2])}]",
            'path': f'/{folder_name}'
        }
        
        r = self.session.post(url=url, params=params, headers=self.headers, data=data, timeout=30, allow_redirects=False, verify=False)
        return r.json().get('errno', -1)
    
    def create_share(self, fs_id: int, expiry: int, password: str) -> Union[str, int]:
        """创建分享链接"""
        url = f'{BASE_URL}/share/set'
        params = {
            'channel': 'chunlei',
            'bdstoken': self.bdstoken,
            'clienttype': '0',
            'app_id': '250528',
            'web': '1'
        }
        data = {
            'period': str(expiry),
            'pwd': password,
            'eflag_disable': 'true',
            'channel_list': '[]',
            'schannel': '4',
            'fid_list': f'[{fs_id}]'
        }
        
        r = self.session.post(url=url, params=params, headers=self.headers, data=data, timeout=15, allow_redirects=False, verify=False)
        data = r.json()
        if data.get('errno') != 0:
            return data.get('errno', -1)
        
        return data.get('link', '')
    
    def get_file_list(self, folder_name: str = '/') -> Union[List, int]:
        """获取目录文件列表"""
        url = f'{BASE_URL}/api/list'
        params = {
            'order': 'time',
            'desc': '1',
            'showempty': '0',
            'web': '1',
            'page': '1',
            'num': '1000',
            'dir': folder_name,
            'bdstoken': self.bdstoken
        }
        
        r = self.session.get(url=url, params=params, headers=self.headers, timeout=15, allow_redirects=False, verify=False)
        data = r.json()
        if data.get('errno') != 0:
            return data.get('errno', -1)
        
        return data.get('list', [])
    
    def create_dir(self, folder_name: str) -> int:
        """创建目录"""
        url = f'{BASE_URL}/api/create'
        params = {
            'a': 'commit',
            'bdstoken': self.bdstoken
        }
        data = {
            'path': folder_name,
            'isdir': '1',
            'block_list': '[]',
        }
        
        r = self.session.post(url=url, params=params, headers=self.headers, data=data, timeout=15, allow_redirects=False, verify=False)
        return r.json().get('errno', -1)


def parse_response_content(response: str) -> tuple:
    """解析页面源码提取转存参数"""
    shareid_list = re.findall(r'"shareid":(\d+?),"', response)
    user_id_list = re.findall(r'"share_uk":"(\d+?)","', response)
    fs_id_list = re.findall(r'"fs_id":(\d+?),"', response)
    server_filename_list = re.findall(r'"server_filename":"(.+?)","', response)
    isdir_list = re.findall(r'"isdir":(\d+?),"', response)
    
    if not all([shareid_list, user_id_list, fs_id_list, server_filename_list]):
        return None
    
    return shareid_list[0], user_id_list[0], fs_id_list, server_filename_list[0], isdir_list[0] if isdir_list else '0'


# ============ 输入解析 ============
def parse_input(input_data: str) -> List[Dict[str, str]]:
    """解析输入数据（自动去重）
    
    支持格式：
    1. 只有链接：https://pan.baidu.com/s/xxxxx
    2. 名称+链接：XXX教程 https://pan.baidu.com/s/xxxxx
    """
    links = []
    seen_urls = set()
    MAX_URL_LENGTH = 200
    
    # 判断是文件路径还是直接文本
    if Path(input_data).exists():
        file_path = Path(input_data).resolve()
        
        # 安全检查：防止路径遍历
        allowed_dirs = [Path.cwd(), Path.home()]
        if not any(str(file_path).startswith(str(d)) for d in allowed_dirs):
            print(f"错误: 文件路径不允许超出工作目录")
            return []
        
        # 检查文件大小（最大 50MB）
        if file_path.stat().st_size > 50 * 1024 * 1024:
            print(f"错误: 文件过大（最大 50MB）")
            return []
        
        if file_path.suffix in ['.xlsx', '.xls']:
            # Excel 文件
            df = pd.read_excel(file_path)
            
            # 查找链接列（支持多种列名）
            link_col = None
            for col in df.columns:
                if '链接' in col or '链接' in col.lower() or 'url' in col.lower() or 'link' in col.lower():
                    link_col = col
                    break
            
            if not link_col:
                print("错误: Excel 文件中缺少'链接'列")
                return []
            
            # 查找名称列
            name_col = None
            for col in df.columns:
                if '名称' in col or '名字' in col or 'name' in col.lower():
                    name_col = col
                    break
            
            for _, row in df.iterrows():
                name = str(row.get(name_col, '')).strip() if name_col else ''
                link = str(row.get(link_col, '')).strip()
                if link and link not in seen_urls and 'pan.baidu.com' in link:
                    links.append({'name': name, 'url': link})
                    seen_urls.add(link)
        else:
            # 文本文件
            with open(file_path, 'r', encoding='utf-8') as f:
                for line in f:
                    parsed = _parse_line(line.strip(), seen_urls)
                    if parsed:
                        links.append(parsed)
    else:
        # 直接文本
        for line in input_data.split('\n'):
            parsed = _parse_line(line.strip(), seen_urls)
            if parsed:
                links.append(parsed)
    
    return links


def _parse_line(line: str, seen_urls: set) -> Optional[Dict[str, str]]:
    """解析单行输入"""
    if not line or line.startswith('#'):
        return None
    
    # 提取百度网盘链接
    url_match = re.search(r'https?://pan\.baidu\.com/s/[a-zA-Z0-9_-]+', line)
    
    if not url_match:
        return None
    
    url = url_match.group(0)
    if len(url) > 200:
        return None
    
    # 提取链接前的文本作为名称
    name_part = line[:url_match.start()].strip()
    name_part = re.sub(r'[-:：\s\t]+$', '', name_part)
    name_part = name_part[:200] if name_part else ''
    
    # 提取提取码（如果有）
    pwd_match = re.search(r'(?:提取码|密码|pwd)[:：]?\s*([a-zA-Z0-9]{4})', line, re.IGNORECASE)
    password = pwd_match.group(1) if pwd_match else ''
    
    if url not in seen_urls:
        seen_urls.add(url)
        return {'name': name_part, 'url': url, 'password': password}
    
    return None


# ============ 核心处理 ============
class BaiduProcessor:
    """百度网盘批量处理器"""
    
    def __init__(self, cookie: str, folder_name: str = "来自：分享", trust_env: bool = True):
        self.network = BaiduNetwork(cookie, trust_env)
        self.folder_name = folder_name
        self.results = []
        self.processed_urls = set()  # 已转存的链接记录
        self.existing_files = set()  # 目标目录中已存在的文件名
    
    def initialize(self) -> bool:
        """初始化：获取 bdstoken 和创建目录"""
        print("初始化百度网盘客户端...")
        self.network.bdstoken = self.network.get_bdstoken()
        
        if isinstance(self.network.bdstoken, int):
            print(f"获取 bdstoken 失败，错误代码：{self.network.bdstoken}")
            return False
        
        print(f"✓ 登录成功")
        
        # 创建目标目录（固定：来自：分享）
        if self.folder_name:
            result = self.network.get_file_list(f'/{self.folder_name}')
            if isinstance(result, int):
                return_code = self.network.create_dir(self.folder_name)
                if return_code == 0:
                    print(f"✓ 创建目录: {self.folder_name}")
                else:
                    print(f"创建目录失败，错误代码: {return_code}")
            else:
                # 加载已存在的文件名，用于转存去重
                self.existing_files = {item.get('server_filename', '') for item in result if item.get('server_filename')}
                print(f"  目标目录已有 {len(self.existing_files)} 个文件")
        
        return True
    
    def verify_link(self, url: str, password: str = '') -> tuple:
        """验证链接有效性"""
        # 如果有提取码，先验证
        if password:
            bdclnd = self.network.verify_pass_code(url, password)
            if isinstance(bdclnd, int):
                return None, f"提取码错误({bdclnd})"
            
            # 更新 cookie
            self.network.headers['Cookie'] = update_cookie(bdclnd, self.network.headers['Cookie'])
        
        # 获取转存参数
        response = self.network.get_transfer_params(url)
        params = parse_response_content(response)
        
        if not params:
            return None, "解析失败"
        
        return params, ""
    
    def process_save(self, url: str, password: str = '', custom_folder: str = '') -> Dict[str, Any]:
        """转存单个链接"""
        result = {
            'name': '',
            'new_share_url': '',
            'status': '等待中',
            'error': ''
        }
        
        try:
            # ========== 去重检查 ==========
            # 1. 检查是否已转存过（通过 URL 去重）
            if url in self.processed_urls:
                result['status'] = '跳过'
                result['error'] = '链接已处理过'
                return result
            
            # 验证链接
            params, error = self.verify_link(url, password)
            if not params:
                result['status'] = '失败'
                result['error'] = error or '验证失败'
                return result
            
            shareid, share_uk, fs_ids, file_name, isdir = params
            result['name'] = file_name
            
            # 2. 检查目标目录是否已存在同名文件（文件去重）
            if file_name in self.existing_files:
                result['status'] = '跳过'
                result['error'] = f'文件已存在'
                return result
            
            # 确定目标目录
            target_folder = self.folder_name
            if custom_folder:
                target_folder = f"{self.folder_name}/{custom_folder}" if self.folder_name else custom_folder
            
            # 执行转存
            transfer_result = self.network.transfer_file([shareid, share_uk, fs_ids], target_folder)
            
            if transfer_result != 0:
                error_msg = ERROR_CODES.get(transfer_result, f'转存失败({transfer_result})')
                result['status'] = '失败'
                result['error'] = error_msg
                return result
            
            # 转存成功，记录文件名
            self.existing_files.add(file_name)
            self.processed_urls.add(url)
            
            # ========== 创建分享（永久有效 + 随机提取码）==========
            fs_id = int(fs_ids[0]) if fs_ids else 0
            random_pwd = generate_code()  # 生成随机4位提取码
            # expiry = -1 表示永久有效
            share_url = self.network.create_share(fs_id, -1, random_pwd)
            
            if isinstance(share_url, str) and share_url:
                # 自动填充提取码
                result['new_share_url'] = f"{share_url}?pwd={random_pwd}"
                result['status'] = '成功'
            else:
                result['status'] = '转存成功'
                result['error'] = '分享创建失败'
        
        except Exception as e:
            result['status'] = '失败'
            result['error'] = str(e)
        
        return result
    
    def batch_process(self, links: List[Dict], delay: int = 1, retry: int = 2) -> List[Dict]:
        """批量处理链接"""
        results = []
        
        for i, link in enumerate(links, 1):
            url = link['url']
            name = link.get('name', '')
            password = link.get('password', '')
            
            print(f"[{i}/{len(links)}] 处理: {name or url[:50]}...")
            
            # 重试逻辑
            for attempt in range(retry):
                result = self.process_save(url, password)
                
                if result['status'] == '成功' or '转存成功' in result['status']:
                    break
                
                if attempt < retry - 1:
                    print(f"  第 {attempt + 1} 次失败，重试...")
                    time.sleep(delay)
            
            if result['status'] == '成功':
                print(f"  ✓ 成功: {result['new_share_url']}")
            else:
                print(f"  ✗ 失败: {result['error']}")
            
            results.append(result)
            
            # 处理间隔
            if i < len(links):
                time.sleep(delay)
        
        return results


# ============ 结果保存 ============
def save_results(results: List[Dict], output_path: str):
    """保存结果到 Excel"""
    output_file = Path(output_path).resolve()
    
    # 安全检查
    allowed_dirs = [Path.cwd(), Path.home()]
    if not any(str(output_file).startswith(str(d)) for d in allowed_dirs):
        print(f"错误: 输出路径不允许超出工作目录")
        return
    
    output_file.parent.mkdir(parents=True, exist_ok=True)
    
    # 整理数据
    data = []
    for r in results:
        status = r.get('status', '')
        link = r.get('new_share_url', '') if '成功' in status else ''
        error = r.get('error', '')
        
        data.append({
            '资源名称': r.get('name', ''),
            '网盘链接': link,
            '状态': status,
            '备注': error if error and '成功' not in status else ''
        })
    
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False, engine='openpyxl')
    
    # 统计
    success = len([r for r in results if '成功' in r.get('status', '')])
    failed = len(results) - success
    
    print(f"\n处理完成!")
    print(f"成功: {success} 个, 失败: {failed} 个")
    print(f"结果已保存到: {output_file}")


# ============ 主函数 ============
def main():
    parser = argparse.ArgumentParser(description="百度网盘批量转存+分享工具")
    parser.add_argument("--input", required=True, help="输入文件路径或链接文本")
    parser.add_argument("--output", default="outputs/tables/result.xlsx", help="输出Excel文件路径")
    parser.add_argument("--folder", default="来自：分享", help="转存目标文件夹名称（默认：来自：分享）")
    parser.add_argument("--cookie", help="百度网盘 Cookie（可选）")
    parser.add_argument("--delay", type=int, default=1, help="请求间隔秒数（默认1秒）")
    parser.add_argument("--retry", type=int, default=2, help="失败重试次数（默认2次）")
    parser.add_argument("--no-proxy", action="store_true", help="禁用系统代理")
    
    args = parser.parse_args()
    
    # 获取 Cookie
    cookie = args.cookie or load_cookie_from_env()
    if not cookie:
        print("错误: 请通过 --cookie 参数提供 Cookie 或创建 .env 文件")
        print("  .env 文件应包含: BAIDU_COOKIE=你的Cookie值")
        sys.exit(1)
    
    # 初始化处理器
    processor = BaiduProcessor(
        cookie=cookie,
        folder_name=args.folder,
        trust_env=not args.no_proxy
    )
    
    # 初始化连接
    if not processor.initialize():
        print("初始化失败，请检查 Cookie 是否正确")
        sys.exit(1)
    
    # 解析输入
    print(f"解析输入: {args.input}")
    links = parse_input(args.input)
    print(f"找到 {len(links)} 个链接（已去重）")
    
    if not links:
        print("没有找到有效的百度网盘链接")
        sys.exit(1)
    
    # 处理
    results = processor.batch_process(links, args.delay, args.retry)
    
    # 保存结果
    save_results(results, args.output)


if __name__ == "__main__":
    main()