import os
import aiohttp
import getpass
import json
import hashlib
from Crypto.Cipher import DES
import base64
from bs4 import BeautifulSoup
import time
import asyncio
from rich import print as rprint
from rich.progress import Progress, BarColumn, TextColumn
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox

url = "https://i-learning.cycu.edu.tw/"

if sys.platform == 'win32':
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

# MD5 Encrypt
def md5_encode(input_string) -> str:
    md5_hash = hashlib.md5()
    md5_hash.update(input_string.encode('utf-8'))
    return md5_hash.hexdigest()

# DES Encrypt ECB NoPadding
def des_encode(key: str, data) -> str:
    cipher = DES.new(key.encode('utf-8'), DES.MODE_ECB)
    encrypted_data = cipher.encrypt(data.encode('utf-8'))
    return str(base64.encodebytes(encrypted_data), encoding='utf-8').replace("\n", "")

async def fetch_login_key(session):
    while True:
        async with session.get(url + "sys/door/re_gen_loginkey.php?xajax=reGenLoginKey", headers=headers) as response:
            res = await response.text()
            if "loginForm.login_key.value = \"" in res:
                return res.split("loginForm.login_key.value = \"")[1].split("\"")[0]

async def login(session, id, pwd, loginKey) -> bool:
    async with session.post(url + "login.php", headers=headers, data={
        "username": id,
        "pwd": pwd,
        "password": "*" * len(pwd),
        "login_key": loginKey,
        "encrypt_pwd": des_encode(md5_encode(pwd)[:4] + loginKey[:4], pwd + " " * (16 - len(pwd) % 16) if len(pwd) % 16 != 0 else pwd),
    }) as response:
        res = await response.text()
        if "lang=\"big5" in res:
            rprint("[red]登入失敗，請重新再試![/red]")
            return False
    rprint("[green]登入成功！[/green]")
    return True

async def fetch_courses(session):
    async with session.get(url + "learn/mooc_sysbar.php", headers=headers) as response:
        soup = BeautifulSoup(await response.text(), 'lxml')
        courses = {
            option["value"]: option.text.strip()
            for option in soup.select("select#selcourse option")
            if option["value"] != "10000000"
        }
        return courses

async def fetch_hrefs(session, course_id):
    async with session.get(url + f"xmlapi/index.php?action=my-course-path-info&cid={course_id}", headers=headers) as response:
        items = json.loads(await response.text())
        hrefs = dict()
        if items['code'] == 0:
            def search_hrefs(data):
                if isinstance(data, dict):
                    for key, value in data.items():
                        if key == 'href' and (value.endswith('.pdf') or value.endswith('.pptx') or value.endswith('.mp4')):
                            pattern = r'[<>:"/\\|?*\x00-\x1F\x7F]'
                            name = re.sub(pattern, '', str(data['text']))
                            hrefs[name] = str(value)
                        elif isinstance(value, (dict, list)):
                            search_hrefs(value)
                elif isinstance(data, list):
                    for item in data:
                        search_hrefs(item)
            search_hrefs(items['data']['path']['item'])
        return hrefs

async def download_material(session: aiohttp.ClientSession, href, filename, course_name, base_path):
    async with session.get(href, headers=headers) as response:
        if response.status != 200:
            return
        filename += f".{str(response.url).split('.')[-1]}"
        materials_path = os.path.join(base_path, "materials")
        save_path = os.path.join(materials_path, course_name)
        if not os.path.exists(save_path):
            os.makedirs(save_path)
        file_path = os.path.join(save_path, filename)
        if os.path.exists(file_path):
            return
        with open(file_path, 'wb') as file:
            async for chunk in response.content.iter_chunked(8192):
                if chunk:
                    file.write(chunk)

def select_folder():
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="選擇下載資料夾", initialdir=os.getcwd())
    root.destroy()
    default = os.path.join(os.getcwd())
    return folder if folder else default

def ask_download_more():
    root = tk.Tk()
    root.withdraw()
    result = messagebox.askyesno("下載完成", "課程資料已下載完成！\n是否要下載其他課程的資料？\n(按 '否' 將結束程式)")
    root.destroy()
    return "more" if result else "exit"

headers = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.0 Safari/605.1.15"}

async def download_courses(session, courses, base_path, progress):
    course_list = list(courses.items())
    rprint("\n[bold cyan]可用課程列表：[/bold cyan]")
    for i, (course_id, course_name) in enumerate(course_list, 1):
        rprint(f"[yellow]{i}.[/yellow] {course_name} (ID: {course_id})")
    
    while True:
        rprint("\n[cyan]請選擇要下載的課程：[/cyan]")
        selection = input("輸入要下載的課程編號 (用逗號分隔，例如 '1, 3, 5') 或輸入 'all' 下載全部，或輸入 'end' 結束程式：")
        selection = selection.lower().strip()
        
        if selection in ('end', 'exit'):
            rprint("[cyan]使用者選擇結束程式...[/cyan]")
            return False  # Signal to exit the program
        
        if selection == 'all':
            selected_courses = course_list
            break
        try:
            selected_indices = [int(x.strip()) - 1 for x in selection.split(',')]
            if all(0 <= idx < len(course_list) for idx in selected_indices):
                selected_courses = [course_list[idx] for idx in selected_indices]
                break
            else:
                rprint("[red]無效的編號，請輸入範圍內的數字！[/red]")
        except ValueError:
            rprint("[red]請輸入有效數字、'all' 或 'end'！[/red]")
    
    rprint(f"\n[green]即將下載 {len(selected_courses)} 門課程...[/green]")
    start = time.time()
    
    course_task = progress.add_task("[cyan]時間差不多囉...", total=len(selected_courses))
    with progress:
        for course_id, course_name in selected_courses:
            progress.update(course_task, description=f"[cyan] 正在下載 {course_name}")
            hrefs = await fetch_hrefs(session, course_id)
            tasks = [download_material(session, hrefs[filename], filename, course_name, base_path) for filename in hrefs.keys()]
            
            task_start = time.time()
            sub_task = progress.add_task("[orange3]", total=len(hrefs))
            for i, task in enumerate(tasks):
                progress.update(sub_task, description=f"[orange3] {list(hrefs.keys())[i][:50]} ...")
                await task
                progress.update(sub_task, advance=1)
            
            progress.remove_task(sub_task)
            progress.update(course_task, advance=1, description="")
            progress.console.print(f"[green] {course_name} 下載完成, 耗時: %.2fs" % (time.time() - task_start))

    progress.remove_task(course_task)
    rprint(f"[bold green]下載完成! 總耗時: %.2fs[/bold green]" % (time.time() - start))
    return True  # Signal that download completed normally

async def main():
    os.system("title CYCU-iLearning-Scraper")
    rprint("[yellow]<!!! 尊重版權/著作權 尊重版權/著作權 尊重版權/著作權 !!!>[/yellow]\n")
    
    # Load saved username
    config_file = "config.json"
    saved_id = None
    if os.path.exists(config_file):
        try:
            with open(config_file, 'r') as f:
                config = json.load(f)
                saved_id = config.get("username")
        except (json.JSONDecodeError, IOError):
            rprint("[red]無法載入設定檔，將重新輸入...[/red]")
    
    if saved_id:
        rprint(f"[green]已儲存的學號：{saved_id}[/green]")
        use_saved = input("使用此學號？(y/n)：").lower().startswith('y')
        id = saved_id if use_saved else input("輸入您的學號：")
    else:
        id = input("輸入您的學號：")
    
    if not saved_id or id != saved_id:
        with open(config_file, 'w') as f:
            json.dump({"username": id}, f)
        rprint(f"[green]學號 {id} 已儲存至 {config_file}[/green]")
    
    pwd = getpass.getpass("輸入您的itouch密碼：")
    
    # Folder selection
    rprint("\n[cyan]請選擇儲存課程資料的資料夾[/cyan]")
    base_path = select_folder()
    materials_path = os.path.join(base_path, "materials")
    rprint(f"[green]下載資料將儲存至：{materials_path}[/green]")
    
    resolver = aiohttp.AsyncResolver(nameservers=["1.1.1.1", "1.0.0.1"])
    connector = aiohttp.TCPConnector(limit=50, resolver=resolver)
    async with aiohttp.ClientSession(connector=connector) as session:
        login_key = await fetch_login_key(session)
        if not await login(session, id, pwd, login_key):
            input("\n按 Enter 鍵重試...")
            return
        
        progress = Progress(
            TextColumn("{task.description}", justify="left"),
            BarColumn(),
            TextColumn("{task.completed}/{task.total}", justify="right"),
        )
        
        while True:
            courses = await fetch_courses(session)
            continue_program = await download_courses(session, courses, base_path, progress)
            if not continue_program:
                rprint("[cyan]程式結束。[/cyan]")
                break
            
            choice = ask_download_more()
            if choice == "exit":
                rprint("[cyan]程式結束。[/cyan]")
                break
            else:
                rprint("[cyan]準備下載其他課程...\n[/cyan]")

if __name__ == "__main__":
    asyncio.run(main())