# -*- coding: utf-8 -*-
"""一键启动：Flask服务 + 公网隧道（带自动重连）"""
import subprocess
import sys
import os
import time
import re
import signal
import urllib.request

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PYTHON = r'C:\Users\EDY\.workbuddy\binaries\python\versions\3.12.8\python.exe'
APP_PY = os.path.join(BASE_DIR, 'app.py')
URL_FILE = os.path.join(BASE_DIR, 'current_url.txt')


def test_local_flask(timeout=5):
    """测试本地Flask服务是否可访问"""
    try:
        urllib.request.urlopen('http://127.0.0.1:5000/', timeout=timeout)
        return True
    except Exception:
        return False


def start_flask():
    """启动Flask服务，返回进程对象"""
    print("[1/2] 正在启动 Flask 服务...")
    proc = subprocess.Popen(
        [PYTHON, APP_PY],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
    )
    # 等待并确认启动
    for _ in range(10):
        time.sleep(1)
        if test_local_flask(timeout=3):
            print("      ✓ Flask 服务已启动 (端口 5000)")
            return proc
        if proc.poll() is not None:
            break
    print("      ✗ Flask 服务启动失败！")
    proc.terminate()
    return None


def start_ssh_tunnel():
    """启动SSH隧道，返回 (进程对象, 公网URL)"""
    print("\n[2/2] 正在创建公网隧道...")
    print("      (连接到 localhost.run，可能需要几秒钟...)")

    proc = subprocess.Popen(
        ['ssh', '-o', 'StrictHostKeyChecking=no', '-o', 'ConnectTimeout=15',
         '-o', 'ServerAliveInterval=60', '-o', 'ServerAliveCountMax=3',
         '-o', 'ExitOnForwardFailure=yes',
         '-R', '80:localhost:5000', 'nokey@localhost.run'],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding='utf-8',
        errors='replace'
    )

    public_url = None
    start_time = time.time()
    while time.time() - start_time < 30:
        line = ssh_proc.stdout.readline()
        if not line:
            if proc.poll() is not None:
                print("      ✗ SSH 隧道连接失败！")
                return None, None
            time.sleep(0.2)
            continue

        line = line.strip()
        url_match = re.search(r'https://[a-z0-9-]+\.[a-z.]+', line)
        if url_match and 'localhost.run' not in url_match.group():
            public_url = url_match.group()
            break

    return proc, public_url


def save_url(public_url):
    """保存公网URL到文件"""
    try:
        with open(URL_FILE, 'w', encoding='utf-8') as f:
            f.write(public_url)
    except Exception as e:
        print(f"      保存URL文件失败: {e}")


def main():
    print("=" * 55)
    print("     电商审单核对系统 - 正在启动...")
    print("     (公网隧道断开后会自动重连)")
    print("=" * 55)
    print()

    # 1. 启动 Flask 服务
    flask_proc = start_flask()
    if not flask_proc:
        print("\n按任意键退出...")
        input()
        return

    # 2. 启动 SSH 隧道（带自动重连）
    running = True
    first_time = True
    try:
        while running:
            ssh_proc = subprocess.Popen(
                ['ssh', '-o', 'StrictHostKeyChecking=no', '-o', 'ConnectTimeout=15',
                 '-o', 'ServerAliveInterval=60', '-o', 'ServerAliveCountMax=3',
                 '-o', 'ExitOnForwardFailure=yes',
                 '-R', '80:localhost:5000', 'nokey@localhost.run'],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding='utf-8',
                errors='replace'
            )

            public_url = None
            start_time = time.time()
            while time.time() - start_time < 30:
                line = ssh_proc.stdout.readline()
                if not line:
                    if ssh_proc.poll() is not None:
                        break
                    time.sleep(0.2)
                    continue

                line = line.strip()
                url_match = re.search(r'https://[a-z0-9-]+\.[a-z.]+', line)
                if url_match and 'localhost.run' not in url_match.group():
                    public_url = url_match.group()
                    break

            if public_url:
                save_url(public_url)
                if first_time:
                    print(f"      ✓ 公网隧道已创建!")
                    print()
                    print("=" * 55)
                    print("  🎉 启动成功！访问地址如下：")
                    print("=" * 55)
                    print()
                    print(f"  📌 公网访问（任何网络）：{public_url}")
                    print(f"  📌 局域网访问：         http://172.16.10.84:5000")
                    print(f"  📌 本机访问：           http://127.0.0.1:5000")
                    print()
                    print("  ⚠️  注意事项：")
                    print("     • 公网地址每次启动会变化，请以本次显示为准")
                    print("     • 本程序需要保持运行，关闭后网页将无法访问")
                    print("     • 隧道断开时会自动重连，无需手动操作")
                    print("     • 按 Ctrl+C 可停止服务")
                    print("=" * 55)
                    print()
                    first_time = False
                else:
                    print(f"      ✓ 公网隧道已重连: {public_url}")
            else:
                print("      ✗ 未能获取公网URL，3秒后重试...")
                try:
                    ssh_proc.terminate()
                except Exception:
                    pass
                time.sleep(3)
                continue

            # 监控隧道和Flask服务
            while True:
                time.sleep(2)
                if flask_proc.poll() is not None:
                    print("\n Flask 服务已停止，正在重启...")
                    flask_proc = start_flask()
                    if not flask_proc:
                        running = False
                        break
                if ssh_proc.poll() is not None:
                    print("\n SSH 隧道已断开，正在自动重连...")
                    break

    except KeyboardInterrupt:
        print("\n\n 正在停止所有服务...")
    finally:
        try:
            ssh_proc.terminate()
        except Exception:
            pass
        try:
            flask_proc.terminate()
        except Exception:
            pass
        print(" 服务已停止。")


if __name__ == '__main__':
    main()
