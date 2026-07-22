#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
精准推送指定文件到 GitHub 仓库（绕开 git 协议，仅依赖 api.github.com）。
用法:
    GH_TOKEN=<token> python push_files.py
- 文件清单在 FILES 中维护：每一项为 (本地相对路径, 仓库内相对路径)
- 已存在则自动带上 sha 进行更新
"""
import os
import sys
import json
import base64
import urllib.request
import urllib.error
import urllib.parse

API = "https://api.github.com"
USERNAME = "maquan1234"
REPO = "order-verify-web"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# (本地相对路径, 仓库内相对路径)
FILES = [
    ("app.py", "app.py"),
    ("verifier.py", "verifier.py"),
    ("build.spec", "build.spec"),
    ("打包.bat", "打包.bat"),
    ("打包说明.md", "打包说明.md"),
    ("requirements.txt", "requirements.txt"),
    ("products.json", "products.json"),
    ("templates/index.html", "templates/index.html"),
    (".github/workflows/build-exe.yml", ".github/workflows/build-exe.yml"),
    ("push_files.py", "push_files.py"),
    ("push_via_api.py", "push_via_api.py"),
    ("DEPLOY.md", "DEPLOY.md"),
]


def api_call(method, path, token, data=None):
    url = API + path
    req = urllib.request.Request(url, method=method)
    req.add_header("Authorization", f"token {token}")
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("User-Agent", "order-verify-push")
    if data is not None:
        req.add_header("Content-Type", "application/json")
        req.data = json.dumps(data).encode("utf-8")
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            return resp.read().decode("utf-8"), resp.status
    except urllib.error.HTTPError as e:
        return e.read().decode("utf-8"), e.code


def main():
    token = os.environ.get("GH_TOKEN")
    if not token:
        print("错误: 请通过环境变量 GH_TOKEN 传入 GitHub token")
        sys.exit(1)

    ok = skip = fail = 0
    for local_rel, repo_rel in FILES:
        abs_path = os.path.join(BASE_DIR, local_rel)
        if not os.path.isfile(abs_path):
            print(f"[skip] 本地不存在: {local_rel}")
            skip += 1
            continue
        size = os.path.getsize(abs_path)
        if size > 800 * 1024:
            print(f"[skip] 过大 ({size} bytes): {local_rel}")
            skip += 1
            continue
        with open(abs_path, "rb") as fh:
            content_b64 = base64.b64encode(fh.read()).decode("ascii")

        data = {"message": f"update {repo_rel}", "content": content_b64}
        enc_rel = urllib.parse.quote(repo_rel)
        _, gcode = api_call("GET", f"/repos/{USERNAME}/{REPO}/contents/{enc_rel}", token)
        if gcode == 200:
            try:
                data["sha"] = json.loads(_).get("sha")
            except Exception:
                pass
        _, pcode = api_call("PUT", f"/repos/{USERNAME}/{REPO}/contents/{enc_rel}", token, data)
        if pcode in (200, 201):
            ok += 1
            print(f"[ok] {repo_rel}")
        else:
            fail += 1
            print(f"[fail {pcode}] {repo_rel}")

    print(f"\n完成: 成功 {ok}, 跳过 {skip}, 失败 {fail}")
    print(f"仓库: https://github.com/{USERNAME}/{REPO}")


if __name__ == "__main__":
    main()
