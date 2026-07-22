#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通过 GitHub REST API 直接把仓库文件上传到 GitHub（绕开 git 协议）。
适用于 git push 因网络限制不可达、但 api.github.com 可达的环境。

用法:
    GH_TOKEN=<你的token> python push_via_api.py <GITHUB_USERNAME> <REPO_NAME> [BASE_DIR]

- TOKEN 通过环境变量 GH_TOKEN 传入（避免出现在进程列表/命令行历史）
- TOKEN 需要 repo 权限（classic token 勾选 repo，或 fine-grained 给该仓库 Contents 写权限）
- REPO_NAME 不存在会自动创建（public）
"""
import os
import sys
import json
import base64
import subprocess
import urllib.request
import urllib.error
import urllib.parse

API = "https://api.github.com"


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
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.read().decode("utf-8"), resp.status
    except urllib.error.HTTPError as e:
        return e.read().decode("utf-8"), e.code


def main():
    if len(sys.argv) < 3:
        print("用法: GH_TOKEN=<token> python push_via_api.py <USERNAME> <REPO_NAME> [BASE_DIR]")
        sys.exit(1)

    token = os.environ.get("GH_TOKEN")
    if not token:
        print("错误: 请通过环境变量 GH_TOKEN 传入 GitHub token")
        sys.exit(1)

    username = sys.argv[1]
    repo = sys.argv[2]
    base_dir = sys.argv[3] if len(sys.argv) > 3 else os.getcwd()

    # 1) 创建仓库（已存在则忽略 409）
    body, code = api_call(
        "POST", "/user/repos", token,
        {"name": repo, "private": False, "auto_init": False,
         "description": "电商审单核对系统 - 销售单备注与发货条码数量自动核对 Web 应用"}
    )
    print(f"[create repo] {code} -> {repo}")
    if code not in (201, 409, 422):
        print("创建仓库失败:", body)
        sys.exit(1)

    # 2) 取得 git 跟踪的文件清单（关闭 quotePath，避免中文路径被八进制转义导致找不到文件）
    files = subprocess.check_output(
        ["git", "-c", "core.quotePath=false", "ls-files"], cwd=base_dir
    ).decode("utf-8").splitlines()
    files = [f for f in files if f.strip()]

    ok = 0
    skip = 0
    for rel in files:
        abs_path = os.path.join(base_dir, rel)
        if not os.path.isfile(abs_path):
            skip += 1
            continue
        # 跳过超大文件（API 单文件上限 ~1MB，保险起见限 800KB）
        size = os.path.getsize(abs_path)
        if size > 800 * 1024:
            print(f"[skip] {rel} 过大 ({size} bytes)")
            skip += 1
            continue
        with open(abs_path, "rb") as fh:
            content_b64 = base64.b64encode(fh.read()).decode("ascii")

        data = {"message": f"add {rel}", "content": content_b64}
        # 中文文件名需做百分号编码，否则 URL 含非 ASCII 字符会报错
        enc_rel = urllib.parse.quote(rel)
        # 若文件已存在，需要带上 sha 才能更新
        _, gcode = api_call("GET", f"/repos/{username}/{repo}/contents/{enc_rel}", token)
        if gcode == 200:
            try:
                data["sha"] = json.loads(_).get("sha")
            except Exception:
                pass

        _, pcode = api_call("PUT", f"/repos/{username}/{repo}/contents/{enc_rel}", token, data)
        if pcode in (200, 201):
            ok += 1
            print(f"[ok] {rel}")
        else:
            print(f"[fail {pcode}] {rel}")

    print(f"\n完成: 成功 {ok} 个, 跳过 {skip} 个")
    print(f"仓库地址: https://github.com/{username}/{repo}")


if __name__ == "__main__":
    main()
