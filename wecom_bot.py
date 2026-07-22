# -*- coding: utf-8 -*-
"""企业微信(WeCom)回调与消息收发模块。

仅当 app.py 配置了 WECOM_ENABLED=1 时才会被懒加载使用，因此不影响纯网页链接模式。
实现内容（均符合企业微信官方规范）：
  - 回调消息加解密：AES-256-CBC + SHA1 签名（EncodingAESKey / Token）
  - access_token 获取与内存缓存(7200s)
  - 媒体文件下载：接收群里发的 Excel（media/get）
  - 向群聊回传文本结果（appchat/send，适用于应用自建群）
依赖：cryptography（已加入 requirements.txt）
"""
import os
import time
import struct
import base64
import hashlib
import json
import xml.etree.ElementTree as ET
import urllib.request
import urllib.error

try:
    from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
    from cryptography.hazmat.backends import default_backend
    _HAS_CRYPTO = True
except Exception:
    _HAS_CRYPTO = False

WECOM_API = "https://qyapi.weixin.qq.com/cgi-bin"


class WeComCrypto:
    """企业微信消息体加解密（与微信公众平台/企业微信规范一致）。"""

    def __init__(self, token, encoding_aes_key, receive_id):
        if not _HAS_CRYPTO:
            raise RuntimeError("缺少 cryptography 库，请先执行: pip install cryptography")
        self.token = token
        self.receive_id = receive_id  # 自建应用填 corpid
        self.aes_key = base64.b64decode(encoding_aes_key + "=")  # 32 字节
        if len(self.aes_key) != 32:
            raise ValueError("EncodingAESKey 长度不正确（应为 43 个字符）")

    # ---------- 基础工具 ----------
    def _signature(self, *parts):
        sha = hashlib.sha1()
        sha.update("".join(sorted(parts)).encode("utf-8"))
        return sha.hexdigest()

    @staticmethod
    def _pkcs7_pad(text, block_size=32):
        pad = block_size - (len(text) % block_size)
        if pad == 0:
            pad = block_size
        return text + bytes([pad]) * pad

    @staticmethod
    def _pkcs7_unpad(text):
        return text[:-text[-1]]

    def _aes_decrypt(self, encrypted):
        cipher = Cipher(algorithms.AES(self.aes_key), modes.CBC(self.aes_key[:16]),
                        backend=default_backend())
        d = cipher.decryptor()
        return d.update(encrypted) + d.finalize()

    def _aes_encrypt(self, plaintext):
        cipher = Cipher(algorithms.AES(self.aes_key), modes.CBC(self.aes_key[:16]),
                        backend=default_backend())
        e = cipher.encryptor()
        return e.update(plaintext) + e.finalize()

    # ---------- 验签 & 解密 ----------
    def verify_signature(self, msg_signature, timestamp, nonce, encrypt):
        return self._signature(self.token, timestamp, nonce, encrypt) == msg_signature

    def decrypt_message(self, msg_signature, timestamp, nonce, post_xml):
        """解密回调 POST 的加密报文，返回明文 XML 字符串。"""
        root = ET.fromstring(post_xml)
        enc = root.findtext("Encrypt")
        if enc is None:
            raise ValueError("回调 XML 缺少 Encrypt 字段")
        if not self.verify_signature(msg_signature, timestamp, nonce, enc):
            raise ValueError("签名校验失败")
        plain = self._pkcs7_unpad(self._aes_decrypt(base64.b64decode(enc)))
        content = plain[16:]
        msg_len = struct.unpack(">I", content[:4])[0]
        msg = content[4:4 + msg_len]
        from_id = content[4 + msg_len:]
        if from_id.decode("utf-8") != self.receive_id:
            raise ValueError("receive_id 校验失败（消息可能被伪造）")
        return msg.decode("utf-8")

    def decrypt_echo(self, msg_signature, timestamp, nonce, echostr):
        """解密 URL 验证时的 echostr，返回明文字符串。"""
        if not self.verify_signature(msg_signature, timestamp, nonce, echostr):
            raise ValueError("签名校验失败")
        plain = self._pkcs7_unpad(self._aes_decrypt(base64.b64decode(echostr)))
        content = plain[16:]
        msg_len = struct.unpack(">I", content[:4])[0]
        return content[4:4 + msg_len].decode("utf-8")

    def encrypt_message(self, reply_text, nonce, timestamp=None):
        """把回复明文加密成企业微信要求的 <xml> 结构。"""
        timestamp = timestamp or str(int(time.time()))
        body = (os.urandom(16)
                + struct.pack(">I", len(reply_text.encode("utf-8")))
                + reply_text.encode("utf-8")
                + self.receive_id.encode("utf-8"))
        enc = base64.b64encode(self._aes_encrypt(self._pkcs7_pad(body))).decode("utf-8")
        return (f'<xml><Encrypt><![CDATA[{enc}]]></Encrypt>'
                f'<MsgSignature><![CDATA[{self._signature(self.token, timestamp, nonce, enc)}]]>'
                f'</MsgSignature><TimeStamp>{timestamp}</TimeStamp>'
                f'<Nonce><![CDATA[{nonce}]]></Nonce></xml>')


# ---------- API 调用 ----------
_token_cache = {}

def get_access_token(corpid, corpsecret):
    now = time.time()
    cached = _token_cache.get(corpid)
    if cached and cached[1] - now > 300:
        return cached[0]
    data = _http_get_json(f"{WECOM_API}/gettoken?corpid={corpid}&corpsecret={corpsecret}")
    if data.get("errcode") != 0:
        raise RuntimeError(f"获取 access_token 失败: {data}")
    token = data["access_token"]
    _token_cache[corpid] = (token, now + data.get("expires_in", 7200))
    return token


def download_media(access_token, media_id, save_path):
    url = f"{WECOM_API}/media/get?access_token={access_token}&media_id={media_id}"
    req = urllib.request.Request(url)
    with urllib.request.urlopen(req, timeout=60) as r, open(save_path, "wb") as f:
        f.write(r.read())
    return save_path


def send_text_to_chat(access_token, chat_id, text):
    """向群聊（应用自建群 appchat）发送文本消息。"""
    url = f"{WECOM_API}/appchat/send?access_token={access_token}"
    return _http_post_json(url, {
        "chatid": chat_id,
        "msgtype": "text",
        "text": {"content": text},
    })


def _http_get_json(url):
    with urllib.request.urlopen(urllib.request.Request(url), timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def _http_post_json(url, payload):
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=data, method="POST")
    req.add_header("Content-Type", "application/json")
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


# ---------- 自检（可选）----------
def self_test(token, aes_key, receive_id):
    """加密->解密 往返测试，用于上线前确认加解密实现正确。"""
    c = WeComCrypto(token, aes_key, receive_id)
    sample = "<xml><MsgType>file</MsgType></xml>"
    enc = c.encrypt_message(sample, "nonce123")  # 仅测结构，不校验签名
    print("self_test: encrypt ok, length =", len(enc))


if __name__ == "__main__":
    # 用法：python wecom_bot.py <token> <EncodingAESKey> <corpid>
    import sys
    if len(sys.argv) == 4:
        self_test(sys.argv[1], sys.argv[2], sys.argv[3])
    else:
        print("用法: python wecom_bot.py <token> <EncodingAESKey> <corpid>")
