#!/usr/bin/env python3
"""Outlook .msg 본문 추출기.

사용법
  python3 tools/read_msg.py <파일> [--attach-dir DIR] [--html] [--full]

<파일> 은 둘 중 하나.
  1) .msg 파일 자체
  2) Drive MCP `download_file_content` 결과가 저장된 .txt (JSON: {content: base64, ...})

의존성: pip install olefile
"""
import base64
import io
import json
import re
import struct
import sys

import olefile

# ---------------------------------------------------------------- 속성 태그
TAGS = {
    "0037": "subject",
    "0C1A": "sender_name",
    "0C1F": "sender_email",
    "0042": "sent_repr_name",
    "0065": "sent_repr_email",
    "0E04": "to",
    "0E03": "cc",
    "0E02": "bcc",
    "1000": "body",
    "1013": "html_body",
    "1009": "rtf_compressed",
    "007D": "headers",
    "3707": "attach_long_name",
    "3704": "attach_name",
    "3701": "attach_data",
}
TIME_TAGS = {0x0039: "sent", 0x0E06: "delivered", 0x3007: "created"}


def _decode(data, typ):
    if typ == "001F":
        return data.decode("utf-16-le", "ignore").rstrip("\x00")
    if typ == "001E":
        for enc in ("cp949", "utf-8", "latin-1"):
            try:
                return data.decode(enc).rstrip("\x00")
            except UnicodeDecodeError:
                continue
    return data


# ------------------------------------------------------- RTF (LZFu) 압축 해제
_LZFU_DICT = (
    "{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman "
    "\\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArialTimes New "
    "RomanCourier{\\colortbl\\red0\\green0\\blue0\r\n\\par "
    "\\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx"
)


def rtf_decompress(data):
    if len(data) < 16:
        return b""
    _csize, rsize, magic, _crc = struct.unpack("<4I", data[:16])
    if magic == 0x414C454D:              # MELA - 비압축
        return data[16:]
    if magic != 0x75465A4C:              # LZFu
        return b""
    buf = bytearray(_LZFU_DICT.encode("latin-1"))
    buf.extend(b"\x00" * (4096 - len(buf)))
    wp = len(_LZFU_DICT)
    out = bytearray()
    src = data[16:]
    i = 0
    while i < len(src) and len(out) < rsize:
        flags = src[i]
        i += 1
        for bit in range(8):
            if i >= len(src) or len(out) >= rsize:
                break
            if flags & (1 << bit):        # 사전 참조
                if i + 1 >= len(src):
                    break
                a, b = src[i], src[i + 1]
                i += 2
                offset = (a << 4) | (b >> 4)
                length = (b & 0x0F) + 2
                if offset == wp:
                    return bytes(out)
                for _ in range(length):
                    ch = buf[offset % 4096]
                    offset += 1
                    out.append(ch)
                    buf[wp % 4096] = ch
                    wp += 1
            else:                          # 리터럴
                ch = src[i]
                i += 1
                out.append(ch)
                buf[wp % 4096] = ch
                wp += 1
    return bytes(out)


def rtf_to_text(rtf):
    """RTF 에서 평문만 뽑는다. \\'XX 바이트는 cp949 로 되돌린다."""
    s = rtf.decode("latin-1", "ignore")
    s = re.sub(r"\{\\\*?\\(fonttbl|colortbl|stylesheet|generator|pict|object)[^{}]*"
               r"(\{[^{}]*\}[^{}]*)*\}", "", s)
    s = s.replace("\\par", "\n").replace("\\line", "\n").replace("\\tab", "\t")
    s = re.sub(r"\\u(-?\d+)\??", lambda m: chr(int(m.group(1)) % 65536), s)
    out = bytearray()
    i = 0
    while i < len(s):
        c = s[i]
        if c == "\\":
            if s[i + 1:i + 2] == "'":
                out.extend(bytes([int(s[i + 2:i + 4], 16)]))
                i += 4
                continue
            m = re.match(r"\\([a-zA-Z]+)(-?\d+)? ?", s[i:])
            if m:
                i += m.end()
                continue
            if s[i + 1:i + 2] == "*":          # \* 목적지 표시자는 버린다
                i += 2
                continue
            out.extend(s[i + 1].encode("latin-1", "ignore"))
            i += 2
            continue
        if c in "{}":
            i += 1
            continue
        out.extend(c.encode("latin-1", "ignore"))
        i += 1
    txt = out.decode("cp949", "ignore")
    if re.search(r"(?i)<(html|body|div|p)\b", txt):   # HTML 캡슐화 RTF
        return html_to_text(txt)
    return _squeeze(txt)


def _squeeze(text):
    """공백뿐인 줄을 없애고 빈 줄이 2개 이상 이어지지 않게 한다."""
    lines = [ln.rstrip() for ln in text.replace("\r\n", "\n").split("\n")]
    lines = ["" if not ln.strip() else ln for ln in lines]
    out = []
    for ln in lines:
        if ln == "" and (not out or out[-1] == ""):
            continue
        out.append(ln)
    return "\n".join(out).strip()


def html_to_text(html):
    s = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", "", html)
    s = re.sub(r"(?i)<br\s*/?>", "\n", s)
    s = re.sub(r"(?i)</(p|div|tr|li|h[1-6])>", "\n", s)
    s = re.sub(r"<[^>]+>", "", s)
    for a, b in (("&nbsp;", " "), ("&amp;", "&"), ("&lt;", "<"),
                 ("&gt;", ">"), ("&quot;", '"'), ("&#39;", "'")):
        s = s.replace(a, b)
    return _squeeze(s)


# ----------------------------------------------------------------- 시간 속성
def read_times(ole):
    res = {}
    try:
        raw = ole.openstream("__properties_version1.0").read()
    except OSError:
        return res
    for off in range(32, len(raw) - 15, 16):
        tag = struct.unpack("<I", raw[off:off + 4])[0]
        pid, ptyp = tag >> 16, tag & 0xFFFF
        if ptyp != 0x0040 or pid not in TIME_TAGS:
            continue
        ft = struct.unpack("<Q", raw[off + 8:off + 16])[0]
        if ft:
            import datetime
            res[TIME_TAGS[pid]] = (
                datetime.datetime(1601, 1, 1)
                + datetime.timedelta(microseconds=ft // 10)
            ).strftime("%Y-%m-%d %H:%M:%S UTC")
    return res


# -------------------------------------------------------------------- 메인
def load(path):
    with open(path, "rb") as fh:
        head = fh.read(8)
    if head[:4] == b"\xd0\xcf\x11\xe0":
        return open(path, "rb").read()
    with open(path, encoding="utf-8") as fh:
        return base64.b64decode(json.load(fh)["content"])


def parse(raw, attach_dir=None):
    ole = olefile.OleFileIO(io.BytesIO(raw))
    msg = {"attachments": []}
    msg.update(read_times(ole))
    slots = {}
    for entry in ole.listdir():
        m = re.match(r"__substg1\.0_([0-9A-F]{4})([0-9A-F]{4})", entry[-1])
        if not m:
            continue
        tag, typ = m.group(1), m.group(2)
        name = TAGS.get(tag)
        if not name:
            continue
        data = ole.openstream(entry).read()
        branch = "/".join(entry[:-1])
        slots.setdefault(branch, {})[name] = (
            data if typ == "0102" else _decode(data, typ))

    msg.update(slots.get("", {}))

    # 본문: plain → html → rtf 순
    if not msg.get("body"):
        if msg.get("html_body"):
            hb = msg["html_body"]
            hb = hb.decode("cp949", "ignore") if isinstance(hb, bytes) else hb
            msg["body"] = html_to_text(hb)
            msg["body_source"] = "html"
        elif msg.get("rtf_compressed"):
            msg["body"] = rtf_to_text(rtf_decompress(msg["rtf_compressed"]))
            msg["body_source"] = "rtf"
    else:
        msg["body_source"] = "plain"

    for branch, props in sorted(slots.items()):
        if not branch.startswith("__attach"):
            continue
        fname = props.get("attach_long_name") or props.get("attach_name") or branch
        blob = props.get("attach_data")
        msg["attachments"].append({"name": fname,
                                   "size": len(blob) if blob else 0})
        if attach_dir and blob:
            import os
            os.makedirs(attach_dir, exist_ok=True)
            safe = re.sub(r"[/\\]", "_", fname)
            with open(os.path.join(attach_dir, safe), "wb") as fh:
                fh.write(blob)
    for k in ("html_body", "rtf_compressed"):
        msg.pop(k, None)
    return msg


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    opts = {a for a in sys.argv[1:] if a.startswith("--")}
    if not args:
        print(__doc__)
        return 1
    attach_dir = None
    if "--attach-dir" in sys.argv:
        attach_dir = sys.argv[sys.argv.index("--attach-dir") + 1]
        args = [a for a in args if a != attach_dir]
    msg = parse(load(args[0]), attach_dir)
    limit = None if "--full" in opts else 4000
    print("제목  :", msg.get("subject", ""))
    print("발신  :", msg.get("sender_name") or msg.get("sent_repr_name", ""),
          f"<{msg.get('sender_email') or msg.get('sent_repr_email', '')}>")
    print("수신  :", msg.get("to", ""))
    print("참조  :", msg.get("cc", ""))
    print("일시  :", msg.get("sent") or msg.get("delivered", ""))
    if msg["attachments"]:
        print("첨부  :", ", ".join(
            f"{a['name']} ({a['size']:,}B)" for a in msg["attachments"]))
    print(f"--- 본문 ({msg.get('body_source', 'none')}) ---")
    print((msg.get("body") or "")[:limit])
    return 0


if __name__ == "__main__":
    sys.exit(main())
