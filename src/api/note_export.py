"""Obsidian / Word 笔记生成：每日邮件附件与手动导出（scripts/export_course.py）共用。

纯本地处理，不调用任何 API：Markdown 直接使用模型已经生成的摘要文本，
Word 由 pandoc 在本地转换。

Obsidian 笔记包结构（解压到 vault 根目录即可，重复解压直接覆盖）：
    iCourse/<课程>/<课程>.md            课程主页，用内置搜索块自动列出全部课次
    iCourse/<课程>/<课程> <课次>.md     每节课一篇，带 YAML 属性
公式保持 $...$ / $$...$$ 原样，由 Obsidian 原生渲染，不依赖任何图片。
"""

import io
import os
import re
import tempfile
import zipfile

_BAD_CHARS = set('\\/:*?"<>|#^[]')
_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def obsidian_name(text) -> str:
    """把 Obsidian 文件名 / 双链里不允许的字符换成下划线。"""
    cleaned = "".join("_" if c in _BAD_CHARS else c for c in str(text or ""))
    return cleaned.strip().strip(".") or "untitled"


def obsidian_tag(text) -> str:
    """Obsidian 标签只允许字母、数字、中文和 _-/，其余字符换成下划线。"""
    tag = "".join(c if (c.isalnum() or c in "_-/") else "_" for c in str(text or ""))
    return tag or "course"


def yaml_str(value) -> str:
    """输出 YAML 双引号字符串（转义反斜杠与双引号）。"""
    s = "" if value is None else str(value)
    return '"' + s.replace("\\", "\\\\").replace('"', '\\"') + '"'


def lecture_title(lec: dict) -> str:
    return str(lec.get("sub_title") or lec.get("sub_id") or "untitled")


def lecture_note(course_title: str, teacher: str, lec: dict) -> tuple[str, str]:
    """返回单节课笔记的 (zip 内路径, 内容)。"""
    folder = obsidian_name(course_title)
    title = lecture_title(lec)
    note = obsidian_name(f"{course_title} {title}")
    date = str(lec.get("date") or "")
    date_line = f"date: {date}" if _DATE_RE.match(date) else f"date: {yaml_str(date)}"
    header = f"> 课程：[[{folder}]]" + (f"　教师：{teacher}" if teacher else "")
    content = "\n".join([
        "---",
        f"course: {yaml_str(course_title)}",
        f"teacher: {yaml_str(teacher or '')}",
        date_line,
        f"lecture: {yaml_str(title)}",
        "tags:",
        "  - iCourse",
        f"  - {obsidian_tag(course_title)}",
        "---",
        "",
        header,
        "",
        (lec.get("summary") or "").strip(),
        "",
    ])
    return f"iCourse/{folder}/{note}.md", content


def course_index(course_title: str) -> tuple[str, str]:
    """返回课程主页的 (zip 内路径, 内容)。

    主页用 Obsidian 内置的搜索块实时列出该课程文件夹下的笔记，内容与
    具体有哪些课次无关，所以每次覆盖都是安全的，永远不会"变旧"。
    """
    folder = obsidian_name(course_title)
    scope = f'path:"iCourse/{folder}/" -file:"{folder}.md"'
    content = "\n".join([
        "---",
        f"course: {yaml_str(course_title)}",
        "tags:",
        "  - iCourse",
        f"  - {obsidian_tag(course_title)}",
        "---",
        "",
        "## 课次",
        "",
        "```query",
        scope,
        "```",
        "",
        "## 课程事项提醒",
        "",
        "```query",
        f'{scope} "课程事项提醒"',
        "```",
        "",
    ])
    return f"iCourse/{folder}/{folder}.md", content


def build_obsidian_zip(courses: list[tuple[str, str, list[dict]]]) -> bytes:
    """courses: [(课程名, 教师, [lecture dict, ...]), ...] → zip 字节。"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for course_title, teacher, lectures in courses:
            for lec in lectures:
                zf.writestr(*lecture_note(course_title, teacher, lec))
            zf.writestr(*course_index(course_title))
    return buf.getvalue()


def build_docx(course_title: str, teacher: str, lectures: list[dict]) -> bytes:
    """用 pandoc 把 Markdown 转成 Word；公式转成 Word 原生公式（可编辑），不是图片。"""
    import pypandoc  # noqa: PLC0415  仅 Word 导出需要

    parts = [f"# {course_title}", f"任课教师：{teacher or ''}"]
    for lec in lectures:
        title = lecture_title(lec)
        date = str(lec.get("date") or "")
        heading = title if (not date or date in title) else f"{title}（{date}）"
        parts.append(f"## {heading}")
        parts.append((lec.get("summary") or "").strip())
    md = "\n\n".join(parts)
    with tempfile.TemporaryDirectory() as tmp:
        out = os.path.join(tmp, "notes.docx")
        # commonmark_x：标题、列表前不强制空行，且支持 $...$ / $$...$$ 公式
        pypandoc.convert_text(md, "docx", format="commonmark_x", outputfile=out)
        with open(out, "rb") as f:
            return f.read()
