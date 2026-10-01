"""Local resume text extraction and conservative experience splitting.

The parser never invents facts. It keeps source text and reports conflicting
dates instead of silently deciding which version is correct.
"""
from collections import Counter, defaultdict
from pathlib import Path
import io, re, zipfile

import fitz
from lxml import etree as E

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
GROUP_ALIASES = {
    "教育背景": "教育背景", "教育经历": "教育背景", "education": "教育背景",
    "实习经历": "实习工作", "实习工作": "实习工作", "工作经历": "实习工作",
    "experience": "实习工作", "legal experience": "实习工作",
    "社会实践": "社会实践", "学工经历": "社会实践", "学工及实践": "社会实践",
    "校园经历": "社会实践", "activities": "社会实践",
    "科研经历": "科研课题", "科研课题": "科研课题", "项目经历": "科研课题",
    "research": "科研课题", "projects": "科研课题",
    "荣誉及奖励": "荣誉与技能", "荣誉奖励": "荣誉与技能", "荣誉与技能": "荣誉与技能",
    "honors": "荣誉与技能", "skills": "荣誉与技能",
    "兴趣爱好": "兴趣爱好", "interests": "兴趣爱好",
}
DATE_RE = re.compile(r"(?:(?:19|20)\d{2}[./年-]\d{1,2}(?:月)?|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\s+(?:19|20)\d{2})(?:\s*(?:[-–—至~]|to)\s*(?:(?:19|20)?\d{2}[./-]?\d{0,2}|至今|现在|Present|Current))?", re.I)

def _clean_lines(lines):
    out=[]
    for line in lines:
        line=re.sub(r"[\u200b\uf0b7]", "", line)
        line=re.sub(r"\s+", " ", line).strip(" \t•·")
        if line and (not out or line != out[-1]):out.append(line)
    # Some Word files expose every textbox twice. Remove an exact repeated half.
    if len(out)%2==0 and out[:len(out)//2] == out[len(out)//2:]:out=out[:len(out)//2]
    return out

def extract_text(name, blob):
    suffix=Path(name).suffix.lower()
    if suffix == ".pdf":
        doc=fitz.open(stream=blob,filetype="pdf")
        lines=[]
        for page in doc:lines.extend(page.get_text("text").splitlines())
        doc.close();return _clean_lines(lines)
    if suffix == ".docx":
        with zipfile.ZipFile(io.BytesIO(blob)) as z:
            info=z.getinfo("word/document.xml")
            if info.file_size>20_000_000:raise ValueError(f"{name} 解压后过大")
            root=E.fromstring(z.read(info),E.XMLParser(resolve_entities=False,no_network=True))
        return _clean_lines("".join(p.xpath(".//w:t/text()",namespaces=NS)) for p in root.findall(".//w:p",NS))
    if suffix in {".txt",".md"}:return _clean_lines(blob.decode("utf-8-sig").splitlines())
    raise ValueError(f"{name}：仅支持 PDF、DOCX、UTF-8 TXT 和 Markdown")

def _heading(line):
    key=re.sub(r"[\s:：|｜]", "", line).lower()
    for alias,group in GROUP_ALIASES.items():
        if key == re.sub(r"[\s:：|｜]", "", alias).lower():return group

def _looks_like_title(line):
    if not DATE_RE.search(line):return False
    no_date=DATE_RE.sub("",line).strip(" |｜·—–-")
    return len(no_date)>=2 and ("｜" in line or "|" in line or len(no_date)<=80)

def _split_title_date(line):
    matches=list(DATE_RE.finditer(line))
    if not matches:return line,""
    m=matches[-1];title=(line[:m.start()]+line[m.end():]).strip(" |｜·—–-")
    return title,m.group(0).replace("-","–")

def _guess_group(title,current):
    if current:return current
    t=title.lower()
    if any(x in t for x in ["大学","学院","university","school"]):return "教育背景"
    if any(x in t for x in ["基金","课题","研究","论文","competition","project","research"]):return "科研课题"
    if any(x in t for x in ["志愿","学生会","记者团","辩论队","班级","volunteer","association"]):return "社会实践"
    return "实习工作"

def parse_lines(lines,source):
    result=[];group=None;current=None
    def finish():
        nonlocal current
        if not current:return
        description="".join(current.pop("body")).strip()
        current["description"]=description;current["raw"]=(current["title"]+" "+current["date"]+"\n"+description).strip()
        if current["title"] and (description or current["group"]=="教育背景"):result.append(current)
        current=None
    for line in lines:
        maybe=_heading(line)
        if maybe:finish();group=maybe;continue
        if _looks_like_title(line):
            finish();title,date=_split_title_date(line)
            current={"group":_guess_group(title,group),"title":title,"date":date,"body":[],"source":source}
        elif current:
            if len("".join(current["body"]))<6000:current["body"].append(line)
    finish();return result

def _key(title):
    return re.sub(r"[^\w\u4e00-\u9fff]", "", title).lower()

def parse_resume_files(files):
    parsed=[];warnings=[]
    for item in files:
        lines=extract_text(item["name"],item["blob"])
        candidates=parse_lines(lines,item["name"])
        if not candidates:warnings.append(f"{item['name']}：未识别出结构化经历，文件未被丢弃，但需要手动录入")
        parsed.extend(candidates)
    buckets=defaultdict(list)
    for item in parsed:buckets[_key(item["title"])].append(item)
    merged=[]
    for variants in buckets.values():
        first=variants[0];dates=[v["date"] for v in variants if v["date"]];descs=[v["description"] for v in variants if v["description"]]
        date=Counter(dates).most_common(1)[0][0] if dates else ""
        description=max(descs,key=len) if descs else ""
        conflicts=[]
        if len(set(dates))>1:conflicts.append("不同文件中的日期不一致："+" / ".join(sorted(set(dates))))
        merged.append({
            "group":Counter(v["group"] for v in variants).most_common(1)[0][0],
            "title":first["title"],"date":date,"description":description,
            "raw":"\n\n".join(v["raw"] for v in variants),
            "english_title":"","english_date":"","english":"","english_pending":True,
            "sources":sorted({v["source"] for v in variants}),"conflicts":conflicts,
        })
    return merged,warnings
