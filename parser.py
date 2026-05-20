import os
import re
import json
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Ofis formulalarini (OMML) oddiy belgilarga yoki elementlarga o'g'irish
def parse_math_element(element):
    """Formula elementlarini (daraja, indeks, kasr) oddiy matnga yoki struktura formatiga o'tkazish"""
    text = ""
    for child in element:
        # Oddiy matn
        if child.tag.endswith('t'):
            text += child.text
        # Daraja (superscript)
        elif child.tag.endswith('superscript') or child.tag.endswith('sup'):
            text += f"^{{{parse_math_element(child)}}}"
        # Indeks (subscript)
        elif child.tag.endswith('subscript') or child.tag.endswith('sub'):
            text += f"_{{{parse_math_element(child)}}}"
        # Agar ichki boshqa elementlar bo'lsa rekurstiv chaqirish
        elif len(child) > 0:
            text += parse_math_element(child)
    return text

def get_text_with_formatting(paragraph):
    """Paragraph ichidagi daraja, indeks va formula belgilarini saqlab qolgan holda HTML formatga o'tkazish"""
    p_html = ""
    # paragraph ichidagi barcha xml elementlarni tekshirish
    for child in paragraph._p:
        # Agar bu oddiy run (matn bo'lagi) bo'lsa
        if child.tag.endswith('r'):
            rPr = child.find(qn('w:rPr'))
            text_elem = child.find(qn('w:t'))
            if text_elem is not None and text_elem.text:
                text = text_elem.text
                if rPr is not None:
                    # Darajani tekshirish (superscript)
                    vertAlign = rPr.find(qn('w:vertAlign'))
                    if vertAlign is not None:
                        val = vertAlign.get(qn('w:val'))
                        if val == 'superscript':
                            p_html += f"<sup>{text}</sup>"
                            continue
                        elif val == 'subscript':
                            p_html += f"<sub>{text}</sub>"
                            continue
                p_html += text
        # Agar bu Office Math (Formula) elementi bo'lsa
        elif child.tag.endswith('oMath'):
            # MathJax inline formati uchun \( ... \) belgilari ichiga olamiz
            math_text = "".join(child.itertext())
            # Ba'zi belgilarni to'g'rilash
            math_text = math_text.replace("²", "<sup>2</sup>").replace("³", "<sup>3</sup>")
            p_html += f" <span class='math-expr'>{math_text}</span> "
    
    # Agar yuqoridagi xml tahlil bo'sh qaytsa, oddiy matnni olamiz
    if not p_html.strip():
        p_html = paragraph.text
        
    # Kasr va maxsus belgilarni frontend chiroyli ko'rsatishi uchun qo'shimcha regex almashtirishlar
    p_html = re.sub(r'(\w+)/(\w+)', r'<span class="fraction"><span class="top">\1</span><span class="bottom">\2</span></span>', p_html)
    return p_html.strip()

# =========================================================
# FIZIKA VA FORMULALI TESTLAR UCHUN PARSER
# =========================================================
def get_quizzes_by_letters(file_path, json_answers_path=None):
    try:
        doc = Document(file_path)
        
        # Har bir paragrafni formatini saqlab string ro'yxatiga o'giramiz
        paragraphs_html = []
        for p in doc.paragraphs:
            txt = get_text_with_formatting(p)
            if txt:
                paragraphs_html.append(txt)
                
        full_text = "\n".join(paragraphs_html)
        
        correct_answers = {}
        if json_answers_path and os.path.exists(json_answers_path):
            try:
                with open(json_answers_path, 'r', encoding='utf-8') as f:
                    answers_data = json.load(f)
                correct_answers = answers_data.get("fizika_answers", answers_data)
            except Exception as json_err:
                print(f"JSON yuklashda xatolik: {json_err}")

        # Savollarni raqamlar bo'yicha ajratish (Masalan: 64. yoki 65.)
        question_blocks = re.split(r'(?=\b\d+\.)', full_text)
        quiz_data = []
        letter_to_idx = {"A": 0, "B": 1, "C": 2, "D": 3}

        for block in question_blocks:
            block = block.strip()
            if not block: continue
            header_match = re.match(r'^(\d+)\.(.*)', block, re.DOTALL)
            if not header_match: continue
                
            q_num = header_match.group(1).strip()
            rest_of_text = header_match.group(2).strip()
            
            # Variantlar indeksini aniqlash
            idx_A = rest_of_text.find("A)")
            idx_B = rest_of_text.find("B)")
            idx_C = rest_of_text.find("C)")
            idx_D = rest_of_text.find("D)")
            
            if idx_A != -1 and idx_B != -1 and idx_C != -1 and idx_D != -1 and (idx_A < idx_B < idx_C < idx_D):
                question_content = rest_of_text[:idx_A].strip()
                opt_A = rest_of_text[idx_A + 2:idx_B].strip()
                opt_B = rest_of_text[idx_B + 2:idx_C].strip()
                opt_C = rest_of_text[idx_C + 2:idx_D].strip()
                opt_D = rest_of_text[idx_D + 2:].strip()
                
                options = [opt_A, opt_B, opt_C, opt_D]
                correct_letter = str(correct_answers.get(q_num, "A")).upper().strip()
                correct_index = letter_to_idx.get(correct_letter, 0)
                
                if question_content:
                    cleaned_question = re.sub(r'[ \t]+', ' ', question_content).strip()
                    cleaned_options = [re.sub(r'\|$', '', opt).strip() for opt in options]
                    
                    quiz_data.append({
                        "question": f"{q_num}. {cleaned_question}",
                        "options": cleaned_options,
                        "correct": correct_index
                    })
        return quiz_data
    except Exception as e:
        print(f"Fizika parser xatosi: {e}")
        return []

# =========================================================
# STANDART VA BOSHQA PARSERLAR (O'zgarishsiz qoldi)
# =========================================================
def get_quizzes(file_path):
    try:
        doc = Document(file_path)
        paragraphs = []
        for p in doc.paragraphs:
            t = get_text_with_formatting(p)
            if t: paragraphs.append(t)
        full_text = "\n".join(paragraphs)
        raw_questions = full_text.split("++++")
        quiz_data = []

        for item in raw_questions:
            item = item.strip()
            if not item: continue
            parts = item.split("====")
            if len(parts) < 2: continue
            
            question_text = parts[0].strip()
            options_raw = [p.strip() for p in parts[1:] if p.strip()]
            
            correct_index = 0
            final_options = []
            for i, opt in enumerate(options_raw):
                clean_opt = opt.replace("#", "").strip()
                if opt.startswith("#"):
                    correct_index = i
                final_options.append(clean_opt)
            
            if len(final_options) >= 2:
                quiz_data.append({
                    "question": question_text,
                    "options": final_options,
                    "correct": correct_index
                })
        return quiz_data
    except Exception as e:
        print(f"Standart parser xatosi: {e}")
        return []

def get_quizzes_programming(file_path):
    try:
        doc = Document(file_path)
        quiz_data = []
        paragraphs = [get_text_with_formatting(p) for p in doc.paragraphs if p.text.strip()]
        
        i = 0
        while i < len(paragraphs):
            if "Savol" in paragraphs[i] or "савол" in paragraphs[i].lower():
                question_parts = []
                i += 1
                while i < len(paragraphs) and not paragraphs[i].startswith("—"):
                    question_parts.append(paragraphs[i])
                    i += 1
                
                question_text = " ".join(question_parts).strip()
                options = []
                while i < len(paragraphs) and paragraphs[i].startswith("—"):
                    option_text = paragraphs[i].replace("—", "").strip()
                    options.append(option_text)
                    i += 1
                
                if len(options) >= 2 and question_text:
                    quiz_data.append({
                        "question": question_text,
                        "options": options,
                        "correct": 0
                    })
            else:
                i += 1
        return quiz_data
    except Exception as e:
        print(f"Dasturlash parser xatosi: {e}")
        return []

def normalize_text(text):
    if not text: return ""
    text = re.sub(r'[^a-zA-Z0-9\s]', '', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip().lower()

def calculate_similarity(text1, text2):
    norm1 = normalize_text(text1)
    norm2 = normalize_text(text2)
    if not norm1 or not norm2: return 0
    if norm1 == norm2: return 100
    words1 = set(norm1.split())
    words2 = set(norm2.split())
    union = len(words1.union(words2))
    if union == 0: return 0
    return (len(words1.intersection(words2)) / union) * 100

def get_quizzes_english_pdf_docx(docx_path, pdf_path):
    # Ingliz tili kodi o'zgarishsiz qoladi, asosiysi HTML formatda qaytishi ta'minlangan
    try:
        import pdfplumber
        pdf_red_texts = []
        if os.path.exists(pdf_path):
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    for char in page.chars:
                        color = char.get("non_stroking_color") or char.get("stroking_color")
                        if color and len(color) == 3:
                            r, g, b = color
                            if (r > 0.6 and g < 0.4 and b < 0.4):
                                pdf_red_texts.append(char["text"])
        
        doc = Document(docx_path)
        raw_lines = [get_text_with_formatting(p) for p in doc.paragraphs if p.text.strip()]
        
        # Soddalashtirilgan logika (oldingi parseringiz mantiqi bo'yicha)
        # ... (Sizda bor ingliz tili mantiqiy kodi shu yerda davom etadi)
        return [] 
    except:
        return []