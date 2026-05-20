import os
import random
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

try:
    from parser import get_quizzes, get_quizzes_programming, get_quizzes_english_pdf_docx, get_quizzes_by_letters
except ImportError:
    raise ImportError("Iltimos, 'parser.py' fayli ushbu papkada borligini va nomi to'g'ri yozilganini tekshiring!")

app = FastAPI(title="Quiz Website API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

SUBJECTS = {
    "Falsafa": "Falsafa.docx",
    "MT-V-A": "Mtuzilma.docx",
    "Dasturlash": "Dasturlash.docx",
    "Dinshunoslik": "Dinshunoslik.docx",
    "Ingliz tili-Di": "Ingliz2.docx",
    "Ingliz tili-KIN": "Ingliz.docx",
    "Fizika": "Fizika.docx"  
}
ENGLISH_PDF_PATH = "Ingliz_javoblar.pdf"
ANSWERS_JSON_PATH = "Fizika_answers.json"

@app.get("/api/subjects")
def get_subjects():
    return list(SUBJECTS.keys())

@app.get("/api/quiz/{subject}")
def get_quiz_questions(subject: str, count: str = "20"):
    if subject not in SUBJECTS:
        raise HTTPException(status_code=404, detail="Fan topilmadi")
    
    file_path = os.path.join(BASE_DIR, SUBJECTS[subject])
    if not os.path.exists(file_path):
         raise HTTPException(status_code=500, detail=f"{SUBJECTS[subject]} fayli serverda topilmadi")

    try:
        if subject == "Dasturlash":
            all_tests = get_quizzes_programming(file_path)
        elif subject == "Ingliz tili-KIN":
            pdf_full_path = os.path.join(BASE_DIR, ENGLISH_PDF_PATH)
            all_tests = get_quizzes_english_pdf_docx(file_path, pdf_full_path)
        elif subject == "Fizika":
            json_full_path = os.path.join(BASE_DIR, ANSWERS_JSON_PATH)
            all_tests = get_quizzes_by_letters(file_path, json_full_path)
        else:
            all_tests = get_quizzes(file_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Parser xatosi: {str(e)}")

    if not all_tests:
        return []

    # Agar "all" tanlansa hamma testlarni, aks holda so'ralgan miqdorni olamiz
    if count.lower() == "all":
        sample_size = len(all_tests)
    else:
        try:
            sample_size = min(int(count), len(all_tests))
        except ValueError:
            sample_size = min(20, len(all_tests))

    selected_tests = random.sample(all_tests, sample_size)
    
    formatted_tests = []
    for idx, test in enumerate(selected_tests):
        options = [str(opt) for opt in test["options"]]
        correct_text = options[test["correct"]]
        
        random.shuffle(options)
        new_correct_idx = options.index(correct_text)
        
        formatted_tests.append({
            "id": idx + 1,
            "question": test["question"],
            "options": options,
            "correct": new_correct_idx
        })
        
    return formatted_tests

@app.get("/")
def read_root():
    html_path = os.path.join(BASE_DIR, "index.html")
    if os.path.exists(html_path):
        return FileResponse(html_path)
    raise HTTPException(status_code=404, detail="index.html fayli topilmadi!")