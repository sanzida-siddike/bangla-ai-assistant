# Shikkha.ai 📚
**Teach smarter. Learn smarter. In Bangla.**

Shikkha.ai is a Bangla-first AI learning assistant for **students and teachers**.
Upload notes, books, slides or research papers (PDF, Word, PowerPoint, Excel, TXT) and:

🔗 Live app: https://bangla-ai-sanzida.streamlit.app

## Features

### 💬 চ্যাট (Chat)
- Ask questions about your files in Bangla or English
- RAG search (TF-IDF word + character n-grams) with page/slide sources for every answer
- Add many files at once with the **+** button

### 🎯 অনুশীলন (Practice, for students)
- **MCQ Quiz**: normal or university-admission level
- **Weak-topic finder**: tracks wrong answers by topic and makes a new quiz on weak topics
- **Viva practice**: AI asks oral questions, grades your answer out of 10 with feedback
- **Past-question analysis**: finds repeated topics and likely questions from old papers
- Choose a file and a page/slide range

### 📝 প্রশ্নপত্র (Question paper, for teachers)
- **SSC/HSC সৃজনশীল** (উদ্দীপক + ক খ গ ঘ)
- **Diploma (BTEB)**: editable pattern (final / class test)
- **University**: presets for National, Public and Private universities (Final, Mid, Incourse, Quiz/CT) and research papers,
  all editable (sections × questions × marks, "answer any N")
- **Auto-detect pattern** from an old question paper
- Research-paper questions: methodology, findings, critique, contribution
- Optional Bloom level + CO tags
- Download as **Word (.docx)**, with or without answers

## Tech
Python · Streamlit · Groq (`openai/gpt-oss-120b`) · scikit-learn TF-IDF RAG · pypdf · python-docx · python-pptx · openpyxl

## Run locally
```bash
py -m pip install -r requirements.txt
py -m streamlit run app.py
```
Put your key in `.streamlit/secrets.toml` (this file is never uploaded to GitHub):
```toml
GROQ_API_KEY = "your-key-here"
```

## Author
Made by [Sanzida Siddike](https://github.com/sanzida-siddike), Computer Instructor & M.Sc SE student (DIU).
