from flask import Flask, render_template, request, jsonify
import requests
import csv
import os
import re

app = Flask(__name__)

OLLAMA_API = "http://localhost:11434/api/chat"
RESPONSES_CSV = "responses.csv"
PROHIBIDAS_CSV = "prohibidas.csv"
SIMILARITY_THRESHOLD = 0.6

SYSTEM_PROMPT = {
    "role": "system",
    "content": (
        "Gobi, el asistente virtual oficial de la USICAMM. "
        "Responde con precisión y amabilidad las dudas relacionadas con la USICAMM."
    )
}

def load_prohibited_words():
    if not os.path.exists(PROHIBIDAS_CSV):
        return []
    with open(PROHIBIDAS_CSV, newline='', encoding='utf-8') as f:
        return [row[0].lower() for row in csv.reader(f)]

def contains_prohibited_word(text, prohibited_words):
    text_lower = text.lower()
    return any(palabra in text_lower for palabra in prohibited_words)

def load_responses():
    if not os.path.exists(RESPONSES_CSV):
        return []
    with open(RESPONSES_CSV, newline='', encoding='utf-8') as f:
        return list(csv.reader(f))

def save_response(question, answer):
    exists = os.path.exists(RESPONSES_CSV)
    with open(RESPONSES_CSV, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if not exists:
            writer.writerow(["question", "answer"])
        writer.writerow([question, answer])

def simple_similarity(a, b):
    a_words = set(re.findall(r'\w+', a.lower()))
    b_words = set(re.findall(r'\w+', b.lower()))
    if not a_words or not b_words:
        return 0.0
    inter = a_words.intersection(b_words)
    return len(inter) / max(len(a_words), len(b_words))

def find_similar_answer(question, responses):
    for q, a in responses:
        if simple_similarity(q, question) >= SIMILARITY_THRESHOLD:
            return a
    return None

def chat_with_ollama(user_prompt):
    payload = {
        "model": "deepseek-llm:7b",
        "messages": [
            SYSTEM_PROMPT,
            {"role": "user", "content": user_prompt}
        ],
        "stream": False
    }

    response = requests.post(OLLAMA_API, json=payload)
    if response.status_code == 200:
        return response.json().get("message", {}).get("content", "")
    else:
        return f"Error: {response.status_code} - {response.text}"

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/chat", methods=["POST"])
def chat():
    user_input = request.json.get("message")
    prohibited_words = load_prohibited_words()
    responses = load_responses()

    if contains_prohibited_word(user_input, prohibited_words):
        return jsonify({"answer": "Lo siento, no puedo responder esa pregunta."})

    answer = find_similar_answer(user_input, responses)
    if not answer:
        answer = chat_with_ollama(user_input)
        save_response(user_input, answer)

    return jsonify({"answer": answer})

if __name__ == "__main__":
    app.run(debug=True)
