#!/usr/bin/env python3
import os
import random
from datetime import datetime
from gtts import gTTS
from flask import Flask, request, jsonify, send_from_directory, render_template_string
from flask_cors import CORS

# -------------------- CONFIG --------------------
app = Flask(__name__)
CORS(app)

AUDIO_DIR = "audio_outputs"
STATIC_DIR = "static"
os.makedirs(AUDIO_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)


# -------------------- INTERVIEW AGENT --------------------
class InterviewAgent:
    def __init__(self):
        self.questions = [
            {"q": "How do you use VLOOKUP in Excel?", "keywords": ["vlookup", "="],
             "answer": "VLOOKUP searches a value in the first column and returns a value in the same row from another column. Example: =VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])."},
            {"q": "What is the difference between absolute and relative references?", "keywords": ["$", "absolute", "relative"],
             "answer": "Relative references change when copied; absolute references ($A$1) remain fixed."},
            {"q": "Explain Pivot Tables and give an example use case.", "keywords": ["pivot", "summarize", "aggregate"],
             "answer": "Pivot Tables summarize and analyze large datasets. Example: total sales per region."},
            {"q": "How would you find duplicates in a dataset?", "keywords": ["duplicate", "conditional formatting", "remove duplicates"],
             "answer": "Use Conditional Formatting → Highlight Duplicate Values, or Remove Duplicates tool."},
            {"q": "What is the purpose of conditional formatting?", "keywords": ["highlight", "color", "formatting"],
             "answer": "Apply formatting automatically based on rules, e.g., highlight sales above $1000."},
            {"q": "Explain the use of IF, AND, OR functions in Excel.", "keywords": ["if", "and", "or"],
             "answer": "IF performs conditional logic; AND checks multiple conditions; OR checks if any is true."},
            {"q": "What is the difference between COUNT, COUNTA, and COUNTIF?", "keywords": ["count", "counta", "countif"],
             "answer": "COUNT counts numbers, COUNTA counts non-empty cells, COUNTIF counts cells matching a condition."},
            {"q": "How do you create a drop-down list in Excel?", "keywords": ["data validation", "drop-down"],
             "answer": "Data → Data Validation → Allow: List → enter values."},
            {"q": "What are named ranges and why are they useful?", "keywords": ["named range", "reference"],
             "answer": "Named ranges assign a name to a cell/range for easier formula reference."},
            {"q": "Explain the use of INDEX and MATCH in Excel.", "keywords": ["index", "match"],
             "answer": "INDEX returns value by row/column; MATCH returns position. Together: flexible VLOOKUP alternative."}
        ]
        self.index = 0
        self.answers = []

    def reset(self):
        self.index = 0
        self.answers = []

    def intro(self):
        text = ("Hello, I’ll be your Excel mock interviewer today. "
                "I’ll ask you a series of questions to test your Excel knowledge. "
                "Please answer carefully. Let’s begin!")
        audio_file = self._speech(text)
        return text, audio_file

    def next_question(self):
        if self.index < len(self.questions):
            q_data = self.questions[self.index]
            audio_file = self._speech(q_data["q"])
            return True, {"index": self.index, "question": q_data["q"], "audio": audio_file}
        else:
            farewell = "Thank you for attending the interview. Best of luck with your preparation!"
            audio_file = self._speech(farewell)
            return False, {"message": farewell, "audio": audio_file}

    def submit_answer(self, ans_text):
        if self.index < len(self.questions):
            q_data = self.questions[self.index]
            score, feedback = self._evaluate(ans_text, q_data["keywords"], q_data["answer"])
            self.answers.append({"q": q_data["q"], "a": ans_text, "score": score, "feedback": feedback})
            self.index += 1
            audio_file = self._speech(feedback)
            return True, {"text": feedback, "audio": audio_file}
        return False, {"message": "No active question"}

    def _evaluate(self, ans, keywords, correct_answer):
        ans_lower = (ans or "").lower()
        matched = [k for k in keywords if k in ans_lower]
        score = (len(matched) / len(keywords)) if keywords else 0
        feedback = f"You mentioned {', '.join(matched) if matched else 'none of the keywords'}. "
        if score == 1:
            feedback += "Excellent!"
        elif score > 0:
            feedback += f"Partial. Correct: {correct_answer}"
        else:
            feedback = f"Missed key points. Correct answer: {correct_answer}"
        return round(score, 2), feedback

    def get_report(self):
        total = len(self.answers)
        score = sum(a["score"] for a in self.answers)
        return {"total": total, "score": round(score, 2), "answers": self.answers}

    def _speech(self, text):
        fname = f"ans_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{random.randint(0,9999)}.mp3"
        path = os.path.join(AUDIO_DIR, fname)
        try:
            gTTS(text=text, lang='en', slow=False).save(path)
            return fname
        except Exception as e:
            print("gTTS error:", e)
            return None


agent = InterviewAgent()


# -------------------- HTML PAGES --------------------
DASHBOARD_HTML = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>AI-Powered Excel Mock Interviewer</title>
  <style>
    body {
      margin:0; height:100vh;
      background:url('/static/photo1.png') no-repeat center center fixed;
      background-size:cover;
      display:flex; flex-direction:column;
      justify-content:center; align-items:center;
      color:white; text-shadow:1px 1px 5px black;
      font-family:Segoe UI, Arial;
    }
    h1 { font-size:2.5rem; margin-bottom:20px; }
    .btn {
      padding:15px 40px; font-size:1.2rem;
      border:none; border-radius:10px;
      background:#4CAF50; color:white; cursor:pointer;
    }
    .btn:hover { background:#45a049; }
  </style>
</head>
<body>
  <h1>AI-Powered Excel Mock Interviewer</h1>
  <button class="btn" onclick="location.href='/interview'">Start Interview</button>
</body>
</html>
"""

INTERVIEW_HTML = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>Excel Interview</title>
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <style>
    body {
      margin:0;
      height:100vh;
      background:url('/static/photo1.png') no-repeat center center fixed;
      background-size:cover;
      display:flex;
      justify-content:center;
      align-items:center;
      font-family:Segoe UI, Arial, sans-serif;
      color:white;
      text-shadow:1px 1px 4px black;
    }
    .card {
      background: rgba(0,0,0,0.7);
      padding: 30px;
      border-radius: 20px;
      max-width: 800px;
      width: 90%;
      text-align: center;
      box-shadow: 0 6px 15px rgba(0,0,0,0.5);
    }
    h1 { margin-top:0; font-size:2rem; margin-bottom:20px; }
    #question { margin:20px 0; font-size:1.3rem; }
    textarea {
      width:100%; height:80px; padding:10px; margin:10px 0;
      border-radius:8px; border:none; font-size:1rem;
    }
    button {
      padding:12px 24px; margin:8px;
      border:none; border-radius:10px;
      font-size:1rem;
      background:#4CAF50; color:white; cursor:pointer;
      transition: background 0.3s;
    }
    button:hover { background:#45a049; }
    .secondary { background:#2563eb; }
    .secondary:hover { background:#1e4fd1; }
    audio { margin-top:15px; width: 100%; }
    #log { margin-top:15px; min-height:30px; }
  </style>
</head>
<body>
  <div class="card">
    <h1>Excel Mock Interview</h1>
    <div id="question">Press Start to begin.</div>
    <button id="startBtn" class="secondary">Start</button>
    <button id="nextBtn" disabled>Next Question</button>
    <div class="options">
      <button id="speakBtn">🎙️ Speak Answer</button>
      <br>
      <textarea id="textAnswer" placeholder="Or type your answer here..."></textarea>
      <button id="submitTextBtn">Submit Text Answer</button>
    </div>
    <button id="finishBtn" disabled>Submit Interview</button>
    <button id="reportBtn" style="display:none">View Report</button>
    <br>
    <button onclick="location.href='/'" class="secondary">⬅ Back to Dashboard</button>
    <audio id="audioPlayer" controls style="display:none"></audio>
    <audio id="audioFeedback" controls style="display:none"></audio>
    <div id="log"></div>
  </div>

<script>
const startBtn=document.getElementById('startBtn');
const nextBtn=document.getElementById('nextBtn');
const speakBtn=document.getElementById('speakBtn');
const submitTextBtn=document.getElementById('submitTextBtn');
const textAnswer=document.getElementById('textAnswer');
const finishBtn=document.getElementById('finishBtn');
const reportBtn=document.getElementById('reportBtn');
const qBox=document.getElementById('question');
const log=document.getElementById('log');
const audioPlayer=document.getElementById('audioPlayer');
const audioFeedback=document.getElementById('audioFeedback');

let recognition=null, listening=false;
if('webkitSpeechRecognition' in window || 'SpeechRecognition' in window){
  const SR=window.SpeechRecognition||window.webkitSpeechRecognition;
  recognition=new SR();
  recognition.lang='en-US';
  recognition.onstart=()=>{listening=true;speakBtn.innerText="Stop Listening";};
  recognition.onend=()=>{listening=false;speakBtn.innerText="🎙️ Speak Answer";};
  recognition.onresult=async (e)=>{
    const transcript=e.results[0][0].transcript;
    log.innerText="You: "+transcript;
    await submitAnswer(transcript);
  };
}

function playAudio(url){
  if(url){
    audioPlayer.src=url;
    audioPlayer.style.display='block';
    audioPlayer.play().catch(()=>{});
  }
}

async function submitAnswer(answer){
  const res=await fetch('/api/submit_answer',{
    method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({answer})
  });
  const d=await res.json();
  if(d.text){
    qBox.innerText=d.text;
    if(d.audio_url){
      audioFeedback.src=d.audio_url;
      audioFeedback.style.display='block';
      audioFeedback.play().catch(()=>{});
    }
    nextBtn.disabled=false;
    finishBtn.disabled=false;
  }
}

startBtn.onclick=async()=>{
  const r=await fetch('/api/intro');const d=await r.json();
  qBox.innerText=d.text; if(d.audio_url) playAudio(d.audio_url);
  nextBtn.disabled=false;
};

nextBtn.onclick=nextQuestion;
async function nextQuestion(){
  const r=await fetch('/api/next_question');const d=await r.json();
  if(d.ok){ qBox.innerText=d.payload.question; playAudio(d.payload.audio_url);}
  else { qBox.innerText=d.payload.message; playAudio(d.payload.audio_url); nextBtn.disabled=true; finishBtn.disabled=false;}
}

speakBtn.onclick=()=>{ if(!recognition){alert("SpeechRecognition not supported");return;} if(listening) recognition.stop(); else recognition.start(); };
submitTextBtn.onclick=()=>{ if(textAnswer.value.trim()!=="") submitAnswer(textAnswer.value.trim()); };

finishBtn.onclick=()=>{ qBox.innerText="Interview finished! Click 'View Report' to see results."; reportBtn.style.display="inline-block"; };
reportBtn.onclick=async()=>{ const r=await fetch('/api/report'); const d=await r.json(); log.innerText="Report: Score "+d.score+" out of "+d.total+" questions."; };
</script>
</body>
</html>
"""


# -------------------- ROUTES --------------------
@app.route("/")
def dashboard():
    return render_template_string(DASHBOARD_HTML)

@app.route("/interview")
def interview():
    agent.reset()
    return render_template_string(INTERVIEW_HTML)

@app.route("/api/intro")
def api_intro():
    text, audio = agent.intro()
    return jsonify({"text": text, "audio_url": f"/audio/{audio}" if audio else None})

@app.route("/api/next_question")
def api_next_question():
    ok, payload = agent.next_question()
    if ok:
        return jsonify({"ok": True, "payload": {"index": payload["index"], "question": payload["question"], "audio_url": f"/audio/{payload['audio']}"}})
    return jsonify({"ok": False, "payload": {"message": payload["message"], "audio_url": f"/audio/{payload['audio']}"}})

@app.route("/api/submit_answer", methods=["POST"])
def api_submit_answer():
    data=request.get_json() or {}
    ans=data.get("answer","")
    ok,payload=agent.submit_answer(ans)
    if ok:
        return jsonify({"text":payload["text"],"audio_url":f"/audio/{payload['audio']}"})
    return jsonify({"error":payload["message"]}),400

@app.route("/api/report")
def api_report():
    return jsonify(agent.get_report())

@app.route("/audio/<path:filename>")
def audio(filename):
    return send_from_directory(AUDIO_DIR, filename)


# -------------------- RUN --------------------
if __name__=="__main__":
    app.run(host="0.0.0.0",port=5000,debug=True)
