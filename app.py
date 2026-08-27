import hashlib
import random
import re
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
import streamlit as st
from docx import Document
from supabase import create_client

# ============================================================

# STREAMLIT CONFIGURATION

# ============================================================

st.set_page_config(
page_title="NBME Study Quiz",
page_icon="🧠",
layout="wide",
)

# ============================================================

# FILE LOCATION

# ============================================================

# The Word document is stored in the same GitHub repository

# as this Python file.

WORD_FILE = (
Path(**file**).parent
/ "HBA - NBME Questions.docx"
)

# ============================================================

# SUPABASE CONNECTION

# ============================================================

@st.cache_resource
def get_supabase_client():
"""
Create one Supabase client for the Streamlit server.

```
IMPORTANT:
SUPABASE_SECRET_KEY must be stored in Streamlit Secrets.
Never put the secret key directly in this Python file.
"""

return create_client(
    st.secrets["SUPABASE_URL"],
    st.secrets["SUPABASE_SECRET_KEY"],
)
```

supabase = get_supabase_client()

# ============================================================

# TEXT CLEANING

# ============================================================

def clean_text(text):
"""
Clean up text extracted from Word.

```
Removes:
- non-breaking spaces
- zero-width characters
- markdown-style formatting
- excessive whitespace
"""

if text is None:
    return ""

text = text.replace("\u00a0", " ")
text = text.replace("\u200b", "")
text = text.replace("\u200c", "")
text = text.replace("\u200d", "")
text = text.replace("\ufeff", "")

text = text.replace("**", "")
text = text.replace("__", "")

text = text.replace("\t", " ")

return text.strip()
```

# ============================================================

# READ WORD DOCUMENT

# ============================================================

def read_word_as_text(file_path):
"""
Read the Word document and return its text.

```
Reads:
- normal paragraphs
- text inside tables
"""

document = Document(file_path)

chunks = []

# --------------------------------------------------------
# Normal paragraphs
# --------------------------------------------------------

for paragraph in document.paragraphs:

    text = paragraph.text

    if text:
        chunks.append(text)

# --------------------------------------------------------
# Tables
# --------------------------------------------------------

for table in document.tables:

    for row in table.rows:

        row_parts = []

        for cell in row.cells:

            cell_text = cell.text

            if cell_text:
                row_parts.append(cell_text)

        if row_parts:

            chunks.append(
                " | ".join(row_parts)
            )

full_text = "\n".join(chunks)

full_text = full_text.replace(
    "\r\n",
    "\n"
)

full_text = full_text.replace(
    "\r",
    "\n"
)

return full_text
```

# ============================================================

# NORMALIZE DOCUMENT

# ============================================================

def normalize_document_text(text):

```
lines = []

for raw_line in text.split("\n"):

    line = clean_text(raw_line)

    if line:
        lines.append(line)

return "\n".join(lines)
```

# ============================================================

# IDENTIFY QUESTION NUMBER

# ============================================================

def parse_question_number(line):

```
line = clean_text(line)

line = line.replace("*", "")

match = re.fullmatch(
    r"(\d+)\s*[\.\)]?",
    line
)

if match:

    return int(
        match.group(1)
    )

return None
```

# ============================================================

# IDENTIFY ANSWER CHOICE

# ============================================================

def parse_answer_choice(line):

```
line = clean_text(line)

line = line.replace("*", "")

match = re.match(
    r"^([A-E])\s*[\.\)]\s*(.+)$",
    line,
    flags=re.IGNORECASE
)

if not match:
    return None

letter = match.group(1).upper()

answer = match.group(2).strip()

return letter, answer
```

# ============================================================

# FIND SECTION

# ============================================================

def find_section(text, heading):

```
match = re.search(
    rf"(?im)^\s*{re.escape(heading)}\s*$",
    text
)

if match:
    return match

return None
```

# ============================================================

# PARSE QUESTIONS

# ============================================================

def parse_questions(question_text):

```
lines = question_text.split("\n")

questions = {}

current_number = None

current_question_lines = []

current_options = {}

def save_current_question():

    nonlocal current_number
    nonlocal current_question_lines
    nonlocal current_options

    if current_number is None:
        return

    question_text_combined = " ".join(
        current_question_lines
    ).strip()

    if (
        question_text_combined
        and len(current_options) >= 4
    ):

        questions[current_number] = {
            "question": question_text_combined,
            "options": dict(current_options)
        }

for raw_line in lines:

    line = clean_text(raw_line)

    if not line:
        continue

    # ----------------------------------------------------
    # New question
    # ----------------------------------------------------

    question_number = parse_question_number(
        line
    )

    if question_number is not None:

        save_current_question()

        current_number = question_number

        current_question_lines = []

        current_options = {}

        continue

    # ----------------------------------------------------
    # Answer option
    # ----------------------------------------------------

    option = parse_answer_choice(
        line
    )

    if (
        option is not None
        and current_number is not None
    ):

        letter, answer = option

        current_options[letter] = answer

        continue

    # ----------------------------------------------------
    # Regular question text
    # ----------------------------------------------------

    if current_number is not None:

        current_question_lines.append(
            line
        )

save_current_question()

return questions
```

# ============================================================

# FIND EXPLANATION SECTION

# ============================================================

def extract_explanation_section(
full_text,
answer_sheet_match
):

```
text_after_answer_sheet = full_text[
    answer_sheet_match.end():
]

heading_pattern = re.compile(
    r"(?im)^\s*(\d+)\.\s+(.+?)\s*$"
)

for match in heading_pattern.finditer(
    text_after_answer_sheet
):

    start = match.end()

    preview = text_after_answer_sheet[
        start:start + 2000
    ]

    if re.search(
        r"Correct\s+answer\s*:",
        preview,
        flags=re.IGNORECASE
    ):

        return text_after_answer_sheet[
            match.start():
        ]

raise ValueError(
    "Could not find the explanation section."
)
```

# ============================================================

# PARSE EXPLANATIONS

# ============================================================

def parse_explanations(explanation_text):

```
explanations = {}

heading_pattern = re.compile(
    r"(?im)^\s*(\d+)\.\s+(.+?)\s*$"
)

matches = list(
    heading_pattern.finditer(
        explanation_text
    )
)

for index, match in enumerate(matches):

    number = int(
        match.group(1)
    )

    if index + 1 < len(matches):

        end = matches[
            index + 1
        ].start()

    else:

        end = len(
            explanation_text
        )

    block = explanation_text[
        match.start():end
    ].strip()

    answer_match = re.search(
        r"Correct\s+answer\s*:\s*([A-E])",
        block,
        flags=re.IGNORECASE
    )

    if answer_match is None:
        continue

    correct_letter = answer_match.group(
        1
    ).upper()

    explanations[number] = {
        "correct_letter": correct_letter,
        "explanation": block
    }

return explanations
```

# ============================================================

# LOAD COMPLETE QUESTION BANK

# ============================================================

def load_question_bank(file_path):

```
if not file_path.exists():

    raise FileNotFoundError(
        "The Word document could not be found:\n\n"
        f"{file_path}"
    )

raw_text = read_word_as_text(
    file_path
)

full_text = normalize_document_text(
    raw_text
)

questions_match = find_section(
    full_text,
    "Questions"
)

answer_sheet_match = find_section(
    full_text,
    "Answer Sheet"
)

if questions_match is None:

    raise ValueError(
        "Could not find 'Questions' "
        "in the Word document."
    )

if answer_sheet_match is None:

    raise ValueError(
        "Could not find 'Answer Sheet' "
        "in the Word document."
    )

# --------------------------------------------------------
# QUESTION SECTION
# --------------------------------------------------------

question_section = full_text[
    questions_match.end():
    answer_sheet_match.start()
]

questions = parse_questions(
    question_section
)

# --------------------------------------------------------
# EXPLANATIONS
# --------------------------------------------------------

explanation_section = (
    extract_explanation_section(
        full_text,
        answer_sheet_match
    )
)

explanations = parse_explanations(
    explanation_section
)

# --------------------------------------------------------
# MATCH QUESTIONS + EXPLANATIONS
# --------------------------------------------------------

final_questions = []

for number in sorted(questions):

    if number not in explanations:

        continue

    question = questions[number]

    explanation = explanations[number]

    correct_letter = (
        explanation[
            "correct_letter"
        ]
    )

    if (
        correct_letter
        not in question["options"]
    ):

        continue

    final_questions.append({

        "number": number,

        "question": question[
            "question"
        ],

        "options": question[
            "options"
        ],

        "correct_letter": correct_letter,

        "correct_answer": question[
            "options"
        ][correct_letter],

        "explanation": explanation[
            "explanation"
        ]
    })

if not final_questions:

    raise ValueError(
        "No complete questions were found "
        "in the Word document."
    )

return final_questions
```

# ============================================================

# CACHE QUESTION BANK

# ============================================================

@st.cache_data
def load_cached_question_bank(
file_path,
modified_time
):
"""
Cache the parsed Word document.

```
modified_time ensures Streamlit reloads the document
when the Word file changes.
"""

return load_question_bank(
    Path(file_path)
)
```

# ============================================================

# PARTICIPANT ID

# ============================================================

def create_participant_id(
participant_code,
app_salt
):
"""
Convert a participant's private code into a
deterministic identifier.

```
The raw participant code is NOT stored.

The user can therefore return later and recover
their previous statistics by entering the same code.

IMPORTANT:
This is identification, not full authentication.
Users should choose a unique private code.
"""

value = (
    str(app_salt)
    + "|"
    + participant_code.strip()
)

return hashlib.sha256(
    value.encode("utf-8")
).hexdigest()
```

# ============================================================

# DATABASE FUNCTIONS

# ============================================================

def save_participant(
participant_id,
display_name
):
"""
Save/update a participant.
"""

```
response = (
    supabase
    .table("participants")
    .upsert(
        {
            "participant_id": participant_id,
            "display_name": display_name,
        },
        on_conflict="participant_id"
    )
    .execute()
)

return response
```

def save_answer(
participant_id,
display_name,
question,
selected_letter
):
"""
Save one answer attempt permanently in Supabase.
"""

```
is_correct = (
    selected_letter
    == question["correct_letter"]
)

record = {
    "participant_id": participant_id,
    "display_name": display_name,
    "question_number": question["number"],
    "selected_letter": selected_letter,
    "correct_letter": question["correct_letter"],
    "is_correct": is_correct,
    "answered_at": datetime.now(
        timezone.utc
    ).isoformat(),
}

(
    supabase
    .table("answer_logs")
    .insert(record)
    .execute()
)

return is_correct
```

# ============================================================

# DATABASE STATISTICS

# ============================================================

def get_overall_stats(
participant_id
):
"""
Get all answer attempts belonging to this participant.

```
Used only for internal statistics where needed.
"""

response = (
    supabase
    .table("answer_logs")
    .select(
        "question_number, "
        "selected_letter, "
        "correct_letter, "
        "is_correct, "
        "answered_at"
    )
    .eq(
        "participant_id",
        participant_id
    )
    .order(
        "answered_at",
        desc=False
    )
    .execute()
)

rows = response.data or []

total_attempts = len(rows)

total_correct = sum(
    1
    for row in rows
    if row.get("is_correct") is True
)

return rows, total_attempts, total_correct
```

def get_question_stats(
participant_id,
question_number
):
"""
Get historical statistics for ONE specific question.

```
Includes:

USER:
- attempts
- correct answers
- accuracy
- answer distribution

GLOBAL:
- attempts
- correct answers
- accuracy
"""

# ========================================================
# THIS USER'S STATISTICS FOR THIS QUESTION
# ========================================================

response = (
    supabase
    .table("answer_logs")
    .select(
        "selected_letter, "
        "is_correct"
    )
    .eq(
        "participant_id",
        participant_id
    )
    .eq(
        "question_number",
        question_number
    )
    .execute()
)

rows = response.data or []

attempts = len(rows)

correct = sum(
    1
    for row in rows
    if row.get("is_correct") is True
)

answer_counts = {
    "A": 0,
    "B": 0,
    "C": 0,
    "D": 0,
    "E": 0,
}

for row in rows:

    letter = row.get(
        "selected_letter"
    )

    if letter in answer_counts:

        answer_counts[letter] += 1

user_accuracy = (
    correct
    / attempts
    * 100
    if attempts
    else 0
)

# ========================================================
# GLOBAL STATISTICS FOR THIS QUESTION
# ========================================================

global_response = (
    supabase
    .table("answer_logs")
    .select(
        "*",
        count="exact",
        head=True
    )
    .eq(
        "question_number",
        question_number
    )
    .execute()
)

global_attempts = (
    global_response.count or 0
)

# --------------------------------------------------------
# Global correct attempts
# --------------------------------------------------------

global_correct_response = (
    supabase
    .table("answer_logs")
    .select(
        "*",
        count="exact",
        head=True
    )
    .eq(
        "question_number",
        question_number
    )
    .eq(
        "is_correct",
        True
    )
    .execute()
)

global_correct = (
    global_correct_response.count or 0
)

global_accuracy = (
    global_correct
    / global_attempts
    * 100
    if global_attempts
    else 0
)

return {
    "attempts": attempts,
    "correct": correct,
    "accuracy": user_accuracy,
    "answers": answer_counts,
    "global_attempts": global_attempts,
    "global_correct": global_correct,
    "global_accuracy": global_accuracy,
}
```

# ============================================================

# ADMIN FUNCTIONS

# ============================================================

def get_all_answers():

```
response = (
    supabase
    .table("answer_logs")
    .select("*")
    .order(
        "answered_at",
        desc=True
    )
    .execute()
)

return response.data or []
```

# ============================================================

# INITIALIZE SESSION STATE

# ============================================================

def initialize_quiz_state():

```
defaults = {

    "quiz_started": False,

    "participant_id": None,

    "display_name": "",

    "questions": [],

    "current_index": 0,

    "current_correct": False,

    "quiz_complete": False,

    "last_feedback": "",

    "last_feedback_type": "",

    "show_explanation": False,

    "flash_type": None,
}

for key, value in defaults.items():

    if key not in st.session_state:

        st.session_state[key] = value
```

initialize_quiz_state()

# ============================================================

# ANSWER FLASH

# ============================================================

if st.session_state.get("flash_type"):

```
if st.session_state.flash_type == "correct":

    flash_color = "rgba(0, 200, 0, 0.40)"

else:

    flash_color = "rgba(255, 0, 0, 0.40)"

st.markdown(
    f"""
    <style>

    @keyframes answerFlash {{
        0% {{
            opacity: 0;
        }}

        20% {{
            opacity: 1;
        }}

        60% {{
            opacity: 1;
        }}

        100% {{
            opacity: 0;
        }}
    }}

    .answer-flash {{
        position: fixed;
        top: 0;
        left: 0;
        width: 100vw;
        height: 100vh;
        background: {flash_color};
        z-index: 999999;
        pointer-events: none;
        animation: answerFlash 0.65s ease-out forwards;
    }}

    </style>

    <div class="answer-flash"></div>
    """,
    unsafe_allow_html=True
)

# Clear the flash so it occurs only once.
st.session_state.flash_type = None
```

# ============================================================

# LOAD QUESTIONS

# ============================================================

try:

```
modified_time = (
    WORD_FILE.stat().st_mtime_ns
)

QUESTIONS = load_cached_question_bank(
    str(WORD_FILE),
    modified_time
)
```

except Exception as error:

```
st.error(
    "There was a problem loading the "
    "question bank."
)

st.exception(error)

st.stop()
```

# ============================================================

# SIDEBAR

# ============================================================

with st.sidebar:

```
st.title("🧠 NBME Study Quiz")

st.write(
    f"**Questions loaded:** {len(QUESTIONS)}"
)

st.divider()

if st.session_state.quiz_started:

    st.write(
        f"**Participant:** "
        f"{st.session_state.display_name}"
    )

    if st.button(
        "Restart Quiz",
        use_container_width=True
    ):

        st.session_state.questions = []

        st.session_state.current_index = 0

        st.session_state.current_correct = False

        st.session_state.quiz_complete = False

        st.session_state.last_feedback = ""

        st.session_state.last_feedback_type = ""

        st.session_state.show_explanation = False

        st.session_state.flash_type = None

        st.rerun()
```

# ============================================================

# TITLE

# ============================================================

st.title("🧠 NBME Study Quiz")

st.caption(
"Choose an answer. Incorrect answers can be "
"retried until the correct answer is selected."
)

# ============================================================

# LOGIN / START PAGE

# ============================================================

if not st.session_state.quiz_started:

```
st.subheader(
    "Log in to start"
)

username = st.text_input(
    "Username",
    placeholder="Enter your username"
)

password = st.text_input(
    "Password",
    type="password",
    placeholder="Enter your password"
)

if st.button(
    "Log In & Start Quiz",
    type="primary",
    use_container_width=True
):

    username = username.strip()

    # ----------------------------------------------------
    # CHECK USERNAME
    # ----------------------------------------------------

    users = st.secrets["users"]

    if username not in users:

        st.error(
            "Invalid username or password."
        )

        st.stop()

    # ----------------------------------------------------
    # CHECK PASSWORD
    # ----------------------------------------------------

    expected_password = users[username]

    if password != expected_password:

        st.error(
            "Invalid username or password."
        )

        st.stop()

    # ----------------------------------------------------
    # CREATE PARTICIPANT ID
    # ----------------------------------------------------

    participant_id = create_participant_id(
        username,
        st.secrets["PARTICIPANT_SALT"]
    )

    # ----------------------------------------------------
    # SAVE PARTICIPANT
    # ----------------------------------------------------

    try:

        save_participant(
            participant_id,
            username
        )

    except Exception as error:

        st.error(
            "Could not connect to the database."
        )

        st.exception(error)

        st.stop()

    # ----------------------------------------------------
    # RANDOMIZE QUESTIONS
    # ----------------------------------------------------

    question_copy = QUESTIONS.copy()

    random.shuffle(
        question_copy
    )

    # ----------------------------------------------------
    # STORE LOGIN INFORMATION
    # ----------------------------------------------------

    st.session_state.quiz_started = True

    st.session_state.participant_id = (
        participant_id
    )

    st.session_state.display_name = (
        username
    )

    st.session_state.questions = (
        question_copy
    )

    st.session_state.current_index = 0

    st.session_state.current_correct = False

    st.session_state.quiz_complete = False

    st.session_state.last_feedback = ""

    st.session_state.last_feedback_type = ""

    st.session_state.show_explanation = False

    st.session_state.flash_type = None

    st.rerun()

st.divider()

st.write(
    "Please use the username and password "
    "provided to you."
)

st.stop()
```

# ============================================================

# ADMIN PANEL

# ============================================================

st.markdown(
""" <details> <summary><strong>Administrator</strong></summary>
""",
unsafe_allow_html=True
)

admin_password = st.text_input(
"Administrator password",
type="password"
)

if admin_password:

```
expected_password = st.secrets[
    "ADMIN_PASSWORD"
]

if admin_password == expected_password:

    st.success(
        "Administrator access granted."
    )

    try:

        all_answers = get_all_answers()

        if all_answers:

            admin_df = pd.DataFrame(
                all_answers
            )

            st.write(
                f"Total recorded answer attempts: "
                f"{len(admin_df):,}"
            )

            st.dataframe(
                admin_df,
                use_container_width=True
            )

            csv_data = admin_df.to_csv(
                index=False
            ).encode("utf-8")

            st.download_button(
                label="Download CSV",
                data=csv_data,
                file_name=(
                    "nbme_quiz_results.csv"
                ),
                mime="text/csv"
            )

        else:

            st.info(
                "No answer records yet."
            )

    except Exception as error:

        st.error(
            "Could not load administrator data."
        )

        st.exception(error)

else:

    st.error(
        "Incorrect administrator password."
    )
```

st.markdown(
"</details>",
unsafe_allow_html=True
)

# ============================================================

# CURRENT QUIZ

# ============================================================

questions = st.session_state.questions

current_index = st.session_state.current_index

# ============================================================

# FINISHED

# ============================================================

if st.session_state.quiz_complete:

```
st.success(
    "🎉 Quiz complete!"
)

total_questions = len(
    questions
)

st.write(
    f"You completed all **{total_questions}** "
    "questions."
)

st.divider()

# --------------------------------------------------------
# LAST QUESTION / CURRENT QUESTION STATISTICS
# --------------------------------------------------------

st.info(
    "Your historical statistics are shown for each "
    "individual question while you are answering it."
)

if st.button(
    "Start another randomized quiz",
    type="primary",
    use_container_width=True
):

    new_questions = QUESTIONS.copy()

    random.shuffle(
        new_questions
    )

    st.session_state.questions = (
        new_questions
    )

    st.session_state.current_index = 0

    st.session_state.current_correct = False

    st.session_state.quiz_complete = False

    st.session_state.last_feedback = ""

    st.session_state.last_feedback_type = ""

    st.session_state.show_explanation = False

    st.session_state.flash_type = None

    st.rerun()

st.stop()
```

# ============================================================

# CURRENT QUESTION

# ============================================================

question = questions[current_index]

question_number = question["number"]

total_questions = len(
questions
)

# ============================================================

# PROGRESS

# ============================================================

progress = (
(current_index + 1)
/ total_questions
)

st.progress(
progress
)

st.write(
f"### Question {current_index + 1} "
f"of {total_questions}"
)

st.caption(
f"Document question number: "
f"{question_number}"
)

# ============================================================

# HISTORICAL QUESTION STATISTICS

# ============================================================

try:

```
historical_stats = get_question_stats(
    st.session_state.participant_id,
    question_number
)
```

except Exception:

```
historical_stats = {
    "attempts": 0,
    "correct": 0,
    "accuracy": 0,
    "answers": {
        "A": 0,
        "B": 0,
        "C": 0,
        "D": 0,
        "E": 0,
    },
    "global_attempts": 0,
    "global_correct": 0,
    "global_accuracy": 0,
}
```

# ============================================================

# USER STATISTICS FOR CURRENT QUESTION

# ============================================================

historical_attempts = (
historical_stats["attempts"]
)

historical_correct = (
historical_stats["correct"]
)

historical_accuracy = (
historical_stats["accuracy"]
)

# ============================================================

# GLOBAL STATISTICS FOR CURRENT QUESTION

# ============================================================

global_attempts = (
historical_stats["global_attempts"]
)

global_correct = (
historical_stats["global_correct"]
)

global_accuracy = (
historical_stats["global_accuracy"]
)

# ============================================================

# HISTORICAL DISPLAY

# ============================================================

if historical_attempts > 0:

```
distribution_parts = []

for letter in [
    "A",
    "B",
    "C",
    "D",
    "E"
]:

    count = (
        historical_stats[
            "answers"
        ].get(letter, 0)
    )

    percentage = (
        count
        / historical_attempts
        * 100
    )

    distribution_parts.append(
        f"{letter}: {percentage:.0f}%"
    )

if global_attempts > 0:

    st.caption(
        f"Your previous attempts: "
        f"**{historical_attempts}** | "
        f"Your historical accuracy: "
        f"**{historical_accuracy:.0f}%** | "
        f"Global historical accuracy: "
        f"**{global_accuracy:.0f}%** | "
        f"Previously selected: "
        f"{' '.join(distribution_parts)}"
    )

else:

    st.caption(
        f"Your previous attempts: "
        f"**{historical_attempts}** | "
        f"Your historical accuracy: "
        f"**{historical_accuracy:.0f}%** | "
        f"Global historical accuracy: "
        f"No previous attempts"
    )
```

else:

```
if global_attempts > 0:

    st.caption(
        "Your previous attempts: **none** | "
        f"Global historical accuracy: "
        f"**{global_accuracy:.0f}%**"
    )

else:

    st.caption(
        "Your previous attempts: none | "
        "Global historical accuracy: no attempts yet"
    )
```

# ============================================================

# QUESTION

# ============================================================

st.markdown(
f"### {question['question']}"
)

# ============================================================

# ANSWER BUTTONS

# ============================================================

for letter in [
"A",
"B",
"C",
"D",
"E"
]:

```
if letter not in question["options"]:
    continue

option_text = (
    question[
        "options"
    ][letter]
)

button_text = (
    f"{letter}. {option_text}"
)

# --------------------------------------------------------
# CORRECT / WRONG DISPLAY
# --------------------------------------------------------

if st.session_state.current_correct:

    if letter == question["correct_letter"]:

        st.success(
            button_text
        )

    else:

        st.button(
            button_text,
            disabled=True,
            use_container_width=True,
            key=f"disabled_{current_index}_{letter}"
        )

    continue

# --------------------------------------------------------
# NORMAL BUTTON
# --------------------------------------------------------

if st.button(
    button_text,
    use_container_width=True,
    key=f"question_{current_index}_{letter}"
):

    try:

        is_correct = save_answer(
            st.session_state.participant_id,
            st.session_state.display_name,
            question,
            letter
        )

    except Exception as error:

        st.error(
            "Your answer could not be saved."
        )

        st.exception(error)

        st.stop()

    # ----------------------------------------------------
    # CORRECT ANSWER
    # ----------------------------------------------------

    if is_correct:

        st.session_state.current_correct = True

        st.session_state.last_feedback = (
            "✓ Correct!"
        )

        st.session_state.last_feedback_type = (
            "correct"
        )

        st.session_state.show_explanation = True

        # Green screen flash
        st.session_state.flash_type = (
            "correct"
        )

    # ----------------------------------------------------
    # INCORRECT ANSWER
    # ----------------------------------------------------

    else:

        st.session_state.last_feedback = (
            "✗ Incorrect — try again."
        )

        st.session_state.last_feedback_type = (
            "incorrect"
        )

        # Red screen flash
        st.session_state.flash_type = (
            "incorrect"
        )

    st.rerun()
```

# ============================================================

# FEEDBACK

# ============================================================

if st.session_state.last_feedback:

```
if (
    st.session_state.last_feedback_type
    == "correct"
):

    st.success(
        st.session_state.last_feedback
    )

elif (
    st.session_state.last_feedback_type
    == "incorrect"
):

    st.error(
        st.session_state.last_feedback
    )
```

# ============================================================

# EXPLANATION

# ============================================================

if st.session_state.show_explanation:

```
st.divider()

st.subheader(
    "Explanation"
)

st.info(
    question["explanation"]
)

st.write(
    f"**Correct answer: "
    f"{question['correct_letter']}. "
    f"{question['correct_answer']}**"
)

st.divider()

# --------------------------------------------------------
# NEXT QUESTION
# --------------------------------------------------------

if current_index == total_questions - 1:

    next_button_text = (
        "Finish Quiz"
    )

else:

    next_button_text = (
        "Next Question →"
    )

if st.button(
    next_button_text,
    type="primary",
    use_container_width=True
):

    if current_index == total_questions - 1:

        st.session_state.quiz_complete = True

    else:

        st.session_state.current_index += 1

    st.session_state.current_correct = False

    st.session_state.last_feedback = ""

    st.session_state.last_feedback_type = ""

    st.session_state.show_explanation = False

    st.session_state.flash_type = None

    st.rerun()
```
