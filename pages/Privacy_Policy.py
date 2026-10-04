import streamlit as st

st.set_page_config(
    page_title="Privacy Policy — NBME Study Quiz",
    page_icon="🔒",
)

st.title("Privacy Policy")

st.caption("NBME Study Quiz")

st.write(
    "This Privacy Policy explains what information is collected and "
    "how it is used when you use NBME Study Quiz."
)

st.header("1. Information we collect")

st.write(
    "When you sign in with Google, NBME Study Quiz receives basic "
    "information provided by Google for authentication, which may "
    "include your Google account identifier, name, and email address."
)

st.write(
    "The application also records information about your use of the "
    "quiz, including questions attempted, selected answers, whether "
    "answers were correct, question-bank progress, and the time an "
    "answer was submitted."
)

st.header("2. How we use your information")

st.write(
    "Your information is used to authenticate you, save your quiz "
    "progress across sessions, record quiz attempts, calculate "
    "individual and aggregate question statistics, and provide the "
    "functionality of the NBME Study Quiz."
)

st.header("3. Information sharing")

st.write(
    "Your information is not sold or rented to third parties."
)

st.write(
    "Authentication is provided through Google, and application "
    "data is stored using Supabase. These services process data as "
    "necessary to provide their respective services."
)

st.header("4. Data security")

st.write(
    "Reasonable technical measures are used to protect information "
    "stored by the application. However, no internet-based service "
    "can guarantee absolute security."
)

st.header("5. Data retention and deletion")

st.write(
    "Quiz and account information may be retained while the "
    "application is operating in order to maintain progress and "
    "historical statistics."
)

st.write(
    "If you would like your account information and associated quiz "
    "data deleted, please contact the application administrator."
)

st.header("6. Google authentication")

st.write(
    "NBME Study Quiz uses Google authentication to allow users to "
    "sign in. The application does not receive or store your Google "
    "password."
)

st.header("7. Changes to this Privacy Policy")

st.write(
    "This Privacy Policy may be updated from time to time. Any "
    "changes will be reflected on this page."
)

st.header("8. Contact")

st.write(
    "For questions about this Privacy Policy or requests concerning "
    "your data, please contact the application administrator."
)

st.divider()

st.caption("Last updated: October 3, 2026")
