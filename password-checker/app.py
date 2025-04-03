import re
import random
import string
import streamlit as st

if "password_history" not in st.session_state:
    st.session_state.password_history = []

weak_passwords = ["password123", "123456", "qwerty", "123456789", "password", "12345", "123123", "admin", "letmein"]

def check_password_strength(password):
    score = 0
    feedback = []

    if len(password) >= 12:
        score += 3
    elif len(password) >= 8:
        score += 2
    else:
        feedback.append("❌ Password should be at least 12 characters long.")

    if re.search(r"[A-Z]", password) and re.search(r"[a-z]", password):
        score += 1
    else:
        feedback.append("❌ Include both uppercase and lowercase letters.")

    if re.search(r"\d", password):
        score += 1
    else:
        feedback.append("❌ Add at least one number (0-9).")

    if re.search(r"[!@#$%^&*]", password):
        score += 1
    else:
        feedback.append("❌ Include at least one special character (!@#$%^&*).")

    if password.lower() in weak_passwords:
        feedback.append("❌ This is a common weak password. Please choose another.")

    return score, feedback

def store_password(password, score):
    history = st.session_state.password_history

    if score == 7: 
        if any(password == old_pass for old_pass in history[-3:]):
            return "⚠️ This password is similar to one of your last 3 passwords. Consider changing it."
        
        history.append(password)
        if len(history) > 10:
            history.pop(0)   
        st.session_state.password_history = history
        return None  

    return None

def generate_strong_password():
    length = random.randint(12, 16)
    password = "".join(random.choices(string.ascii_letters + string.digits + "!@#$%^&*", k=length))
    return password

st.title("🔐 Password Checker")

password = st.text_input("Enter your password:", type="password")

if password:
    score, feedback = check_password_strength(password)
    warning = store_password(password, score)

    with st.container():
        if score <= 3:
            st.error("❌ Weak Password - Improve it using the suggestions below.")
            strength_label = "Weak"
            st.progress(score / 7)  
        elif score == 4 or score == 5:
            st.warning("⚠️ Moderate Password - Consider adding more security features.")
            strength_label = "Moderate"
            st.progress(score / 7)
        else:
            st.success("✅ Strong Password!")
            strength_label = "Strong"
            st.progress(score / 7)

        st.markdown(f"**Strength Level: {strength_label}**")

    if warning:
        st.warning(warning)

    with st.container():
        for msg in feedback:
            st.error(msg)

if st.button("Suggest a Strong Password"):
    suggested_password = generate_strong_password()
    st.markdown(f"Suggested Strong Password: **{suggested_password}**")

with st.expander("🔍 View Last 10 Passwords"):
    if st.session_state.password_history:
        st.markdown("**🔹 Your Recent Passwords:**")
        st.markdown("\n".join([f"- {pwd}" for pwd in st.session_state.password_history]))
    else:
        st.info("No strong passwords stored yet. Create one!")
