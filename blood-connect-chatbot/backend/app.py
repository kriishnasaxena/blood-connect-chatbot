from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import openpyxl
import os

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class Message(BaseModel):
    message: str

# File paths
EXCEL_FILES = {
    "donation": "donations.xlsx",
    "request": "requests.xlsx",
    "camp": "camps.xlsx",
}

conversation_state = {
    "mode": None,
    "step": None,
    "data": {},
}

# Initialize Excel files
def init_excel(file_path, headers):
    if not os.path.exists(file_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(headers)
        wb.save(file_path)

init_excel(EXCEL_FILES["donation"], ["Name", "Age", "Blood Group", "Phone", "Location", "Last Donated"])
init_excel(EXCEL_FILES["request"], ["Name", "Age", "Blood Group", "Phone", "Location", "Urgency"])
init_excel(EXCEL_FILES["camp"], ["Organizer Name", "Phone", "Email", "Location", "Preferred Date", "Expected Donors"])

def save_to_excel(file_key, row):
    try:
        path = EXCEL_FILES[file_key]
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        ws.append(row)
        wb.save(path)
        print(f"[✔] Saved to {file_key}: {row}")
    except Exception as e:
        print(f"[✖] Error saving to {file_key}: {e}")

@app.post("/chat")
async def chat(msg: Message):
    user_msg = msg.message.strip()
    state = conversation_state
    data = state["data"]

    # START
    if not state["mode"]:
        if "donate" in user_msg.lower():
            state.update({"mode": "donation", "step": "name"})
            return {"reply": "Great! What's your full name?"}
        elif "request" in user_msg.lower():
            state.update({"mode": "request", "step": "name"})
            return {"reply": "Sure! What's your full name?"}
        elif "camp" in user_msg.lower():
            state.update({"mode": "camp", "step": "organizer"})
            return {"reply": "Awesome! Who's the camp organizer?"}
        else:
            return {
                "reply": "Hi! I’m the Blood Connect Assistant. What would you like to do?",
                "quick_replies": ["Donate Blood", "Request Blood", "Organize Camp"]
            }

    # DONATION
    if state["mode"] == "donation":
        if state["step"] == "name":
            data["name"] = user_msg
            state["step"] = "age"
            return {"reply": "What’s your age?"}
        if state["step"] == "age":
            if not user_msg.isdigit() or int(user_msg) < 18:
                return {"reply": "You must be 18+ to donate. Please enter a valid age."}
            data["age"] = user_msg
            state["step"] = "blood"
            return {"reply": "What is your blood group? (e.g., A+, B-)"}
        if state["step"] == "blood":
            data["blood"] = user_msg.upper()
            state["step"] = "phone"
            return {"reply": "Enter your 10-digit phone number."}
        if state["step"] == "phone":
            if not user_msg.isdigit() or len(user_msg) != 10:
                return {"reply": "Invalid phone number. Enter a 10-digit number."}
            data["phone"] = user_msg
            state["step"] = "location"
            return {"reply": "Where are you located?"}
        if state["step"] == "location":
            data["location"] = user_msg
            state["step"] = "last_donated"
            return {"reply": "Have you donated blood in the last 3 months? (Yes/No)"}
        if state["step"] == "last_donated":
            data["last_donated"] = user_msg
            save_to_excel("donation", [
                data["name"], data["age"], data["blood"], data["phone"], data["location"], data["last_donated"]
            ])
            state.update({"mode": None, "step": None, "data": {}})
            return {"reply": "Thanks for registering as a donor! We'll contact you when needed."}

    # REQUEST
    if state["mode"] == "request":
        if state["step"] == "name":
            data["name"] = user_msg
            state["step"] = "age"
            return {"reply": "What's your age?"}
        if state["step"] == "age":
            if not user_msg.isdigit():
                return {"reply": "Enter a valid age."}
            data["age"] = user_msg
            state["step"] = "blood"
            return {"reply": "Required blood group?"}
        if state["step"] == "blood":
            data["blood"] = user_msg.upper()
            state["step"] = "phone"
            return {"reply": "Enter your 10-digit phone number."}
        if state["step"] == "phone":
            if not user_msg.isdigit() or len(user_msg) != 10:
                return {"reply": "Invalid phone number. Enter a 10-digit number."}
            data["phone"] = user_msg
            state["step"] = "location"
            return {"reply": "Where do you need the blood?"}
        if state["step"] == "location":
            data["location"] = user_msg
            state["step"] = "urgency"
            return {"reply": "Is it urgent? (Yes/No)"}
        if state["step"] == "urgency":
            data["urgency"] = user_msg
            save_to_excel("request", [
                data["name"], data["age"], data["blood"], data["phone"], data["location"], data["urgency"]
            ])
            state.update({"mode": None, "step": None, "data": {}})
            return {"reply": "Request received. We’ll try to match you with available donors."}

    # CAMP
    if state["mode"] == "camp":
        if state["step"] == "organizer":
            data["organizer"] = user_msg
            state["step"] = "phone"
            return {"reply": "Organizer's contact number?"}
        if state["step"] == "phone":
            if not user_msg.isdigit() or len(user_msg) != 10:
                return {"reply": "Please enter a valid 10-digit phone number."}
            data["phone"] = user_msg
            state["step"] = "email"
            return {"reply": "Organizer's email?"}
        if state["step"] == "email":
            data["email"] = user_msg
            state["step"] = "location"
            return {"reply": "Where will the camp be held?"}
        if state["step"] == "location":
            data["location"] = user_msg
            state["step"] = "date"
            return {"reply": "Preferred camp date? (e.g., 22 July)"}
        if state["step"] == "date":
            data["date"] = user_msg
            state["step"] = "donors"
            return {"reply": "Expected number of donors?"}
        if state["step"] == "donors":
            data["donors"] = user_msg
            save_to_excel("camp", [
                data["organizer"], data["phone"], data["email"], data["location"], data["date"], data["donors"]
            ])
            state.update({"mode": None, "step": None, "data": {}})
            return {"reply": "Camp registered! We’ll follow up shortly."}

    return {"reply": "Sorry, I didn’t understand that. Please try again."}
