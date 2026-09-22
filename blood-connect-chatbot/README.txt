Blood Connect Chatbot

Blood Connect Chatbot is a web-based system developed for the Indian Red Cross Society’s Blood Bank. It enables seamless interaction for blood donation, blood requests, and camp registrations through a simple and efficient chat interface.

Overview

The system combines a lightweight HTML + JavaScript frontend with a FastAPI backend, storing all operational data in Excel files — eliminating the need for a traditional database. It is designed for quick deployment and ease of use, making blood bank management more accessible and automated.

Key Features

Interactive chatbot interface for donations, requests, and camp registration

Simple HTML + JavaScript front-end with real-time response handling

FastAPI-powered backend for chatbot logic and data processing

Excel-based data storage using openpyxl (no database setup required)

Local deployment with minimal configuration

Project Structure
blood-connect-chatbot/
├── frontend/
│   └── index.html
├── backend/
│   ├── app.py
│   ├── donations.xlsx
│   ├── requests.xlsx
│   └── camps.xlsx

How to Run

Install dependencies:

pip install fastapi uvicorn openpyxl


Start the backend server:

uvicorn app:app --reload


Open the frontend:

Simply open frontend/index.html in your browser.

No additional server is required for the frontend.

Tech Stack
Component	Technology
Frontend	HTML, CSS, JavaScript
Backend	FastAPI (Python)
Storage	Excel via OpenPyXL
Authors

Shreya Rai
BITS Pilani, Goa Campus

Aryaman Narula
BITS Pilani, Hyderabad Campus