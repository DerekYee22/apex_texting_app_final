Apex PT Automated Appointment Reminder System

Automated SMS appointment reminder tool built for Apex Physical Therapy Specialists to reduce staff workload and streamline patient communication.

Project Overview

Apex Physical Therapy staff were spending several hours each week calling patients individually to confirm upcoming appointments.
This project automates the entire workflow by:

Extracting patient appointments directly from Outlook calendars

Parsing and formatting appointment + contact details

Sending automated SMS reminders through the Twilio API

Providing a simple desktop GUI that staff can run with zero technical experience

Logging all outgoing text messages for documentation and accountability

This tool reduced manual calling from hours per week to a few clicks.

Features
Automated Appointment Extraction

Reads exported Outlook events and pulls out:

Patient name

Appointment date & time

Phone number (parsed from the Outlook event body)

(Implemented in app.py.)

✔ Outlook → CSV Export

Uses a GUI to let staff select:

A therapist’s calendar

A date

Export all events into calendar.csv

(Implemented in export_calendar.py.)

✔ Staff-Friendly Desktop Interface

Simple workflow:

Choose therapist

Choose date

Export calendar

Preview reminders

Send

Automatic Logging

All reminders are written into log.csv, stored chronologically.

📬 Contact

Feel free to reach out with questions or ideas for enhancement!
