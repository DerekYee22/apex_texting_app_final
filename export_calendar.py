import win32com.client
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
from tkcalendar import Calendar
import os
from app import text

# add list of therapists (removed for privacy purposes)
THERAPISTS = []

# insert basepath (removed for privacy purposes)
base_path = 
export_path = os.path.join(base_path,'calendar.csv')

def get_outlook_calendar(name):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    for account in outlook.Folders:
        try:
            calendar_folder = account.Folders["Calendar"]
            for sub_calendar in calendar_folder.Folders:
                if sub_calendar.Name.lower() == name.lower():
                    return sub_calendar
        except Exception:
            continue
    return None

def list_outlook_calendars():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    calendars = []
    added_names = set()
    for account in outlook.Folders:
        try:
            calendar_folder = account.Folders["Calendar"]
            for sub_calendar in calendar_folder.Folders:
                if any(therapist.lower() in sub_calendar.Name.lower() for therapist in THERAPISTS):
                    if sub_calendar.Name not in added_names:
                        calendars.append((sub_calendar.Name, sub_calendar))
                        added_names.add(sub_calendar.Name)
        except Exception:
            continue
    return calendars

def ask_for_calendar_name():
    calendars = list_outlook_calendars()
    if not calendars:
        messagebox.showerror("No Calendars", "No Outlook calendars found.")
        return None, None

    root = tk.Tk()
    root.title("Select Calendar")
    root.geometry("300x120")
    root.eval('tk::PlaceWindow . center')

    selected = tk.StringVar(root)
    selected.set(calendars[0][0])  # Default selection

    tk.Label(root, text="Select an Outlook calendar:").pack(pady=10)
    dropdown = tk.OptionMenu(root, selected, *[name for name, _ in calendars])
    dropdown.pack(pady=5)

    def on_ok():
        root.destroy()

    ok_button = tk.Button(root, text="OK", command=on_ok)
    ok_button.pack(pady=10)

    root.mainloop()

    # Find the calendar object by name
    for name, cal in calendars:
        if name == selected.get():
            return cal, name
    return None, None

def ask_for_date():
    selected_date = {}

    def on_ok():
        selected_date["date"] = cal.selection_get()
        root.destroy()

    root = tk.Tk()
    root.title("Select Date")

    cal = Calendar(root, selectmode='day', year=datetime.now().year,
                   month=datetime.now().month, day=datetime.now().day)
    cal.pack(pady=20)

    ok_button = tk.Button(root, text="Export for Selected Date", command=on_ok)
    ok_button.pack(pady=10)

    root.mainloop()
    return selected_date.get("date", None)

def export_outlook_calendar_for_date(target_date, calendar_name):
    calendar = calendar_name

    # Define start and end of selected day
    start = datetime.combine(target_date, datetime.min.time())
    end = datetime.combine(target_date, datetime.max.time())

    restriction = f"[Start] >= '{start.strftime('%m/%d/%Y %H:%M %p')}' AND [End] <= '{end.strftime('%m/%d/%Y %H:%M %p')}'"
    items = calendar.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")
    restricted_items = items.Restrict(restriction)

    # Collect events
    calendar_data = []
    for item in restricted_items:
        calendar_data.append({
            "Subject": item.Subject,
            "Start": item.Start,
            "End": item.End,
            "Location": item.Location,
            "Body": item.Body if hasattr(item, "Body") else ""
        })

    # Export to CSV
    if calendar_data:
        df = pd.DataFrame(calendar_data)
        df.to_csv(export_path, index=False)
        messagebox.showinfo("Success", f"Calendar exported to:\n{export_path}")
        return True
    else:
        messagebox.showinfo("No Events", f"No events found for {target_date.strftime('%Y-%m-%d')} in {calendar_name}.")
        return False

if __name__ == "__main__":
    calendar, calendar_name = ask_for_calendar_name()
    date = ask_for_date()
    if date:
        if (export_outlook_calendar_for_date(date, calendar)):
            text(date, calendar_name)
