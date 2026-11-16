"""
Mock Data Generator for Microsoft 365 Copilot Usage & Daily Trends.
"""

from datetime import datetime, timedelta
import random
from typing import List, Dict, Any


def generate_mock_copilot_users(count: int = 45) -> List[Dict[str, Any]]:
    first_names = [
        "Alex", "Jordan", "Taylor", "Morgan", "Sam", "Chris", "Pat", "Casey", "Jamie",
        "Emma", "Liam", "Olivia", "Noah", "Sophia", "Lucas", "Ava", "Mason", "Isabella",
        "Ethan", "Mia", "James", "Harper", "Benjamin", "Evelyn", "Daniel", "Chloe", "Michael"
    ]
    last_names = [
        "Smith", "Johnson", "Williams", "Brown", "Jones", "Miller", "Davis", "Wilson", "Anderson",
        "Taylor", "Thomas", "Moore", "Jackson", "Martin", "Lee", "Perez", "Thompson", "White", "Harris"
    ]

    now = datetime.now()
    records = []
    random.seed(42)

    apps = ["Chat", "Teams", "Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Loop"]

    for i in range(count):
        fname = random.choice(first_names)
        lname = random.choice(last_names)
        display_name = f"{fname} {lname}"
        email = f"{fname.lower()}.{lname.lower()}{i if i > 25 else ''}@contoso.com"

        user_type = random.choices(["power", "regular", "casual", "inactive"], weights=[15, 35, 30, 20])[0]

        app_last_activity = {app: "" for app in apps}
        apps_used = []

        if user_type == "inactive":
            last_activity_date = ""
            is_active = False
            total_active_days = 0
            prompts = 0
            prompts_web = 0
            prompts_work = 0
        else:
            is_active = True
            days_ago = random.randint(0, 25)
            last_activity = now - timedelta(days=days_ago)
            last_activity_date = last_activity.strftime("%Y-%m-%d")

            if user_type == "power":
                total_active_days = random.randint(15, 24)
                prompts = random.randint(150, 480)
                chosen_apps = random.sample(apps, k=random.randint(5, 7))
            elif user_type == "regular":
                total_active_days = random.randint(5, 14)
                prompts = random.randint(30, 140)
                chosen_apps = random.sample(apps, k=random.randint(3, 5))
            else:  # casual
                total_active_days = random.randint(1, 4)
                prompts = random.randint(2, 25)
                chosen_apps = random.sample(apps, k=random.randint(1, 2))

            for app in chosen_apps:
                act_date = now - timedelta(days=random.randint(0, 25))
                app_last_activity[app] = act_date.strftime("%Y-%m-%d")
                apps_used.append(app)

            prompts_web = int(prompts * 0.4)
            prompts_work = prompts - prompts_web

        records.append({
            "user_principal_name": email,
            "display_name": display_name,
            "report_refresh_date": now.strftime("%Y-%m-%d"),
            "last_activity_date": last_activity_date,
            "is_active": is_active,
            "apps_used": apps_used,
            "total_active_days": total_active_days,
            "prompts_submitted": prompts,
            "prompts_chat_web": prompts_web,
            "prompts_chat_work": prompts_work,
            "app_last_activity": app_last_activity,
        })

    records.sort(key=lambda x: (x["is_active"], x.get("prompts_submitted", 0), x["total_active_days"]), reverse=True)
    return records


def generate_mock_copilot_trend(days: int = 28) -> List[Dict[str, Any]]:
    """Generates synthetic daily active users and prompt volume trend."""
    now = datetime.now()
    trend = []
    random.seed(101)

    for i in range(days, 0, -1):
        dt = now - timedelta(days=i)
        is_weekend = dt.weekday() >= 5
        if is_weekend:
            active_users = random.randint(1, 3)
            prompts = random.randint(8, 25)
        else:
            active_users = random.randint(6, 12)
            prompts = random.randint(65, 190)

        trend.append({
            "reportDate": dt.strftime("%Y-%m-%d"),
            "anyAppActiveUsers": active_users,
            "anyAppEnabledUsers": 28,
            "promptsSubmitted": prompts,
        })
    return trend
