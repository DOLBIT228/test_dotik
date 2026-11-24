import streamlit as st
import requests
import pandas as pd
from datetime import datetime
from io import BytesIO

# =====================================================
# PAGE CONFIG
# =====================================================

st.set_page_config(
    page_title="CRM Analytics",
    layout="wide"
)

# =====================================================
# SECRETS
# =====================================================

WEBHOOK = st.secrets["WEBHOOK"]

APP_LOGIN = st.secrets["APP_LOGIN"]
APP_PASSWORD = st.secrets["APP_PASSWORD"]

MANAGERS = dict(st.secrets["MANAGERS"])

# =====================================================
# AUTH
# =====================================================

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:

    st.title("Авторизація")

    login = st.text_input("Логін")
    password = st.text_input("Пароль", type="password")

    if st.button("Увійти"):

        if (
            login == APP_LOGIN
            and password == APP_PASSWORD
        ):

            st.session_state.authenticated = True
            st.rerun()

        else:

            st.error("Невірний логін або пароль")

    st.stop()

# =====================================================
# TITLE
# =====================================================

st.title("Аналітика спроб дзвінків")

# =====================================================
# STATUSES
# =====================================================

STATUSES = {
    "C65:NEW": "Новий запит",
    "C65:PREPAYMENT_INVOIC": "Взяті у роботу",
    "C65:EXECUTING": "Додзвон"
}

# =====================================================
# FILTERS
# =====================================================

col1, col2 = st.columns(2)

with col1:

    date_from = st.date_input(
        "Дата від",
        value=datetime.now().replace(day=1)
    )

with col2:

    date_to = st.date_input(
        "Дата до",
        value=datetime.now()
    )

# =====================================================
# API
# =====================================================

@st.cache_data(ttl=300)
def bx_api(method, params=None):

    if params is None:
        params = {}

    url = f"{WEBHOOK}{method}.json"

    response = requests.get(
        url,
        params=params
    )

    data = response.json()

    if "error" in data:

        st.error(data)

        return []

    return data.get("result", [])

# =====================================================
# DATE
# =====================================================

def parse_date(date_string):

    if not date_string:
        return None

    try:

        return datetime.fromisoformat(
            date_string.replace("Z", "+00:00")
        ).replace(tzinfo=None)

    except:

        return None

# =====================================================
# CALL CHECK
# =====================================================

def is_call_attempt(activity):

    subject = str(
        activity.get("SUBJECT", "")
    ).lower()

    type_id = str(
        activity.get("TYPE_ID", "")
    )

    if type_id != "2":
        return False

    if (
        "outgoing" in subject
        or "incoming" in subject
    ):
        return True

    return False

# =====================================================
# GROUPS
# =====================================================

def get_group(count):

    if count == 0:
        return "0"

    elif 1 <= count <= 3:
        return "1-3"

    elif 4 <= count <= 6:
        return "4-6"

    else:
        return "7+"

# =====================================================
# EXCEL
# =====================================================

def create_excel(
    summary_df,
    deals_df,
    managers_df
):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        summary_df.to_excel(
            writer,
            sheet_name="Загальна статистика",
            index=False
        )

        deals_df.to_excel(
            writer,
            sheet_name="Угоди",
            index=False
        )

        managers_df.to_excel(
            writer,
            sheet_name="Менеджери",
            index=False
        )

    return output.getvalue()

# =====================================================
# ANALYZE
# =====================================================

if st.button("Запустити аналіз"):

    summary_rows = []
    deals_rows = []
    managers_rows = []

    manager_stats = {}

    progress = st.progress(0)

    total_statuses = len(STATUSES)

    for index, (status_id, status_name) in enumerate(
        STATUSES.items()
    ):

        deals = bx_api(
            "crm.deal.list",
            {
                "filter[STAGE_ID]": status_id,

                "filter[>=DATE_CREATE]": str(date_from),
                "filter[<=DATE_CREATE]": str(date_to),

                "select[]": [
                    "ID",
                    "TITLE",
                    "DATE_CREATE",
                    "ASSIGNED_BY_ID"
                ]
            }
        )

        group_stats = {
            "0": 0,
            "1-3": 0,
            "4-6": 0,
            "7+": 0
        }

        total_calls_status = 0

        for deal in deals:

            deal_id = deal.get("ID")

            title = deal.get("TITLE")

            created_date = deal.get("DATE_CREATE")

            manager_id = str(
                deal.get("ASSIGNED_BY_ID")
            )

            manager_name = MANAGERS.get(
                manager_id,
                f"ID {manager_id}"
            )

            activities = bx_api(
                "crm.activity.list",
                {
                    "filter[OWNER_TYPE_ID]": 2,
                    "filter[OWNER_ID]": deal_id
                }
            )

            incoming_calls = 0
            outgoing_calls = 0
            calls_count = 0

            last_call = "-"

            call_dates = []

            for activity in activities:

                if is_call_attempt(activity):

                    calls_count += 1

                    subject = str(
                        activity.get("SUBJECT", "")
                    ).lower()

                    if "incoming" in subject:
                        incoming_calls += 1

                    if "outgoing" in subject:
                        outgoing_calls += 1

                    created = parse_date(
                        activity.get("CREATED")
                    )

                    if created:
                        call_dates.append(created)

            if call_dates:

                last_call = max(
                    call_dates
                ).strftime(
                    "%Y-%m-%d %H:%M"
                )

            group = get_group(calls_count)

            group_stats[group] += 1

            total_calls_status += calls_count

            # =========================================
            # DEALS
            # =========================================

            deals_rows.append({

                "Статус": status_name,
                "ID угоди": deal_id,
                "Назва": title,
                "Дата створення": created_date,
                "Менеджер": manager_name,

                "Спроб дзвінка": calls_count,

                "Incoming": incoming_calls,
                "Outgoing": outgoing_calls,

                "Група": group,

                "Останній дзвінок": last_call
            })

            # =========================================
            # MANAGERS
            # =========================================

            if manager_name not in manager_stats:

                manager_stats[manager_name] = {
                    "deals": 0,
                    "calls": 0
                }

            manager_stats[manager_name]["deals"] += 1
            manager_stats[manager_name]["calls"] += calls_count

        avg_calls = 0

        if len(deals) > 0:

            avg_calls = round(
                total_calls_status / len(deals),
                2
            )

        summary_rows.append({

            "Статус": status_name,

            "Угод": len(deals),

            "Середнє дзвінків": avg_calls,

            "0 дзвінків": group_stats["0"],
            "1-3 дзвінки": group_stats["1-3"],
            "4-6 дзвінків": group_stats["4-6"],
            "7+ дзвінків": group_stats["7+"]
        })

        progress.progress(
            (index + 1) / total_statuses
        )

    # =================================================
    # MANAGERS SUMMARY
    # =================================================

    for manager_name, stats in manager_stats.items():

        avg_calls = 0

        if stats["deals"] > 0:

            avg_calls = round(
                stats["calls"] / stats["deals"],
                2
            )

        managers_rows.append({

            "Менеджер": manager_name,

            "Угод": stats["deals"],

            "Всього дзвінків": stats["calls"],

            "Середнє дзвінків": avg_calls
        })

    # =================================================
    # DATAFRAMES
    # =================================================

    summary_df = pd.DataFrame(summary_rows)

    deals_df = pd.DataFrame(deals_rows)

    managers_df = pd.DataFrame(managers_rows)

    # =================================================
    # TABLES
    # =================================================

    st.success("Аналіз завершено")

    st.subheader("Загальна статистика")

    st.dataframe(
        summary_df,
        use_container_width=True
    )

    st.subheader("Детально по угодах")

    st.dataframe(
        deals_df,
        use_container_width=True
    )

    st.subheader("Менеджери")

    st.dataframe(
        managers_df,
        use_container_width=True
    )

    # =================================================
    # EXCEL
    # =================================================

    excel_file = create_excel(
        summary_df,
        deals_df,
        managers_df
    )

    st.download_button(
        label="Завантажити Excel",
        data=excel_file,
        file_name="crm_analytics.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )