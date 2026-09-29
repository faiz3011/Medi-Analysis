import streamlit as st
import pandas as pd
import os
import io
import plotly.express as px
from datetime import datetime, date


# ============================================================
# MEDI ANALYSIS - HOSPITAL INTELLIGENCE SYSTEM
# ============================================================

st.set_page_config(
    page_title="Medi Analysis",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# CONFIGURATION
# ============================================================

DATA_FOLDER = "data"

FILE_PATH = os.path.join(
    DATA_FOLDER,
    "patients.xlsx"
)

ACCOUNT_FILE = os.path.join(
    DATA_FOLDER,
    "hospital_account.xlsx"
)


DEPARTMENTS = [
    "General Medicine",
    "Cardiology",
    "Neurology",
    "Orthopedics",
    "Pediatrics",
    "Gynecology",
    "Emergency",
    "Dermatology",
    "ENT",
    "Other"
]


GENDERS = [
    "Male",
    "Female",
    "Other"
]


REQUIRED_COLUMNS = {
    "Patient ID": "",
    "Name": "",
    "Age": 0,
    "Gender": "Other",
    "Phone": "",
    "Address": "",
    "Department": "General Medicine",
    "Diagnosis": "",
    "Admission Date": "",
    "Admission Time": "",
    "Enrollment Date": "",
    "Enrollment Time": "",
    "Bill Amount": 0.0
}


# ============================================================
# CREATE DATA FOLDER
# ============================================================

os.makedirs(DATA_FOLDER, exist_ok=True)


# ============================================================
# CUSTOM CSS
# ============================================================


st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', Arial, sans-serif;
    }

    .stApp {
        background:
            radial-gradient(circle at 10% 10%, rgba(99, 60, 180, 0.18), transparent 28%),
            radial-gradient(circle at 90% 15%, rgba(45, 100, 210, 0.14), transparent 25%),
            radial-gradient(circle at 50% 90%, rgba(115, 55, 180, 0.10), transparent 30%),
            linear-gradient(135deg, #050507 0%, #0b0912 38%, #0c1020 70%, #08070d 100%);
        color: #f7f7fb;
        min-height: 100vh;
    }

    .block-container {
        max-width: 1480px;
        padding-top: 2rem;
        padding-bottom: 3.5rem;
    }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #050508 0%, #090811 52%, #080d18 100%);
        border-right: 1px solid rgba(145, 120, 255, 0.18);
    }

    section[data-testid="stSidebar"] > div {
        padding-top: 1.4rem;
    }

    .main-logo {
        font-size: 48px;
        font-weight: 800;
        letter-spacing: -1.5px;
        background: linear-gradient(100deg, #ffffff 0%, #d9caff 45%, #8eb8ff 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 2px;
    }

    .main-subtitle {
        color: #a8a4b5;
        font-size: 14px;
        font-weight: 500;
        letter-spacing: .2px;
        margin-bottom: 28px;
    }

    .section-title {
        font-size: 32px;
        font-weight: 800;
        letter-spacing: -.7px;
        background: linear-gradient(100deg, #ffffff, #cdbbff 55%, #86adff);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 4px;
    }

    .hero-card {
        background: linear-gradient(135deg, rgba(20,18,30,.92), rgba(13,18,34,.88));
        border: 1px solid rgba(166,145,255,.20);
        border-radius: 22px;
        padding: 24px 26px;
        margin: 10px 0 22px 0;
        box-shadow: 0 16px 45px rgba(0,0,0,.25);
    }

    .hero-kicker {
        color: #9e91c7;
        text-transform: uppercase;
        letter-spacing: 2px;
        font-size: 11px;
        font-weight: 700;
        margin-bottom: 7px;
    }

    .hero-title {
        font-size: 25px;
        font-weight: 750;
        color: #f7f5ff;
        margin-bottom: 5px;
    }

    .hero-text {
        color: #aaa7b7;
        font-size: 14px;
        line-height: 1.6;
    }

    [data-testid="stMetric"] {
        background: linear-gradient(145deg, rgba(255,255,255,.055), rgba(110,72,190,.075));
        border: 1px solid rgba(170,145,255,.16);
        padding: 19px 18px;
        border-radius: 18px;
        box-shadow: 0 10px 28px rgba(0,0,0,.18);
        min-height: 108px;
    }

    [data-testid="stMetricLabel"] {
        color: #aaa7b8;
        font-weight: 600;
    }

    [data-testid="stMetricValue"] {
        color: #f8f7ff;
        font-weight: 800;
    }

    .stButton > button, .stDownloadButton > button {
        border-radius: 12px;
        min-height: 44px;
        font-weight: 700;
        border: 1px solid rgba(168,145,255,.22);
        background: linear-gradient(135deg, rgba(93,66,160,.45), rgba(43,77,145,.35));
        color: #ffffff;
        box-shadow: 0 8px 22px rgba(0,0,0,.18);
        transition: all .18s ease;
    }

    .stButton > button:hover, .stDownloadButton > button:hover {
        border-color: rgba(183,162,255,.55);
        transform: translateY(-1px);
    }

    div[data-baseweb="input"] > div,
    div[data-baseweb="select"] > div,
    div[data-baseweb="textarea"] > div {
        background: rgba(10, 9, 17, .72);
        border: 1px solid rgba(160,145,200,.18);
        border-radius: 11px;
    }

    div[data-baseweb="input"] input,
    div[data-baseweb="textarea"] textarea {
        color: #f5f4fa;
    }

    div[data-baseweb="select"] * {
        color: #f5f4fa;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: rgba(10,9,16,.45);
        padding: 7px;
        border-radius: 13px;
        border: 1px solid rgba(160,145,200,.12);
    }

    .stTabs [data-baseweb="tab"] {
        border-radius: 9px;
        font-weight: 650;
    }

    .stAlert {
        border-radius: 14px;
        border: 1px solid rgba(160,145,200,.16);
    }

    div[data-testid="stDataFrame"] {
        border-radius: 14px;
        overflow: hidden;
        border: 1px solid rgba(160,145,200,.14);
    }

    hr {
        border-color: rgba(160,145,200,.12) !important;
    }

    .sidebar-brand {
        padding: 10px 4px 18px 4px;
    }

    .sidebar-brand-title {
        font-size: 22px;
        font-weight: 800;
        color: #f8f7ff;
    }

    .sidebar-brand-subtitle {
        color: #8e8a9d;
        font-size: 11px;
        margin-top: 3px;
        letter-spacing: .5px;
    }

    .status-card {
        background: rgba(44, 70, 95, .14);
        border: 1px solid rgba(86, 165, 125, .20);
        border-radius: 12px;
        padding: 10px 12px;
        color: #b7e4ca;
        font-size: 12px;
        font-weight: 650;
    }

    .footer-wrap {
        text-align: center;
        padding: 14px 0 2px 0;
    }

    .footer-main {
        color: #9994a8;
        font-size: 12px;
    }

    .footer-credit {
        color: #777284;
        font-size: 11px;
        margin-top: 5px;
    }
    </style>
    """,
    unsafe_allow_html=True
)



# ============================================================
# SESSION STATE
# ============================================================

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "show_welcome" not in st.session_state:
    st.session_state.show_welcome = False

if "hospital_name" not in st.session_state:
    st.session_state.hospital_name = ""


# ============================================================
# HOSPITAL ACCOUNT FUNCTIONS
# ============================================================

def load_hospital_account():

    if not os.path.exists(ACCOUNT_FILE):
        return pd.DataFrame()

    try:
        return pd.read_excel(
            ACCOUNT_FILE,
            engine="openpyxl"
        )

    except Exception:
        return pd.DataFrame()


def save_hospital_account(account_data):

    try:

        account_data.to_excel(
            ACCOUNT_FILE,
            index=False,
            engine="openpyxl"
        )

        return True

    except Exception as e:

        st.error(
            f"Unable to save hospital account: {e}"
        )

        return False


# ============================================================
# PATIENT DATA FUNCTIONS
# ============================================================

def load_data():

    if os.path.exists(FILE_PATH):

        try:

            data = pd.read_excel(
                FILE_PATH,
                engine="openpyxl"
            )

        except Exception:

            data = pd.DataFrame()

    else:

        data = pd.DataFrame()


    for column, default_value in REQUIRED_COLUMNS.items():

        if column not in data.columns:

            data[column] = default_value


    return data


def save_data(data):

    try:

        data.to_excel(
            FILE_PATH,
            index=False,
            engine="openpyxl"
        )

        return True

    except Exception as e:

        st.error(
            f"Unable to save patient database: {e}"
        )

        return False


def generate_patient_id(data):

    if data.empty:

        return "P1001"


    numbers = []


    for patient_id in data["Patient ID"].astype(str):

        digits = "".join(
            character
            for character in patient_id
            if character.isdigit()
        )

        if digits:

            numbers.append(
                int(digits)
            )


    if not numbers:

        return "P1001"


    return f"P{max(numbers) + 1}"


def to_excel_bytes(data):

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        data.to_excel(
            writer,
            index=False,
            sheet_name="Patient Records"
        )

    return output.getvalue()


# ============================================================
# LOGIN / REGISTRATION
# ============================================================

if not st.session_state.logged_in:

    st.markdown("##")

    col1, col2, col3 = st.columns(
        [1, 1.5, 1]
    )


    with col2:

        st.title("🏥 Medi Analysis")

        st.caption(
            "Hospital Intelligence System • Secure Access"
        )


        login_tab, register_tab = st.tabs(
            [
                "🔐 Sign In",
                "📝 Register Hospital"
            ]
        )


        # ====================================================
        # LOGIN
        # ====================================================

        with login_tab:

            st.subheader("Welcome Back")

            with st.form("login_form"):

                username = st.text_input(
                    "Username"
                )

                password = st.text_input(
                    "Password",
                    type="password"
                )

                login_button = st.form_submit_button(
                    "🔐 Sign In",
                    use_container_width=True
                )


            if login_button:

                account_data = load_hospital_account()


                if account_data.empty:

                    st.error(
                        "No account found. Please register first."
                    )

                else:

                    matched = account_data[
                        (
                            account_data["Username"]
                            .astype(str)
                            .str.lower()
                            == username.strip().lower()
                        )
                        &
                        (
                            account_data["Password"]
                            .astype(str)
                            == password
                        )
                    ]


                    if not matched.empty:

                        st.session_state.logged_in = True

                        st.session_state.show_welcome = True

                        st.session_state.hospital_name = str(
                            matched.iloc[0]["Hospital Name"]
                        )

                        st.rerun()

                    else:

                        st.error(
                            "Invalid username or password."
                        )


        # ====================================================
        # REGISTER HOSPITAL
        # ====================================================

        with register_tab:

            st.subheader("Register Hospital")

            with st.form("registration_form"):

                hospital_name = st.text_input(
                    "Hospital Name *"
                )

                admin_name = st.text_input(
                    "Administrator Name *"
                )

                email = st.text_input(
                    "Email Address *"
                )

                phone = st.text_input(
                    "Contact Number *"
                )

                address = st.text_area(
                    "Hospital Address *"
                )

                st.divider()

                new_username = st.text_input(
                    "Create Username *"
                )

                new_password = st.text_input(
                    "Create Password *",
                    type="password"
                )

                confirm_password = st.text_input(
                    "Confirm Password *",
                    type="password"
                )


                register_button = st.form_submit_button(
                    "📝 Create Hospital Account",
                    use_container_width=True
                )


            if register_button:

                if not hospital_name.strip():

                    st.error(
                        "Hospital name is required."
                    )

                elif not admin_name.strip():

                    st.error(
                        "Administrator name is required."
                    )

                elif not email.strip():

                    st.error(
                        "Email is required."
                    )

                elif not phone.strip():

                    st.error(
                        "Phone number is required."
                    )

                elif not address.strip():

                    st.error(
                        "Hospital address is required."
                    )

                elif not new_username.strip():

                    st.error(
                        "Username is required."
                    )

                elif len(new_password) < 6:

                    st.error(
                        "Password must be at least 6 characters."
                    )

                elif new_password != confirm_password:

                    st.error(
                        "Passwords do not match."
                    )

                else:

                    account_data = load_hospital_account()


                    new_account = pd.DataFrame({

                        "Hospital Name": [
                            hospital_name.strip()
                        ],

                        "Administrator Name": [
                            admin_name.strip()
                        ],

                        "Email": [
                            email.strip()
                        ],

                        "Phone": [
                            phone.strip()
                        ],

                        "Address": [
                            address.strip()
                        ],

                        "Username": [
                            new_username.strip()
                        ],

                        "Password": [
                            new_password
                        ]

                    })


                    updated_accounts = pd.concat(
                        [
                            account_data,
                            new_account
                        ],
                        ignore_index=True
                    )


                    if save_hospital_account(
                        updated_accounts
                    ):

                        st.success(
                            "Hospital account created successfully!"
                        )


    st.stop()


# ============================================================
# WELCOME SCREEN
# ============================================================

if st.session_state.show_welcome:

    st.markdown(
        """
        <div style="
            text-align:center;
            padding:75px 20px 45px 20px;
        ">

        <h1 style="
            font-size:60px;
        ">
        🏥 MEDI ANALYSIS
        </h1>

        <h3 style="
            color:#aaa5b8;
        ">
        Hospital Intelligence System
        </h3>

        </div>
        """,
        unsafe_allow_html=True
    )


    st.info(
        """
### Welcome to Medi Analysis

Smart Patient Management • Healthcare Analytics • Hospital Intelligence

Manage patient records, analyse hospital data and generate
meaningful healthcare insights from one powerful system.
"""
    )


    if st.button(
        "🚀 ENTER MEDI ANALYSIS",
        use_container_width=True
    ):

        st.session_state.show_welcome = False

        st.rerun()


    st.stop()


# ============================================================
# LOAD DATABASE
# ============================================================

df = load_data()


# ============================================================
# HEADER
# ============================================================

st.markdown(
    '<div class="main-logo">✦ Medi Analysis</div>',
    unsafe_allow_html=True
)


st.markdown(
    '<div class="main-subtitle">'
    'Hospital Intelligence • Patient Management • Healthcare Analytics'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    f'''
    <div class="hero-card">
        <div class="hero-kicker">Hospital Intelligence System</div>
        <div class="hero-title">Welcome to {st.session_state.hospital_name or "Medi Analysis"}</div>
        <div class="hero-text">Manage patient records, monitor hospital activity and turn healthcare data into clear, actionable insights.</div>
    </div>
    ''',
    unsafe_allow_html=True
)


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.markdown(
    f"""
    <div class="sidebar-brand">
        <div class="sidebar-brand-title">🏥 Medi Analysis</div>
        <div class="sidebar-brand-subtitle">HOSPITAL INTELLIGENCE SYSTEM</div>
    </div>
    """,
    unsafe_allow_html=True
)

st.sidebar.caption(
    f"Hospital: {st.session_state.hospital_name}"
)


st.sidebar.divider()


page = st.sidebar.radio(
    "Navigation",
    [
        "📊 Dashboard",
        "👤 Add Patient",
        "📋 Patient Records",
        "📈 Data Analysis"
    ]
)


st.sidebar.divider()


st.sidebar.markdown(
    '<div class="status-card">🟢 Database Connected</div>',
    unsafe_allow_html=True
)


st.sidebar.caption(
    f"👥 {len(df)} Patient Record(s)"
)


st.sidebar.divider()


if st.sidebar.button(
    "🚪 Logout",
    use_container_width=True
):

    st.session_state.logged_in = False

    st.session_state.show_welcome = False

    st.session_state.hospital_name = ""

    st.rerun()


# ============================================================
# PAGE ROUTING
# ============================================================


# ============================================================
# DASHBOARD
# ============================================================

if page == "📊 Dashboard":

    st.markdown(
        '<div class="section-title">📊 Dashboard</div>',
        unsafe_allow_html=True
    )


    hospital_display_name = st.session_state.hospital_name


    st.subheader(
        f"🏥 {hospital_display_name}"
    )


    st.caption(
        "Real-time patient intelligence and hospital performance."
    )


    if df.empty:

        st.warning(
            "No patient records available. "
            "Please add your first patient."
        )

    else:

        dashboard_df = df.copy()


        dashboard_df["Bill Amount"] = pd.to_numeric(
            dashboard_df["Bill Amount"],
            errors="coerce"
        ).fillna(0)


        total_patients = len(
            dashboard_df
        )


        male = len(
            dashboard_df[
                dashboard_df["Gender"] == "Male"
            ]
        )


        female = len(
            dashboard_df[
                dashboard_df["Gender"] == "Female"
            ]
        )


        total_revenue = dashboard_df[
            "Bill Amount"
        ].sum()


        departments_count = dashboard_df[
            "Department"
        ].nunique()


        col1, col2, col3, col4, col5 = st.columns(5)


        with col1:

            st.metric(
                "👥 Total Patients",
                total_patients
            )


        with col2:

            st.metric(
                "👨 Male",
                male
            )


        with col3:

            st.metric(
                "👩 Female",
                female
            )


        with col4:

            st.metric(
                "💰 Revenue",
                f"₹ {total_revenue:,.0f}"
            )


        with col5:

            st.metric(
                "🏥 Departments",
                departments_count
            )


        st.divider()


        department_data = (
            dashboard_df["Department"]
            .value_counts()
            .reset_index()
        )

        department_data.columns = [
            "Department",
            "Patients"
        ]


        fig = px.bar(
            department_data,
            x="Department",
            y="Patients",
            text="Patients",
            title="Department-wise Patient Distribution"
        )


        fig.update_layout(
            template="plotly_dark",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Inter, Arial", color="#E9E7F2"),
            margin=dict(l=20, r=20, t=70, b=35),
            title_font=dict(size=18)
        )


        st.plotly_chart(
            fig,
            use_container_width=True
        )


# ============================================================
# ADD PATIENT
# ============================================================

elif page == "👤 Add Patient":

    st.markdown(
        '<div class="section-title">👤 Register New Patient</div>',
        unsafe_allow_html=True
    )


    st.caption(
        "Create a complete and secure patient record."
    )


    next_patient_id = generate_patient_id(df)


    with st.form("patient_enrollment_form"):

        st.subheader(
            "🧑 Patient Information"
        )


        st.info(
            f"🆔 Automatic Patient ID: {next_patient_id}"
        )


        col1, col2 = st.columns(2)


        with col1:

            name = st.text_input(
                "Patient Name *"
            )


            age = st.number_input(
                "Age",
                min_value=0,
                max_value=120,
                value=25
            )


            gender = st.selectbox(
                "Gender",
                GENDERS
            )


            phone = st.text_input(
                "Phone Number"
            )


        with col2:

            address = st.text_area(
                "Patient Address"
            )


            department = st.selectbox(
                "Department",
                DEPARTMENTS
            )


            diagnosis = st.text_input(
                "Diagnosis"
            )


        st.divider()


        st.subheader(
            "📅 Admission & Billing"
        )


        col1, col2, col3 = st.columns(3)


        with col1:

            admission_date = st.date_input(
                "Admission Date",
                value=date.today()
            )


        with col2:

            admission_time = st.time_input(
                "Admission Time",
                value=datetime.now().time().replace(
                    second=0,
                    microsecond=0
                )
            )


        with col3:

            bill_amount = st.number_input(
                "Bill Amount (₹)",
                min_value=0.0,
                value=0.0,
                step=100.0
            )


        submitted = st.form_submit_button(
            "➕ Register Patient",
            use_container_width=True
        )


    if submitted:

        name_clean = name.strip()


        if not name_clean:

            st.error(
                "Patient name is required."
            )

        else:

            enrollment_timestamp = datetime.now()


            new_patient = pd.DataFrame({

                "Patient ID": [
                    next_patient_id
                ],

                "Name": [
                    name_clean
                ],

                "Age": [
                    age
                ],

                "Gender": [
                    gender
                ],

                "Phone": [
                    phone.strip()
                ],

                "Address": [
                    address.strip()
                ],

                "Department": [
                    department
                ],

                "Diagnosis": [
                    diagnosis.strip()
                ],

                "Admission Date": [
                    admission_date.strftime("%d-%m-%Y")
                ],

                "Admission Time": [
                    admission_time.strftime("%I:%M %p")
                ],

                "Enrollment Date": [
                    enrollment_timestamp.strftime("%d-%m-%Y")
                ],

                "Enrollment Time": [
                    enrollment_timestamp.strftime("%I:%M:%S %p")
                ],

                "Bill Amount": [
                    bill_amount
                ]

            })


            updated_data = pd.concat(
                [
                    df,
                    new_patient
                ],
                ignore_index=True
            )


            if save_data(updated_data):

                st.success(
                    f"✅ Patient {name_clean} registered successfully!"
                )

                st.balloons()


# ============================================================
# PATIENT RECORDS
# ============================================================

elif page == "📋 Patient Records":

    st.markdown(
        '<div class="section-title">📋 Patient Records</div>',
        unsafe_allow_html=True
    )


    if df.empty:

        st.warning(
            "No patient records available."
        )

    else:

        st.subheader(
            "🔎 Search Patient"
        )


        search = st.text_input(
            "Search by ID, Name or Phone"
        )


        filtered_df = df.copy()


        if search.strip():

            search_text = search.lower().strip()


            filtered_df = filtered_df[
                filtered_df["Patient ID"]
                .astype(str)
                .str.lower()
                .str.contains(
                    search_text,
                    na=False
                )
                |
                filtered_df["Name"]
                .astype(str)
                .str.lower()
                .str.contains(
                    search_text,
                    na=False
                )
                |
                filtered_df["Phone"]
                .astype(str)
                .str.lower()
                .str.contains(
                    search_text,
                    na=False
                )
            ]


        st.dataframe(
            filtered_df,
            use_container_width=True,
            hide_index=True
        )


        st.divider()


        if not filtered_df.empty:

            selected_id = st.selectbox(
                "Select Patient",
                filtered_df["Patient ID"]
                .astype(str)
                .tolist()
            )


            selected_patient = filtered_df[
                filtered_df["Patient ID"]
                .astype(str)
                == selected_id
            ].iloc[0]


            st.subheader(
                f"👤 {selected_patient['Name']}"
            )


            col1, col2, col3, col4 = st.columns(4)


            with col1:

                st.metric(
                    "Age",
                    selected_patient["Age"]
                )


            with col2:

                st.metric(
                    "Gender",
                    selected_patient["Gender"]
                )


            with col3:

                st.metric(
                    "Department",
                    selected_patient["Department"]
                )


            with col4:

                bill = pd.to_numeric(
                    selected_patient["Bill Amount"],
                    errors="coerce"
                )

                if pd.isna(bill):
                    bill = 0

                st.metric(
                    "Bill",
                    f"₹ {bill:,.0f}"
                )


            st.divider()


            # =================================================
            # EDIT PATIENT
            # =================================================

            st.subheader(
                "✏️ Edit Patient"
            )


            with st.form(
                f"edit_patient_{selected_id}"
            ):

                edit_name = st.text_input(
                    "Patient Name",
                    value=str(selected_patient["Name"])
                )


                edit_age = st.number_input(
                    "Age",
                    min_value=0,
                    max_value=120,
                    value=int(selected_patient["Age"])
                )


                edit_gender = st.selectbox(
                    "Gender",
                    GENDERS,
                    index=GENDERS.index(
                        selected_patient["Gender"]
                    )
                    if selected_patient["Gender"] in GENDERS
                    else 0
                )


                edit_department = st.selectbox(
                    "Department",
                    DEPARTMENTS,
                    index=DEPARTMENTS.index(
                        selected_patient["Department"]
                    )
                    if selected_patient["Department"] in DEPARTMENTS
                    else 0
                )


                edit_diagnosis = st.text_input(
                    "Diagnosis",
                    value=str(selected_patient["Diagnosis"])
                )


                edit_bill = st.number_input(
                    "Bill Amount",
                    min_value=0.0,
                    value=float(
                        pd.to_numeric(
                            selected_patient["Bill Amount"],
                            errors="coerce"
                        )
                        if pd.notna(
                            pd.to_numeric(
                                selected_patient["Bill Amount"],
                                errors="coerce"
                            )
                        )
                        else 0
                    )
                )


                update_button = st.form_submit_button(
                    "💾 Save Changes",
                    use_container_width=True
                )


            if update_button:

                condition = (
                    df["Patient ID"]
                    .astype(str)
                    == selected_id
                )


                df.loc[
                    condition,
                    "Name"
                ] = edit_name


                df.loc[
                    condition,
                    "Age"
                ] = edit_age


                df.loc[
                    condition,
                    "Gender"
                ] = edit_gender


                df.loc[
                    condition,
                    "Department"
                ] = edit_department


                df.loc[
                    condition,
                    "Diagnosis"
                ] = edit_diagnosis


                df.loc[
                    condition,
                    "Bill Amount"
                ] = edit_bill


                if save_data(df):

                    st.success(
                        "Patient updated successfully!"
                    )

                    st.rerun()


            st.divider()


            # =================================================
            # DELETE PATIENT
            # =================================================

            confirm_delete = st.checkbox(
                "I understand this action is permanent.",
                key=f"confirm_delete_{selected_id}"
            )


            if confirm_delete:

                if st.button(
                    "🗑️ Delete Patient",
                    use_container_width=True,
                    key=f"delete_{selected_id}"
                ):

                    updated_df = df[
                        df["Patient ID"]
                        .astype(str)
                        != selected_id
                    ]


                    if save_data(updated_df):

                        st.success(
                            "Patient deleted successfully!"
                        )

                        st.rerun()


        st.divider()


        st.subheader(
            "📥 Export Records"
        )


        col1, col2 = st.columns(2)


        with col1:

            csv_data = filtered_df.to_csv(
                index=False
            ).encode("utf-8")


            st.download_button(
                "📄 Download CSV",
                data=csv_data,
                file_name="patient_records.csv",
                mime="text/csv",
                use_container_width=True
            )


        with col2:

            excel_data = to_excel_bytes(
                filtered_df
            )


            st.download_button(
                "📗 Download Excel",
                data=excel_data,
                file_name="patient_records.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )


# ============================================================
# DATA ANALYSIS
# ============================================================

elif page == "📈 Data Analysis":

    st.markdown(
        '<div class="section-title">📈 Advanced Data Analysis</div>',
        unsafe_allow_html=True
    )


    if df.empty:

        st.warning(
            "No patient data available for analysis."
        )

    else:

        analysis_df = df.copy()


        analysis_df["Age"] = pd.to_numeric(
            analysis_df["Age"],
            errors="coerce"
        )


        analysis_df["Bill Amount"] = pd.to_numeric(
            analysis_df["Bill Amount"],
            errors="coerce"
        ).fillna(0)


        total_patients = len(
            analysis_df
        )


        average_age = analysis_df[
            "Age"
        ].mean()


        total_revenue = analysis_df[
            "Bill Amount"
        ].sum()


        average_bill = analysis_df[
            "Bill Amount"
        ].mean()


        col1, col2, col3, col4 = st.columns(4)


        with col1:

            st.metric(
                "👥 Total Patients",
                total_patients
            )


        with col2:

            st.metric(
                "🎂 Average Age",
                f"{average_age:.1f}"
                if pd.notna(average_age)
                else "0"
            )


        with col3:

            st.metric(
                "💰 Revenue",
                f"₹ {total_revenue:,.0f}"
            )


        with col4:

            st.metric(
                "🧾 Average Bill",
                f"₹ {average_bill:,.0f}"
            )


        st.divider()


        analysis_type = st.selectbox(
            "Select Analysis",
            [
                "Department-wise Patients",
                "Gender Distribution",
                "Top Diagnoses",
                "Department-wise Revenue"
            ]
        )


        chart_type = st.selectbox(
            "Select Chart Type",
            [
                "Bar Chart",
                "Pie Chart",
                "Line Chart"
            ]
        )


        if analysis_type == "Department-wise Patients":

            chart_data = (
                analysis_df["Department"]
                .value_counts()
                .reset_index()
            )

            chart_data.columns = [
                "Category",
                "Value"
            ]


        elif analysis_type == "Gender Distribution":

            chart_data = (
                analysis_df["Gender"]
                .value_counts()
                .reset_index()
            )

            chart_data.columns = [
                "Category",
                "Value"
            ]


        elif analysis_type == "Top Diagnoses":

            chart_data = (
                analysis_df["Diagnosis"]
                .value_counts()
                .head(10)
                .reset_index()
            )

            chart_data.columns = [
                "Category",
                "Value"
            ]


        else:

            chart_data = (
                analysis_df
                .groupby("Department")["Bill Amount"]
                .sum()
                .reset_index()
            )

            chart_data.columns = [
                "Category",
                "Value"
            ]


        if chart_type == "Bar Chart":

            fig = px.bar(
                chart_data,
                x="Category",
                y="Value",
                text="Value",
                title=analysis_type
            )


        elif chart_type == "Pie Chart":

            fig = px.pie(
                chart_data,
                names="Category",
                values="Value",
                hole=0.4,
                title=analysis_type
            )


        else:

            fig = px.line(
                chart_data,
                x="Category",
                y="Value",
                markers=True,
                title=analysis_type
            )


        fig.update_layout(
            template="plotly_dark",
            height=520,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Inter, Arial", color="#E9E7F2"),
            margin=dict(l=20, r=20, t=70, b=35),
            title_font=dict(size=18)
        )


        st.plotly_chart(
            fig,
            use_container_width=True
        )


        st.divider()


        st.subheader(
            "📋 Analysis Data"
        )


        st.dataframe(
            chart_data,
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# FOOTER
# ============================================================

st.divider()


st.markdown(
    """
    <div class="footer-wrap">
        <div class="footer-main">Medi Analysis • Hospital Intelligence System • Patient Management & Healthcare Analytics</div>
        <div class="footer-credit">Patented by : FAIZ SAJID SHAIKH</div>
    </div>
    """,
    unsafe_allow_html=True
)