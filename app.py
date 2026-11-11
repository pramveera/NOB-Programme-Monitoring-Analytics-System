
import streamlit as st
import pandas as pd
import plotly.express as px

from pathlib import Path
from io import BytesIO


# ==========================================
# 1. PAGE CONFIGURATION
# ==========================================

st.set_page_config(
    page_title="NOB Programme Dashboard",
    page_icon="📊",
    layout="wide"
)

FILE_PATH = Path("data/Pre_NOBs_Calls_Report 2.xlsx")

st.title("📊 NOB Programme Dashboard")
st.caption("Indian Youth Congress | Calls Analysis")


# ==========================================
# 2. LOAD AND CLEAN DATA
# ==========================================

if not FILE_PATH.exists():
    st.error("Excel file not found. Check your data folder.")
    st.stop()


@st.cache_data
def load_data(file_path):

    received = pd.read_excel(
        file_path,
        sheet_name="Received Calls"
    )

    not_received = pd.read_excel(
        file_path,
        sheet_name="Not Received Calls"
    )

    # Clean column names
    received.columns = received.columns.str.strip()
    not_received.columns = not_received.columns.str.strip()

    # Clean text values
    for df in [received, not_received]:

        for col in df.select_dtypes(include="object").columns:
            df[col] = df[col].str.strip()

    return received, not_received


received_df, not_received_df = load_data(FILE_PATH)


# ==========================================
# 3. HELPER FUNCTIONS
# ==========================================

def get_options(df, column):

    if column not in df.columns:
        return []

    return sorted(
        df[column].dropna().astype(str).unique().tolist()
    )


def convert_to_excel(df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Records"
        )

    return output.getvalue()


def display_value(value):

    if pd.isna(value):
        return "Not available"

    return str(value)


# ==========================================
# 4. INDIVIDUAL PROFILE POPUP
# ==========================================

@st.dialog("Individual Profile", width="large")
def show_profile(person):

    st.markdown("### 👤 Person Information")

    col1, col2 = st.columns(2)

    with col1:

        st.markdown("**Name**")
        st.write(display_value(person.get("Name")))

        st.markdown("**Gender**")
        st.write(display_value(person.get("Gender")))

        st.markdown("**State**")
        st.write(display_value(person.get("State")))

    with col2:

        st.markdown("**Party**")
        st.write(display_value(person.get("Party")))

        st.markdown("**Assembly**")
        st.write(display_value(person.get("Assembly")))

        st.markdown("**Help Required**")
        st.write(display_value(person.get("Help In")))

    st.divider()

    st.markdown("### 📋 Additional Details")

    col1, col2 = st.columns(2)

    with col1:

        for field in ["Name of the Election", "Legal help"]:

            if field in person.index:

                st.markdown(f"**{field}**")
                st.write(display_value(person[field]))

    with col2:

        if "Phone" in person.index:

            phone = person["Phone"]

            if pd.notna(phone):

                phone = str(phone)
                phone = "*" * max(0, len(phone) - 4) + phone[-4:]

            else:
                phone = "Not available"

            st.markdown("**Phone Number**")
            st.write(phone)

    st.caption("Personal contact information is masked.")


# ==========================================
# 5. SELECTABLE RECORDS TABLE
# ==========================================

def show_person_details(data, key):

    if data.empty:
        st.warning("No records available.")
        return

    # Hide phone number in table
    display_df = data.drop(
        columns=["Phone"],
        errors="ignore"
    ).copy()

    event = st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
        key=key
    )

    selected_rows = event.selection.rows

    if selected_rows:

        selected_index = selected_rows[0]

        person = data.iloc[selected_index]

        show_profile(person)

    else:

        st.info("Select a person's row to view their complete profile.")


# ==========================================
# 6. DASHBOARD NAVIGATION
# ==========================================

page = st.radio(
    "Navigation",
    [
        "Overview",
        "Received Calls",
        "Not Received Calls"
    ],
    horizontal=True
)


# ==========================================
# 7. OVERVIEW PAGE
# ==========================================

if page == "Overview":

    st.header("Overall Calls Overview")

    total_received = len(received_df)
    total_not_received = len(not_received_df)
    total_calls = total_received + total_not_received

    # KPI CARDS
    col1, col2, col3 = st.columns(3)

    col1.metric("Total Calls", f"{total_calls:,}")
    col2.metric("Received Calls", f"{total_received:,}")
    col3.metric("Not Received Calls", f"{total_not_received:,}")

    st.divider()

    # CALL STATUS CHART
    st.subheader("Call Status Distribution")

    status_df = pd.DataFrame({
        "Call Status": ["Received", "Not Received"],
        "Records": [total_received, total_not_received]
    })

    fig = px.pie(
        status_df,
        names="Call Status",
        values="Records",
        hole=0.45,
        title="Received vs Not Received"
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )

    # SUMMARY
    st.subheader("Overall Summary")

    st.write(
        f"Out of {total_calls:,} total contacts, "
        f"{total_received:,} calls were received and "
        f"{total_not_received:,} were not received."
    )


# ==========================================
# 8. RECEIVED CALLS PAGE
# ==========================================

elif page == "Received Calls":

    st.header("Received Calls Analysis")

    filtered_df = received_df.copy()

    # --------------------------------------
    # SEARCH BY NAME
    # --------------------------------------

    search_name = st.text_input(
        "🔍 Search Person by Name",
        placeholder="Enter person's name..."
    )

    if search_name and "Name" in filtered_df.columns:

        filtered_df = filtered_df[
            filtered_df["Name"].astype(str).str.contains(
                search_name,
                case=False,
                na=False
            )
        ]

    # --------------------------------------
    # SEARCH BY ASSEMBLY
    # --------------------------------------

    search_assembly = st.text_input(
        "🔍 Search by Assembly",
        placeholder="Enter assembly name..."
    )

    if search_assembly and "Assembly" in filtered_df.columns:

        filtered_df = filtered_df[
            filtered_df["Assembly"].astype(str).str.contains(
                search_assembly,
                case=False,
                na=False
            )
        ]

    # --------------------------------------
    # STATE AND PARTY FILTERS
    # --------------------------------------

    col1, col2 = st.columns(2)

    if "State" in filtered_df.columns:

        states = get_options(received_df, "State")

        with col1:

            selected_states = st.multiselect(
                "Filter by State",
                states
            )

        if selected_states:

            filtered_df = filtered_df[
                filtered_df["State"].astype(str).isin(selected_states)
            ]

    if "Party" in filtered_df.columns:

        parties = get_options(received_df, "Party")

        with col2:

            selected_parties = st.multiselect(
                "Filter by Party",
                parties
            )

        if selected_parties:

            filtered_df = filtered_df[
                filtered_df["Party"].astype(str).isin(selected_parties)
            ]

    # --------------------------------------
    # GENDER FILTER
    # --------------------------------------

    if "Gender" in filtered_df.columns:

        gender_options = ["All"] + get_options(
            received_df,
            "Gender"
        )

        selected_gender = st.selectbox(
            "Filter by Gender",
            gender_options
        )

        if selected_gender != "All":

            filtered_df = filtered_df[
                filtered_df["Gender"].astype(str) == selected_gender
            ]

    # --------------------------------------
    # HELP REQUIRED FILTER
    # --------------------------------------

    if "Help In" in filtered_df.columns:

        help_options = ["All"] + get_options(
            received_df,
            "Help In"
        )

        selected_help = st.selectbox(
            "Filter by Help Required",
            help_options
        )

        if selected_help != "All":

            filtered_df = filtered_df[
                filtered_df["Help In"].astype(str) == selected_help
            ]

    # --------------------------------------
    # ASSEMBLY DROPDOWN FILTER
    # --------------------------------------

    if "Assembly" in filtered_df.columns:

        assembly_options = ["All"] + get_options(
            received_df,
            "Assembly"
        )

        selected_assembly = st.selectbox(
            "Filter by Assembly",
            assembly_options
        )

        if selected_assembly != "All":

            filtered_df = filtered_df[
                filtered_df["Assembly"].astype(str) == selected_assembly
            ]

    # --------------------------------------
    # FILTERED KPI
    # --------------------------------------

    st.divider()

    st.subheader("Filtered Overview")

    col1, col2 = st.columns(2)

    col1.metric(
        "Matching Records",
        f"{len(filtered_df):,}"
    )

    col2.metric(
        "Unique States",
        filtered_df["State"].nunique()
        if "State" in filtered_df.columns else 0
    )

    # --------------------------------------
    # CATEGORY-WISE CHARTS
    # --------------------------------------

    st.divider()

    st.subheader("Category-wise Analysis")

    chart_columns = [
        "State",
        "Party",
        "Help In",
        "Gender"
    ]

    available_columns = [
        col for col in chart_columns
        if col in filtered_df.columns
    ]

    if filtered_df.empty:

        st.info("No records match the selected filters.")

    else:

        for i in range(0, len(available_columns), 2):

            col1, col2 = st.columns(2)

            for j, container in enumerate([col1, col2]):

                index = i + j

                if index >= len(available_columns):
                    continue

                column = available_columns[index]

                counts = (
                    filtered_df[column]
                    .dropna()
                    .astype(str)
                    .value_counts()
                    .head(10)
                    .reset_index()
                )

                counts.columns = [column, "Records"]

                fig = px.bar(
                    counts,
                    x=column,
                    y="Records",
                    title=f"Top 10 - {column}",
                    text="Records"
                )

                fig.update_layout(
                    template="plotly_white"
                )

                with container:

                    st.plotly_chart(
                        fig,
                        use_container_width=True
                    )

    # --------------------------------------
    # INDIVIDUAL RECORDS TABLE
    # --------------------------------------

    st.divider()

    st.subheader("Received Calls Records")

    show_person_details(
        filtered_df,
        key="received_records"
    )

    # --------------------------------------
    # EXPORT RECEIVED CALLS
    # --------------------------------------

    st.divider()

    st.subheader("📥 Download Received Calls")

    # Exclude names and phone numbers
    export_df = filtered_df.drop(
        columns=["Name", "Phone"],
        errors="ignore"
    )

    excel_data = convert_to_excel(export_df)

    st.download_button(
        label="Download Received Calls Excel",
        data=excel_data,
        file_name="NOB_Received_Calls.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=export_df.empty
    )


# ==========================================
# 9. NOT RECEIVED CALLS PAGE
# ==========================================

elif page == "Not Received Calls":

    st.header("Not Received Calls Analysis")

    total_not_received = len(not_received_df)

    st.metric(
        "Total Not Received Contacts",
        f"{total_not_received:,}"
    )

    st.info(
        "This worksheet contains contact records but "
        "does not provide additional categorical fields "
        "for meaningful distribution analysis."
    )

    # --------------------------------------
    # INDIVIDUAL RECORDS TABLE
    # --------------------------------------

    st.subheader("Not Received Contacts")

    show_person_details(
        not_received_df,
        key="not_received_records"
    )

    # --------------------------------------
    # EXPORT NOT RECEIVED CALLS
    # --------------------------------------

    st.divider()

    st.subheader("📥 Download Not Received Calls")

    export_df = not_received_df.drop(
        columns=["Name", "Phone"],
        errors="ignore"
    )

    excel_data = convert_to_excel(export_df)

    st.download_button(
        label="Download Not Received Calls Excel",
        data=excel_data,
        file_name="NOB_Not_Received_Calls.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=export_df.empty
    )

    st.caption(
        "Personal contact details are hidden in this dashboard."
    )


# ==========================================
# 10. FOOTER
# ==========================================

st.divider()

st.caption("NOB Programme Dashboard | Data Analysis Project")
