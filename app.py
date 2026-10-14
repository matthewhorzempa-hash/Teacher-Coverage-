import re
import openpyxl  # Required for Excel reading in Streamlit Playground
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Teacher Coverage Assistant", page_icon="🏫", layout="wide"
)

PERIODS = [
    "Period 1",
    "Period 2",
    "Period 3",
    "Period 4/5",
    "Period 5/6",
    "Period 6/7",
    "Period 7/8",
    "Period 8/9",
    "Period 9/10",
    "Period 11",
    "Period 12",
    "Homeroom",
]

# Specific teacher exclusions per period (No coverage required at all)
EXCLUSIONS = {
    "Period 4": [
        "Scullion",
        "Burkley",
        "Kristen",
        "Harshman",
        "Null",
        "Musto",
        "Weightman",
        "Parker",
        "Horrell",
        "Hamed",
    ],
    "Period 6": [
        "Conlon",
        "Pondexter",
        "Conners",
        "Lubinsky",
        "M. Gouker",
        "Santoro",
        "Flatt",
        "Babko",
        "Fike",
        "Drennen",
        "Kapr",
    ],
    "Period 8": [
        "Espejo",
        "Herrod",
        "Reda",
        "Grice",
        "Sarver",
        "Boggs",
        "McConnell",
        "Flaus",
        "Leonard",
        "Reese",
        "Barbao",
        "Surovchak",
        "Getsy",
        "Fecik",
        "Horzempa",
        "Farquhar",
    ],
    "Period 10": [
        "Tatar",
        "Coneybeer",
        "Rodgers",
        "Brodeur",
        "J. Gouker",
        "Kenney",
        "Mendicino",
        "Jenkins",
        "Edwards",
        "Pappas",
    ],
}

# Specific classes that route to Study Hall if Duty coverage is unavailable (DO NOT use Prep)
STUDY_HALL_FALLBACKS = {
    "Herrod": ["2"],
    "Conlon": ["7", "8", "12"],
    "Burkley": ["1", "9", "10"],
    "Scullion": ["3", "11"],
    "Espejo": ["1", "3", "6", "7", "11"],
    "Pappas": ["6", "7", "8", "9"],
    "Kapr": ["4", "5"],
    "Reda": ["11"],
    "Barbao": ["2"],
}

# Initialize rotational prep pointers in session state
if "prep_pointers" not in st.session_state:
    st.session_state.prep_pointers = {p: 0 for p in PERIODS}


def is_teacher_excluded_for_period(teacher_name, period_str):
    """Checks if a teacher is explicitly listed in the exclusion rules for a given period."""
    p_nums = re.findall(r"\d+", period_str)
    for ex_p, t_list in EXCLUSIONS.items():
        ex_nums = re.findall(r"\d+", ex_p)
        if any(num in p_nums for num in ex_nums):
            if teacher_name in t_list:
                return True
    return False


def is_study_hall_fallback(teacher_name, period_str):
    """Checks if an unassigned class period should route to Study Hall instead of using a Prep period."""
    if teacher_name not in STUDY_HALL_FALLBACKS:
        return False
    allowed_nums = STUDY_HALL_FALLBACKS[teacher_name]
    p_nums = re.findall(r"\d+", period_str)
    return any(num in allowed_nums for num in p_nums)


def load_schedule(file_obj):
    """Loads and parses the uploaded Excel spreadsheet."""
    try:
        xls = pd.ExcelFile(file_obj)
        df = pd.read_excel(xls, "Sheet1")
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return None, [], {}

    current_period = None
    records = []
    all_teachers = set()

    for idx, row in df.iterrows():
        p = row["Unnamed: 0"]
        if pd.notna(p):
            current_period = str(p).strip()

        prep = str(row["Prep"]).strip() if pd.notna(row["Prep"]) else None
        hall = (
            str(row["Hall Duty"]).strip()
            if pd.notna(row["Hall Duty"])
            else None
        )
        lunch = (
            str(row["Lunch Duty"]).strip()
            if pd.notna(row["Lunch Duty"])
            else None
        )
        study = (
            str(row["Study Halls"]).strip()
            if pd.notna(row["Study Halls"])
            else None
        )

        if any([prep, hall, lunch, study]):
            records.append(
                {
                    "Period": current_period,
                    "Prep": prep,
                    "Hall Duty": hall,
                    "Lunch Duty": lunch,
                    "Study Halls": study,
                }
            )

        for col_val in [prep, hall, lunch, study]:
            if col_val and col_val != "None":
                parts = re.split(r"[/]", col_val)
                for part in parts:
                    clean_name = re.sub(
                        r"\s*\((Sm\.\s*2\vert{}Odd\vert{}Even)\)", "", part
                    ).strip()
                    if clean_name:
                        all_teachers.add(clean_name)

    df_parsed = pd.DataFrame(records)

    # Extract Bottom-Up Prep Lists per period (reversed from spreadsheet order)
    prep_lists = {}
    for p in PERIODS:
        p_rows = df_parsed[df_parsed["Period"].str.lower() == p.lower()]
        p_preps = [
            str(x).strip()
            for x in p_rows["Prep"]
            if pd.notna(x) and str(x).strip() != "None"
        ]
        prep_lists[p] = p_preps[::-1]  # Reverse to start from bottom

    return df_parsed, sorted(list(all_teachers)), prep_lists


def is_prep_period(
    df_parsed, teacher_name, period, semester_val=1, day_type="Odd"
):
    """Checks if a specific period is a scheduled Prep period for the teacher."""
    p_rows = df_parsed[df_parsed["Period"].str.lower() == period.lower()]
    for val in p_rows["Prep"]:
        if pd.notna(val) and str(val) != "None":
            clean_names = [
                re.sub(r"\s*\((Sm\.\s*2\vert{}Odd\vert{}Even)\)", "", x).strip()
                for x in re.split(r"[/]", str(val))
            ]
            if teacher_name in clean_names:
                if "Sm. 2" in str(val) and semester_val != 2:
                    continue
                if "Odd" in str(val) and day_type != "Odd":
                    continue
                if "Even" in str(val) and day_type != "Even":
                    continue
                return True
    return False


def get_coverage_needed_periods(
    df_parsed, teacher_name, semester_val=1, day_type="Odd"
):
    """Finds all periods an absent teacher requires coverage for (excluding Prep & explicit exclusions)."""
    coverage_periods = []
    for p in PERIODS:
        if is_prep_period(
            df_parsed,
            teacher_name,
            p,
            semester_val=semester_val,
            day_type=day_type,
        ):
            continue

        if is_teacher_excluded_for_period(teacher_name, p):
            continue

        coverage_periods.append(p)
    return coverage_periods


def generate_coverage(
    df_parsed,
    prep_lists,
    absent_no_sub,
    absent_with_sub,
    semester_str,
    day_type,
):
    """Assigns coverage using strict hierarchy:

    1. Full-Day Sub Free Period
    2. Hall Duty Staff
    3. Lunch Duty Staff
    4. Study Hall Fallback (if configured for teacher/period)
    5. Rotational Bottom-Up Prep (last resort)
    """
    semester_val = 2 if "2" in str(semester_str) else 1
    period_assigned_coverers = {p: set() for p in PERIODS}
    coverage_rows = []

    all_absent = set(absent_no_sub + absent_with_sub)

    for abs_teacher in absent_no_sub:
        t_periods = get_coverage_needed_periods(
            df_parsed, abs_teacher, semester_val=semester_val, day_type=day_type
        )

        for p in t_periods:
            p_rows = df_parsed[df_parsed["Period"].str.lower() == p.lower()]
            assigned_cover = None
            source = None

            # Priority 1: Full-Day Substitute Availability
            for sub_teacher in absent_with_sub:
                sub_label = f"Sub for {sub_teacher}"
                if sub_label in period_assigned_coverers[p]:
                    continue

                has_free_period = False
                for col in ["Hall Duty", "Lunch Duty", "Prep"]:
                    for val in p_rows[col]:
                        if val and str(val) != "None":
                            clean_names = [
                                re.sub(
                                    r"\s*\((Sm\.\s*2\vert{}Odd\vert{}Even)\)", "", x
                                ).strip()
                                for x in re.split(r"[/]", str(val))
                            ]
                            if sub_teacher in clean_names:
                                if "Sm. 2" in str(val) and semester_val != 2:
                                    continue
                                if "Odd" in str(val) and day_type != "Odd":
                                    continue
                                if "Even" in str(val) and day_type != "Even":
                                    continue
                                has_free_period = True
                                break
                    if has_free_period:
                        break

                if has_free_period:
                    assigned_cover = sub_label
                    source = f"All-Day Sub ({sub_teacher}'s Free Period)"
                    period_assigned_coverers[p].add(sub_label)
                    break

            # Priority 2: Hall Duty
            if not assigned_cover:
                for _, row in p_rows.iterrows():
                    val = row["Hall Duty"]
                    if val and str(val) != "None":
                        if "Sm. 2" in str(val) and semester_val != 2:
                            continue
                        if "Odd" in str(val) and day_type != "Odd":
                            continue
                        if "Even" in str(val) and day_type != "Even":
                            continue

                        clean_names = [
                            re.sub(r"\s*\((Sm\.\s*2\vert{}Odd\vert{}Even)\)", "", x).strip()
                            for x in re.split(r"[/]", str(val))
                        ]
                        for clean_name in clean_names:
                            if (
                                clean_name not in all_absent
                                and clean_name
                                not in period_assigned_coverers[p]
                            ):
                                assigned_cover = clean_name
                                source = "Hall Duty"
                                period_assigned_coverers[p].add(clean_name)
                                break
                        if assigned_cover:
                            break

            # Priority 3: Lunch Duty
            if not assigned_cover:
                for _, row in p_rows.iterrows():
                    val = row["Lunch Duty"]
                    if val and str(val) != "None":
                        if "Sm. 2" in str(val) and semester_val != 2:
                            continue
                        if "Odd" in str(val) and day_type != "Odd":
                            continue
                        if "Even" in str(val) and day_type != "Even":
                            continue

                        clean_names = [
                            re.sub(r"\s*\((Sm\.\s*2\vert{}Odd\vert{}Even)\)", "", x).strip()
                            for x in re.split(r"[/]", str(val))
                        ]
                        for clean_name in clean_names:
                            if (
                                clean_name not in all_absent
                                and clean_name
                                not in period_assigned_coverers[p]
                            ):
                                assigned_cover = clean_name
                                source = "Lunch Duty"
                                period_assigned_coverers[p].add(clean_name)
                                break
                        if assigned_cover:
                            break

            # Priority 4: Study Hall Fallback (Avoids pulling a teacher off Prep)
            if not assigned_cover and is_study_hall_fallback(abs_teacher, p):
                assigned_cover = "🏫 SENT TO STUDY HALL"
                source = "Study Hall Fallback (No Prep Used)"

            # Priority 5: Prep (Bottom-Up Rotational / Cycling)
            if not assigned_cover and p in prep_lists and len(prep_lists[p]) > 0:
                p_prep_list = prep_lists[p]
                n = len(p_prep_list)
                start_idx = st.session_state.prep_pointers.get(p, 0)

                for k in range(n):
                    cand_idx = (start_idx + k) % n
                    candidate = p_prep_list[cand_idx]

                    if (
                        candidate not in all_absent
                        and candidate not in period_assigned_coverers[p]
                    ):
                        assigned_cover = candidate
                        source = (
                            f"Prep (Rotational Position #{cand_idx + 1} of {n})"
                        )
                        period_assigned_coverers[p].add(candidate)
                        st.session_state.prep_pointers[p] = (cand_idx + 1) % n
                        break

            if not assigned_cover:
                assigned_cover = "⚠️ UNCOVERED / EXTERNAL SUB"
                source = "None Available"

            coverage_rows.append(
                {
                    "Absent Teacher": abs_teacher,
                    "Period": p,
                    "Assigned Covering Teacher": assigned_cover,
                    "Coverage Source": source,
                }
            )

    return pd.DataFrame(coverage_rows)


# --- STREAMLIT USER INTERFACE ---
st.title("🏫 High School Teacher Coverage Assistant")
st.markdown(
    "Automated coverage generator utilizing **Full-Day Subs**, **Hall Duty**, **Lunch Duty**, **Study Hall Routing**, & **Rotational Prep**."
)

uploaded_file = st.file_uploader(
    "Upload Schedule Spreadsheet (Horzempa.xlsx)", type=["xlsx"]
)

if uploaded_file is not None:
    df_parsed, teacher_list, prep_lists = load_schedule(uploaded_file)

    if df_parsed is not None and not df_parsed.empty:
        col_left, col_right = st.columns([1, 2])

        with col_left:
            st.subheader("Daily Absence Inputs")

            absent_no_sub = st.multiselect(
                "1. Absent Teachers (NEEDING COVERAGE):",
                options=teacher_list,
                help="Select teachers who are absent and do NOT have an all-day substitute.",
            )

            remaining_teachers = [
                t for t in teacher_list if t not in absent_no_sub
            ]

            absent_with_sub = st.multiselect(
                "2. Absent Teachers (COVERED BY ALL-DAY SUB):",
                options=remaining_teachers,
                help="Select teachers who are absent BUT have a dedicated full-day external substitute.",
            )

            semester = st.radio(
                "Semester:", ["Semester 1", "Semester 2"], index=0
            )

            day_type = st.radio("Day Schedule:", ["Odd", "Even"], index=0)

            st.divider()
            if st.button("🔄 Reset Prep Rotation Pointers"):
                st.session_state.prep_pointers = {p: 0 for p in PERIODS}
                st.success("Rotation pointers reset back to the bottom!")

        with col_right:
            st.subheader("Coverage Results")

            if absent_no_sub:
                res_df = generate_coverage(
                    df_parsed,
                    prep_lists,
                    absent_no_sub,
                    absent_with_sub,
                    semester,
                    day_type,
                )

                uncovered_count = (
                    len(
                        res_df[
                            res_df["Assigned Covering Teacher"].str.contains(
                                "UNCOVERED"
                            )
                        ]
                    )
                    if not res_df.empty
                    else 0
                )

                st.markdown(
                    f"**Coverage Needed For:** {', '.join(absent_no_sub)}  \n"
                    f"**All-Day Subs Active:** {', '.join(absent_with_sub) if absent_with_sub else 'None'}  \n"
                    f"**Settings:** {semester} | {day_type} Day  \n"
                    f"**Total Periods Processed:** {len(res_df)}"
                )

                if uncovered_count > 0:
                    st.warning(
                        f"⚠️ {uncovered_count} period(s) require an additional external substitute."
                    )
                else:
                    st.success("✅ All periods successfully assigned!")

                st.dataframe(res_df, use_container_width=True, hide_index=True)

                csv_data = res_df.to_csv(index=False).encode("utf-8")
                st.download_button(
                    label="📥 Download Coverage Schedule (CSV)",
                    data=csv_data,
                    file_name="daily_teacher_coverage.csv",
                    mime="text/csv",
                )
            else:
                st.info(
                    "Select one or more absent teachers needing coverage to display assignments."
                )
else:
    st.info("👆 Please upload `Horzempa.xlsx` above to begin.")
