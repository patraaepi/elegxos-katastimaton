from __future__ import annotations

from datetime import datetime
from io import BytesIO
from urllib.parse import quote_plus

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Έλεγχος Καταστημάτων",
    page_icon="🔎",
    layout="wide",
)


COLUMNS = [
    "ΗΜΕΡΟΜΗΝΙΑ ΕΛΕΓΧΟΥ",
    "ΚΑΤΑΣΤΗΜΑ",
    "ΠΟΛΗ / ΠΕΡΙΟΧΗ",
    "ΔΙΕΥΘΥΝΣΗ",
    "ΤΗΛΕΦΩΝΟ",
    "EMAIL",
    "ΙΣΤΟΣΕΛΙΔΑ",
    "GOOGLE MAPS",
    "GOOGLE PHOTOS",
    "FACEBOOK",
    "INSTAGRAM",
    "TIKTOK",
    "FACEBOOK VIDEOS",
    "INSTAGRAM REELS",
    "TIKTOK VIDEOS",
    "TRIPADVISOR",
    "ΜΟΥΣΙΚΗ",
    "DJ",
    "LIVE",
    "ΚΑΤΑΣΤΑΣΗ",
    "ΒΑΘΜΟΣ ΒΕΒΑΙΟΤΗΤΑΣ",
    "ΠΑΡΑΤΗΡΗΣΕΙΣ",
]


def empty_history() -> pd.DataFrame:
    """Επιστρέφει κενό ιστορικό με τις σωστές στήλες."""
    return pd.DataFrame(columns=COLUMNS)


def clean_text(value: object) -> str:
    """Μετατρέπει ασφαλώς οποιαδήποτε τιμή σε καθαρό κείμενο."""
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def create_search_links(store_name: str, city: str) -> dict[str, str]:
    """Δημιουργεί οργανωμένα links δημόσιας αναζήτησης."""
    plain_query = f"{store_name} {city}".strip()
    query = quote_plus(plain_query)

    return {
        "GOOGLE PHOTOS": (
            f"https://www.google.com/search?tbm=isch&q={query}"
        ),
        "GOOGLE MAPS": (
            "https://www.google.com/maps/search/"
            f"?api=1&query={query}"
        ),
        "GOOGLE REVIEWS": (
            f"https://www.google.com/search?q={query}"
            "+κριτικές+μουσική+DJ+live"
        ),
        "EMAIL": (
            f"https://www.google.com/search?q={query}"
            "+email+επικοινωνία"
        ),
        "FACEBOOK": (
            "https://www.google.com/search?q="
            f"site%3Afacebook.com+{query}"
        ),
        "FACEBOOK VIDEOS": (
            "https://www.google.com/search?q="
            f"site%3Afacebook.com%2Freel+OR+"
            f"site%3Afacebook.com%2Fwatch+{query}"
        ),
        "INSTAGRAM": (
            "https://www.google.com/search?q="
            f"site%3Ainstagram.com+{query}"
        ),
        "INSTAGRAM REELS": (
            "https://www.google.com/search?q="
            f"site%3Ainstagram.com%2Freel+{query}"
        ),
        "TIKTOK": (
            "https://www.google.com/search?q="
            f"site%3Atiktok.com+{query}"
        ),
        "TRIPADVISOR": (
            "https://www.google.com/search?q="
            f"site%3Atripadvisor.com+{query}"
        ),
    }


def load_history(uploaded_file) -> pd.DataFrame:
    """Διαβάζει προηγούμενο Excel ή CSV ιστορικού."""
    if uploaded_file is None:
        return empty_history()

    filename = uploaded_file.name.lower()

    try:
        if filename.endswith(".xlsx"):
            dataframe = pd.read_excel(
                uploaded_file,
                sheet_name="ΙΣΤΟΡΙΚΟ",
            )
        elif filename.endswith(".csv"):
            dataframe = pd.read_csv(uploaded_file)
        else:
            st.error("Υποστηρίζονται μόνο αρχεία Excel και CSV.")
            return empty_history()

        for column in COLUMNS:
            if column not in dataframe.columns:
                dataframe[column] = ""

        return dataframe[COLUMNS].fillna("")

    except Exception as error:
        st.error(f"Δεν ήταν δυνατή η ανάγνωση του αρχείου: {error}")
        return empty_history()


def create_excel(history: pd.DataFrame) -> bytes:
    """Δημιουργεί Excel με πολλά οργανωμένα φύλλα."""
    output = BytesIO()

    history_export = history.copy()

    social_columns = [
        "ΗΜΕΡΟΜΗΝΙΑ ΕΛΕΓΧΟΥ",
        "ΚΑΤΑΣΤΗΜΑ",
        "ΠΟΛΗ / ΠΕΡΙΟΧΗ",
        "FACEBOOK",
        "INSTAGRAM",
        "TIKTOK",
    ]

    video_columns = [
        "ΗΜΕΡΟΜΗΝΙΑ ΕΛΕΓΧΟΥ",
        "ΚΑΤΑΣΤΗΜΑ",
        "ΠΟΛΗ / ΠΕΡΙΟΧΗ",
        "FACEBOOK VIDEOS",
        "INSTAGRAM REELS",
        "TIKTOK VIDEOS",
    ]

    music_columns = [
        "ΗΜΕΡΟΜΗΝΙΑ ΕΛΕΓΧΟΥ",
        "ΚΑΤΑΣΤΗΜΑ",
        "ΠΟΛΗ / ΠΕΡΙΟΧΗ",
        "ΜΟΥΣΙΚΗ",
        "DJ",
        "LIVE",
        "ΒΑΘΜΟΣ ΒΕΒΑΙΟΤΗΤΑΣ",
        "ΠΑΡΑΤΗΡΗΣΕΙΣ",
    ]

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:
        history_export.to_excel(
            writer,
            sheet_name="ΙΣΤΟΡΙΚΟ",
            index=False,
        )

        history_export[social_columns].to_excel(
            writer,
            sheet_name="SOCIAL MEDIA",
            index=False,
        )

        history_export[video_columns].to_excel(
            writer,
            sheet_name="REELS - VIDEOS",
            index=False,
        )

        history_export[music_columns].to_excel(
            writer,
            sheet_name="ΜΟΥΣΙΚΗ - DJ",
            index=False,
        )

        for sheet_name in writer.book.sheetnames:
            worksheet = writer.book[sheet_name]
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions

            for column_cells in worksheet.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter

                for cell in column_cells:
                    cell_value = clean_text(cell.value)
                    max_length = max(max_length, len(cell_value))

                worksheet.column_dimensions[column_letter].width = min(
                    max(max_length + 2, 12),
                    55,
                )

    output.seek(0)
    return output.getvalue()


def create_record(
    store_name: str,
    city: str,
    links: dict[str, str],
    form_data: dict[str, str],
) -> dict[str, str]:
    """Δημιουργεί μία ολοκληρωμένη εγγραφή ελέγχου."""
    return {
        "ΗΜΕΡΟΜΗΝΙΑ ΕΛΕΓΧΟΥ": datetime.now().strftime(
            "%d/%m/%Y %H:%M"
        ),
        "ΚΑΤΑΣΤΗΜΑ": store_name.upper(),
        "ΠΟΛΗ / ΠΕΡΙΟΧΗ": city.upper(),
        "ΔΙΕΥΘΥΝΣΗ": form_data["address"],
        "ΤΗΛΕΦΩΝΟ": form_data["phone"],
        "EMAIL": form_data["email"],
        "ΙΣΤΟΣΕΛΙΔΑ": form_data["website"],
        "GOOGLE MAPS": links["GOOGLE MAPS"],
        "GOOGLE PHOTOS": links["GOOGLE PHOTOS"],
        "FACEBOOK": form_data["facebook"],
        "INSTAGRAM": form_data["instagram"],
        "TIKTOK": form_data["tiktok"],
        "FACEBOOK VIDEOS": form_data["facebook_videos"],
        "INSTAGRAM REELS": form_data["instagram_reels"],
        "TIKTOK VIDEOS": form_data["tiktok_videos"],
        "TRIPADVISOR": form_data["tripadvisor"],
        "ΜΟΥΣΙΚΗ": form_data["music"],
        "DJ": form_data["dj"],
        "LIVE": form_data["live"],
        "ΚΑΤΑΣΤΑΣΗ": form_data["status"],
        "ΒΑΘΜΟΣ ΒΕΒΑΙΟΤΗΤΑΣ": form_data["confidence"],
        "ΠΑΡΑΤΗΡΗΣΕΙΣ": form_data["notes"],
    }


if "history" not in st.session_state:
    st.session_state.history = empty_history()

if "current_search" not in st.session_state:
    st.session_state.current_search = None


st.title("🔎 ΕΛΕΓΧΟΣ ΚΑΤΑΣΤΗΜΑΤΩΝ")

st.info(
    "Αν έχεις προηγούμενο αρχείο ιστορικού, ανέβασέ το πριν "
    "προσθέσεις νέους ελέγχους."
)


# ---------------------------------------------------------
# ΦΟΡΤΩΣΗ ΠΡΟΗΓΟΥΜΕΝΟΥ ΙΣΤΟΡΙΚΟΥ
# ---------------------------------------------------------

with st.expander("📂 ΦΟΡΤΩΣΗ ΠΡΟΗΓΟΥΜΕΝΟΥ ΙΣΤΟΡΙΚΟΥ"):
    uploaded_history = st.file_uploader(
        "Ανέβασε το προηγούμενο Excel ή CSV",
        type=["xlsx", "csv"],
    )

    if uploaded_history is not None:
        if st.button("ΦΟΡΤΩΣΗ ΑΡΧΕΙΟΥ"):
            st.session_state.history = load_history(uploaded_history)
            st.success(
                f"Φορτώθηκαν "
                f"{len(st.session_state.history)} εγγραφές."
            )


# ---------------------------------------------------------
# ΝΕΑ ΑΝΑΖΗΤΗΣΗ
# ---------------------------------------------------------

st.header("1. ΝΕΟΣ ΕΛΕΓΧΟΣ")

with st.form("search_form"):
    col1, col2 = st.columns(2)

    with col1:
        store_name = st.text_input(
            "Όνομα καταστήματος",
            placeholder="π.χ. WALLSTREET",
        )

    with col2:
        city = st.text_input(
            "Πόλη / περιοχή",
            placeholder="π.χ. ΑΓΙΑ ΠΑΡΑΣΚΕΥΗ",
        )

    start_search = st.form_submit_button(
        "ΕΝΑΡΞΗ ΕΛΕΓΧΟΥ",
        type="primary",
        use_container_width=True,
    )


if start_search:
    if not store_name.strip() or not city.strip():
        st.error(
            "Συμπλήρωσε το όνομα του καταστήματος "
            "και την πόλη."
        )
    else:
        normalized_store = store_name.strip().upper()
        normalized_city = city.strip().upper()

        links = create_search_links(
            normalized_store,
            normalized_city,
        )

        st.session_state.current_search = {
            "store_name": normalized_store,
            "city": normalized_city,
            "links": links,
        }


# ---------------------------------------------------------
# ΑΠΟΤΕΛΕΣΜΑΤΑ ΚΑΙ ΕΠΙΒΕΒΑΙΩΣΗ
# ---------------------------------------------------------

if st.session_state.current_search:
    current = st.session_state.current_search
    links = current["links"]
    current_store = current["store_name"]
    current_city = current["city"]

    st.success(f"Έλεγχος: {current_store} – {current_city}")

    st.header("📸 ΦΩΤΟΓΡΑΦΙΕΣ GOOGLE")

    st.link_button(
        "ΑΝΟΙΓΜΑ ΦΩΤΟΓΡΑΦΙΩΝ GOOGLE",
        links["GOOGLE PHOTOS"],
        use_container_width=True,
    )

    st.header("🎥 REELS ΚΑΙ ΒΙΝΤΕΟ")

    video_col1, video_col2, video_col3 = st.columns(3)

    with video_col1:
        st.link_button(
            "INSTAGRAM REELS",
            links["INSTAGRAM REELS"],
            use_container_width=True,
        )

    with video_col2:
        st.link_button(
            "FACEBOOK VIDEOS",
            links["FACEBOOK VIDEOS"],
            use_container_width=True,
        )

    with video_col3:
        st.link_button(
            "TIKTOK VIDEOS",
            links["TIKTOK"],
            use_container_width=True,
        )

    st.header("🌐 ΑΝΑΖΗΤΗΣΗ ΣΤΟΙΧΕΙΩΝ")

    search_col1, search_col2, search_col3 = st.columns(3)

    with search_col1:
        st.link_button(
            "GOOGLE MAPS",
            links["GOOGLE MAPS"],
            use_container_width=True,
        )
        st.link_button(
            "ΑΝΑΖΗΤΗΣΗ EMAIL",
            links["EMAIL"],
            use_container_width=True,
        )

    with search_col2:
        st.link_button(
            "FACEBOOK",
            links["FACEBOOK"],
            use_container_width=True,
        )
        st.link_button(
            "INSTAGRAM",
            links["INSTAGRAM"],
            use_container_width=True,
        )

    with search_col3:
        st.link_button(
            "TIKTOK",
            links["TIKTOK"],
            use_container_width=True,
        )
        st.link_button(
            "TRIPADVISOR",
            links["TRIPADVISOR"],
            use_container_width=True,
        )

    st.link_button(
        "GOOGLE REVIEWS – ΜΟΥΣΙΚΗ / DJ / LIVE",
        links["GOOGLE REVIEWS"],
        use_container_width=True,
    )

    st.header("2. ΕΠΙΒΕΒΑΙΩΣΗ ΚΑΙ ΣΥΜΠΛΗΡΩΣΗ")

    with st.form("confirmation_form", clear_on_submit=False):
        basic_col1, basic_col2 = st.columns(2)

        with basic_col1:
            address = st.text_input("Διεύθυνση")
            phone = st.text_input("Τηλέφωνο")
            email = st.text_input("Email")
            website = st.text_input("Ιστοσελίδα")

        with basic_col2:
            facebook = st.text_input("Facebook link")
            instagram = st.text_input("Instagram link")
            tiktok = st.text_input("TikTok link")
            tripadvisor = st.text_input("Tripadvisor link")

        st.subheader("ΒΙΝΤΕΟ ΚΑΙ REELS")

        facebook_videos = st.text_area(
            "Facebook videos",
            placeholder=(
                "Βάλε ένα ή περισσότερα links, "
                "ένα σε κάθε γραμμή."
            ),
        )

        instagram_reels = st.text_area(
            "Instagram Reels",
            placeholder=(
                "Βάλε ένα ή περισσότερα links, "
                "ένα σε κάθε γραμμή."
            ),
        )

        tiktok_videos = st.text_area(
            "TikTok videos",
            placeholder=(
                "Βάλε ένα ή περισσότερα links, "
                "ένα σε κάθε γραμμή."
            ),
        )

        st.subheader("ΕΛΕΓΧΟΣ ΜΟΥΣΙΚΗΣ")

        result_col1, result_col2, result_col3 = st.columns(3)

        with result_col1:
            music = st.selectbox(
                "ΜΟΥΣΙΚΗ",
                ["ΔΕΝ ΕΛΕΓΧΘΗΚΕ", "ΝΑΙ", "ΟΧΙ"],
            )

        with result_col2:
            dj = st.selectbox(
                "DJ",
                ["ΔΕΝ ΕΛΕΓΧΘΗΚΕ", "ΝΑΙ", "ΟΧΙ"],
            )

        with result_col3:
            live = st.selectbox(
                "LIVE",
                ["ΔΕΝ ΕΛΕΓΧΘΗΚΕ", "ΝΑΙ", "ΟΧΙ"],
            )

        status = st.selectbox(
            "Κατάσταση ελέγχου",
            [
                "ΝΕΟ",
                "ΥΠΟ ΕΛΕΓΧΟ",
                "ΟΛΟΚΛΗΡΩΜΕΝΟ",
                "ΧΡΕΙΑΖΕΤΑΙ ΕΠΑΝΕΛΕΓΧΟ",
                "ΔΕΝ ΕΝΤΟΠΙΣΤΗΚΕ",
                "ΕΚΛΕΙΣΕ",
            ],
        )

        confidence = st.selectbox(
            "Βαθμός βεβαιότητας",
            [
                "ΕΠΙΒΕΒΑΙΩΜΕΝΟ",
                "ΠΙΘΑΝΟ",
                "ΜΗ ΕΠΙΒΕΒΑΙΩΜΕΝΟ",
                "ΛΑΝΘΑΣΜΕΝΟ",
            ],
        )

        notes = st.text_area(
            "Παρατηρήσεις",
            placeholder=(
                "Γράψε τις αναφορές που βρέθηκαν, "
                "την πηγή και οτιδήποτε χρειάζεται έλεγχο."
            ),
        )

        save_record = st.form_submit_button(
            "ΑΠΟΘΗΚΕΥΣΗ ΕΛΕΓΧΟΥ",
            type="primary",
            use_container_width=True,
        )

    if save_record:
        form_data = {
            "address": address.strip(),
            "phone": phone.strip(),
            "email": email.strip(),
            "website": website.strip(),
            "facebook": facebook.strip(),
            "instagram": instagram.strip(),
            "tiktok": tiktok.strip(),
            "facebook_videos": facebook_videos.strip(),
            "instagram_reels": instagram_reels.strip(),
            "tiktok_videos": tiktok_videos.strip(),
            "tripadvisor": tripadvisor.strip(),
            "music": music,
            "dj": dj,
            "live": live,
            "status": status,
            "confidence": confidence,
            "notes": notes.strip(),
        }

        new_record = create_record(
            current_store,
            current_city,
            links,
            form_data,
        )

        new_dataframe = pd.DataFrame([new_record])

        st.session_state.history = pd.concat(
            [
                st.session_state.history,
                new_dataframe,
            ],
            ignore_index=True,
        )

        st.success("Ο έλεγχος αποθηκεύτηκε στο ιστορικό.")


# ---------------------------------------------------------
# ΙΣΤΟΡΙΚΟ
# ---------------------------------------------------------

st.header("3. ΙΣΤΟΡΙΚΟ ΕΛΕΓΧΩΝ")

history = st.session_state.history.copy()

if history.empty:
    st.warning("Δεν υπάρχουν ακόμη αποθηκευμένοι έλεγχοι.")
else:
    search_history = st.text_input(
        "Αναζήτηση στο ιστορικό",
        placeholder="Γράψε κατάστημα, πόλη, email ή τηλέφωνο",
    )

    filtered_history = history.copy()

    if search_history.strip():
        search_value = search_history.strip().casefold()

        mask = filtered_history.astype(str).apply(
            lambda row: row.str.casefold().str.contains(
                search_value,
                regex=False,
            ).any(),
            axis=1,
        )

        filtered_history = filtered_history[mask]

    st.write(
        f"Εμφανίζονται {len(filtered_history)} "
        f"από {len(history)} εγγραφές."
    )

    st.dataframe(
        filtered_history,
        use_container_width=True,
        hide_index=True,
    )

    duplicate_columns = [
        "ΚΑΤΑΣΤΗΜΑ",
        "ΠΟΛΗ / ΠΕΡΙΟΧΗ",
    ]

    duplicates = history[
        history.duplicated(
            subset=duplicate_columns,
            keep=False,
        )
    ]

    if not duplicates.empty:
        with st.expander(
            f"⚠️ ΠΙΘΑΝΕΣ ΔΙΠΛΟΕΓΓΡΑΦΕΣ: {len(duplicates)}"
        ):
            st.dataframe(
                duplicates,
                use_container_width=True,
                hide_index=True,
            )

    st.header("📥 ΕΞΑΓΩΓΗ ΠΛΗΡΟΥΣ EXCEL")

    excel_file = create_excel(history)

    st.download_button(
        label="ΛΗΨΗ ΕΝΗΜΕΡΩΜΕΝΟΥ EXCEL",
        data=excel_file,
        file_name="ΙΣΤΟΡΙΚΟ_ΕΛΕΓΧΟΥ_ΚΑΤΑΣΤΗΜΑΤΩΝ.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True,
    )

    csv_file = history.to_csv(
        index=False,
    ).encode("utf-8-sig")

    st.download_button(
        label="ΛΗΨΗ CSV",
        data=csv_file,
        file_name="ΙΣΤΟΡΙΚΟ_ΕΛΕΓΧΩΝ.csv",
        mime="text/csv",
        use_container_width=True,
    )

    if st.button(
        "ΔΙΑΓΡΑΦΗ ΠΡΟΣΩΡΙΝΟΥ ΙΣΤΟΡΙΚΟΥ",
        use_container_width=True,
    ):
        st.session_state.history = empty_history()
        st.session_state.current_search = None
        st.rerun()
