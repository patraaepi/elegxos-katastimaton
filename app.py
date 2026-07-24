from urllib.parse import quote_plus

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Έλεγχος Καταστημάτων",
    page_icon="🔎",
    layout="wide",
)


def create_search_links(store_name: str, city: str) -> dict[str, str]:
    """Δημιουργεί συνδέσμους αναζήτησης για το κατάστημα."""

    query = quote_plus(f"{store_name} {city}")

    return {
        "Google Photos": (
            "https://www.google.com/search?"
            f"tbm=isch&q={query}"
        ),
        "Google Maps": (
            "https://www.google.com/maps/search/"
            f"?api=1&query={query}"
        ),
        "Google Reviews": (
            "https://www.google.com/search?"
            f"q={query}+κριτικές+μουσική+DJ"
        ),
        "Facebook": (
            "https://www.google.com/search?"
            f"q=site%3Afacebook.com+{query}"
        ),
        "Facebook Videos": (
            "https://www.google.com/search?"
            f"q=site%3Afacebook.com%2Fwatch+OR+"
            f"site%3Afacebook.com%2Freel+{query}"
        ),
        "Instagram": (
            "https://www.google.com/search?"
            f"q=site%3Ainstagram.com+{query}"
        ),
        "Instagram Reels": (
            "https://www.google.com/search?"
            f"q=site%3Ainstagram.com%2Freel+{query}"
        ),
        "TikTok": (
            "https://www.google.com/search?"
            f"q=site%3Atiktok.com+{query}"
        ),
        "Tripadvisor": (
            "https://www.google.com/search?"
            f"q=site%3Atripadvisor.com+{query}"
        ),
        "Email": (
            "https://www.google.com/search?"
            f"q={query}+email+επικοινωνία"
        ),
    }


st.title("🔎 ΕΛΕΓΧΟΣ ΚΑΤΑΣΤΗΜΑΤΩΝ")

st.write(
    "Γράψε μόνο το όνομα του καταστήματος και την πόλη ή περιοχή."
)

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

search_button = st.button(
    "ΕΝΑΡΞΗ ΕΛΕΓΧΟΥ",
    type="primary",
    use_container_width=True,
)

if search_button:
    if not store_name.strip() or not city.strip():
        st.error(
            "Συμπλήρωσε το όνομα του καταστήματος και την πόλη."
        )
    else:
        store_name = store_name.strip().upper()
        city = city.strip().upper()

        links = create_search_links(store_name, city)

        st.success(f"Έλεγχος: {store_name} – {city}")

        st.header("📸 ΦΩΤΟΓΡΑΦΙΕΣ GOOGLE")
        st.link_button(
            "ΑΝΟΙΓΜΑ ΦΩΤΟΓΡΑΦΙΩΝ GOOGLE",
            links["Google Photos"],
            use_container_width=True,
        )

        st.header("🎥 REELS ΚΑΙ ΒΙΝΤΕΟ")

        video_columns = st.columns(3)

        with video_columns[0]:
            st.link_button(
                "INSTAGRAM REELS",
                links["Instagram Reels"],
                use_container_width=True,
            )

        with video_columns[1]:
            st.link_button(
                "FACEBOOK VIDEOS",
                links["Facebook Videos"],
                use_container_width=True,
            )

        with video_columns[2]:
            st.link_button(
                "TIKTOK VIDEOS",
                links["TikTok"],
                use_container_width=True,
            )

        st.header("🌐 SOCIAL MEDIA")

        social_columns = st.columns(3)

        with social_columns[0]:
            st.link_button(
                "FACEBOOK",
                links["Facebook"],
                use_container_width=True,
            )

        with social_columns[1]:
            st.link_button(
                "INSTAGRAM",
                links["Instagram"],
                use_container_width=True,
            )

        with social_columns[2]:
            st.link_button(
                "TIKTOK",
                links["TikTok"],
                use_container_width=True,
            )

        st.header("📍 ΣΤΟΙΧΕΙΑ ΚΑΤΑΣΤΗΜΑΤΟΣ")

        st.link_button(
            "GOOGLE MAPS",
            links["Google Maps"],
            use_container_width=True,
        )

        st.link_button(
            "ΑΝΑΖΗΤΗΣΗ EMAIL",
            links["Email"],
            use_container_width=True,
        )

        st.header("🎵 ΕΛΕΓΧΟΣ ΜΟΥΣΙΚΗΣ ΚΑΙ DJ")

        st.link_button(
            "GOOGLE REVIEWS – ΜΟΥΣΙΚΗ / DJ",
            links["Google Reviews"],
            use_container_width=True,
        )

        st.link_button(
            "TRIPADVISOR",
            links["Tripadvisor"],
            use_container_width=True,
        )

        results = pd.DataFrame(
            [
                {
                    "Κατάστημα": store_name,
                    "Πόλη": city,
                    "Πηγή": source,
                    "Link": url,
                }
                for source, url in links.items()
            ]
        )

        csv_data = results.to_csv(
            index=False,
        ).encode("utf-8-sig")

        st.header("📥 ΕΞΑΓΩΓΗ")

        st.download_button(
            label="ΛΗΨΗ ΑΠΟΤΕΛΕΣΜΑΤΩΝ CSV",
            data=csv_data,
            file_name=(
                f"{store_name}_{city}_έλεγχος.csv"
            ),
            mime="text/csv",
            use_container_width=True,
        )

        with st.expander("ΠΡΟΒΟΛΗ ΟΛΩΝ ΤΩΝ LINKS"):
            st.dataframe(
                results,
                use_container_width=True,
                hide_index=True,
            )
