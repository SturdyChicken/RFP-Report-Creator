"""
--------------------\ \ RFP Report Creator / /--------------------
What the user does:
    - Upload the UCB CSV file.
    - Check which additional sources to include (Stanford / Princeton).
    - Click "Generate Report".
    - Wait ~5 minutes while the app scrapes & runs AI keyword analysis.
    - Click the download button to save the finished .xlsx file.
--------------------/ / RFP Report Creator \ \--------------------
"""

## -------------------- IMPORT LIBRARIES -------------------- ##
import os
import glob
import tempfile
import streamlit as st
import RFP_Functions as rfp

# Install Playwright Chromium browser binaries on start-up
os.system("playwright install chromium")
import nest_asyncio

nest_asyncio.apply()
## -------------------- IMPORT LIBRARIES -------------------- ##


## -------------------- HEADER -------------------- ##
st.set_page_config(
    page_title="RFP Report Creator",
    layout="centered",
)
st.title("RFP Report Creator")
st.markdown(
    "Upload the **UCB Requests for Proposals CSV**, choose your additional "
    "sources, and generate a formatted Excel report."
)
st.divider()

## -------------------- HEADER -------------------- ##


## -------------------- INPUTS -------------------- ##
st.subheader("Upload UCB Data File")
st.markdown(
    "Navigate to the following URL: https://airtable.com/appMeDyM5cOiMZOqT/shrA7IHmFjjWfB918/tblJT0jbbJHBSnbTI\n\n"
    "Then click **'...'** → **'Download CSV'**"
)
ucb_file = st.file_uploader(
    label="UCB Requests for Proposals (.csv)",
    type=["csv"],
    help="This is the CSV file exported from the UCB RFP database.",
)

st.subheader("Additional Sources")
col1, col2 = st.columns(2)
with col1:
    include_stanford = st.toggle("Stanford", value=True)
with col2:
    include_princeton = st.toggle("Princeton", value=True)

st.divider()
## -------------------- INPUTS -------------------- ##


## -------------------- BUTTON -------------------- ##
st.subheader("Generate Report")

if include_princeton:
    st.info(
        "Princeton requires a headless browser to load its data. "
        "Expect the full report to take **~4–6 minutes** to generate.",
        icon="ℹ️",
    )
elif include_stanford:
    st.info("Estimated run time: **2–4 minutes**.", icon="ℹ️")
else:
    st.info("Estimated run time: **1–2 minutes** (UCB only).", icon="ℹ️")

generate_clicked = st.button(
    "Generate Report",
    type="primary",
    disabled=(ucb_file is None),  # Greyed-out until a file is uploaded
    use_container_width=True,
)

if ucb_file is None:
    st.caption("Upload the UCB CSV above to enable this button.")
## -------------------- BUTTON -------------------- ##


## -------------------- GENERATOR -------------------- ##
if generate_clicked:

    # We use two temp directories:
    #   • input_dir  - holds the uploaded UCB csv so rfp.UCB_csv_reader()
    #                  can find it by file path (it uses os.path.exists internally)
    #   • output_dir - receives the finished .xlsx from rfp.report_maker()
    with tempfile.TemporaryDirectory() as input_dir, tempfile.TemporaryDirectory() as output_dir:

        ucb_temp_path = os.path.join(input_dir, ucb_file.name)
        with open(ucb_temp_path, "wb") as f:
            f.write(ucb_file.getbuffer())

        # live status updates
        combined_library = []
        error_occurred = False

        with st.status(
            "Running — please don't close this tab…", expanded=True
        ) as status:

            # UCB
            st.write("Reading UCB CSV…")
            try:
                ucb_library = rfp.UCB_csv_reader(ucb_temp_path)
                st.write(f"UCB: {len(ucb_library)} grants loaded.")
                combined_library += ucb_library
            except Exception as e:
                st.error(f"UCB processing failed: {e}")
                error_occurred = True

            # STANFORD
            if include_stanford and not error_occurred:
                st.write("Scraping Stanford…  please be patient")
                try:
                    stanford_library = rfp.web_scraper_stanford(
                        "https://cfr.stanford.edu/faculty/funding-opportunities-resources/rfps"
                    )
                    st.write(f"✅ Stanford: {len(stanford_library)} grants loaded.")
                    combined_library += stanford_library
                except Exception as e:
                    st.warning(f"Stanford scrape failed (skipping): {e}")

            # PRINCETON
            if include_princeton and not error_occurred:
                st.write("Scraping Princeton…  " "*(launching headless browser)*")
                try:
                    princeton_library = rfp.web_scraper_princeton(
                        r"https://princeton.infoready4.com/#engagementHubResults"
                        r"?fields=Funding+Source%7COR%7CFoundation%5EIndustry"
                    )
                    st.write(f"Princeton: {len(princeton_library)} grants loaded.")
                    combined_library += princeton_library
                except Exception as e:
                    st.warning(f"Princeton scrape failed (skipping): {e}")

            # CLEAN
            if combined_library and not error_occurred:
                st.write("Running AI keyword analysis…  this may take a while")
                try:
                    combined_library = rfp.add_discipline_data_from_keywords(
                        combined_library
                    )
                    st.write("Keywords assigned.")
                except Exception as e:
                    st.error(f"Keyword analysis failed: {e}")
                    error_occurred = True

            if combined_library and not error_occurred:
                st.write("Cleaning & deduplicating…")
                try:
                    combined_library = rfp.clean_and_deduplicate(combined_library)
                    st.write(f"{len(combined_library)} unique grants after dedup.")
                except Exception as e:
                    st.warning(f"Deduplication failed (continuing anyway): {e}")

            # GENERATE REPORT
            if combined_library and not error_occurred:
                st.write("Building Excel report…")
                try:
                    rfp.report_maker(combined_library, output_dir)
                    st.write("Excel file created.")
                except Exception as e:
                    st.error(f"Report generation failed: {e}")
                    error_occurred = True

            if error_occurred:
                status.update(label="Report generation failed.", state="error")
            elif not combined_library:
                status.update(
                    label="No grants found — nothing to report.", state="error"
                )
            else:
                status.update(label="Report ready!", state="complete", expanded=False)

        # report_maker names the file after the current month, so we glob for it
        if not error_occurred and combined_library:
            xlsx_files = glob.glob(os.path.join(output_dir, "*.xlsx"))

            if xlsx_files:
                output_xlsx_path = xlsx_files[0]
                output_filename = os.path.basename(output_xlsx_path)

                with open(output_xlsx_path, "rb") as f:
                    xlsx_bytes = f.read()

                st.divider()
                st.success(
                    f"Your report is ready! "
                    f"**{len(combined_library)} grants** included across all sources."
                )
                st.download_button(
                    label="⬇Download Excel Report",
                    data=xlsx_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )
            else:
                st.error("Report file not found after generation. Please try again.")
## -------------------- GENERATOR -------------------- ##
