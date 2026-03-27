r"""
--------------------\ \ RFP Report Creator / /--------------------
Project Inputs:
    - UCB Requests for Proposals (RFPs).csv
    - Standford Link: "https://cfr.stanford.edu/faculty/funding-opportunities-resources/rfps"
    - Princeton Link: r"https://princeton.infoready4.com/#engagementHubResults?fields=Funding+Source%7COR%7CFoundation%5EIndustry"
    - output_folder:  r"C:\Users\ballen47.ASURITE\Documents\OFFICE AIDE WORK\RFP LISTS"
Project Outputs:
    - A formatted RFP source list in excel with the following formatting information:
        - Deadline
        - Granting Organization
        - Project Name (linked)
        - Funding Amount
        - Keywords
        - Disciplines (NS, SS, H)
        - raw URL
        - raw Grant Name
Project Methodology:
    - Utilizes an open source machine learning algoithm downloaded and run on device to find keywords from given sample
    - Scrapes urls to find relevant data
    - Uses a Python library to format it into a nice looking excel spreadsheet
--------------------/ / RFP Report Creator \ \--------------------
"""

r"""--------------------\ \ IMPORT LIBRARIES / /--------------------"""
import RFP_Functions as rfp

r"""--------------------/ / IMPORT LIBRARIES \ \--------------------"""


r"""--------------------\ \ GET DATA FILES / /--------------------"""

# UCB
UCB_rfp_library = rfp.UCB_csv_reader("UCB Requests for Proposals (RFPs).csv")

# Stanford
Stanford_rfp_library = rfp.web_scraper_stanford(
    "https://cfr.stanford.edu/faculty/funding-opportunities-resources/rfps"
)

# Princeton
Princeton_rfp_library = rfp.web_scraper_princeton(
    r"https://princeton.infoready4.com/#engagementHubResults?fields=Funding+Source%7COR%7CFoundation%5EIndustry"
)

# Get all of the libraries gathered into one library for the final report
combined_raw_library = UCB_rfp_library + Stanford_rfp_library + Princeton_rfp_library

# This function adds displine data based on included keywords and simple (but non robust) rules
complete_rfp_library = rfp.add_discipline_data_from_keywords(combined_raw_library)

# This function makes the dates all the same format and checks for (and deletes) duplicate recources
clean_library = rfp.clean_and_deduplicate(combined_raw_library)

r"""--------------------/ / GET DATA FILES \ \--------------------"""


r"""--------------------\ \ IMPORT LIBRARIES / /--------------------"""
output_folder = r"C:\Users\ballen47.ASURITE\Documents\OFFICE AIDE WORK\RFP LISTS"
rfp.report_maker(complete_rfp_library, output_folder)
r"""--------------------/ / IMPORT LIBRARIES \ \--------------------"""
