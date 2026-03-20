## Program Description and Goals: All in Python
# Create a formatted Excel Spreadsheet of an RFP List or ASU with the following features:
#   The following columns of information will be included:
#       1. Deadline
#       2. Funding Organization Name
#       3. <=HYPERLINK(I#,J#)>
#       4. Funding Amount
#       5. Key Words (Description / Abstract?)
#       6. Natural Science?
#       7. Social Science?
#       8. Humanities?
#       9. Direct URL
#       10. Grant Name
#   The following RFP Sources will be supported:
#       1. UCB Airtable (already formatted well in excel spreadsheet file)
#       2. Stanford (already formatted well in text exported file)
#       3. Princeton (for now only get deadline, fund name, hyperlink, Description)
#       4. Baylor (for now only get fund name, hyperlink, Description)
#       5. Colorado (more info soon)
#       6. RWJF (more info soon)
#       7. Philanthropy News Digest (more info soon)
#       8. Call for additional recource lists, I can add them
#   Double checking for duplicate funding opprotunities before final export
#   The following additional features should be implemented eventually:
#       1. Properly formatted and exported file downloaded (CSV)
#       2. Ease of Use features, like a downloadable executable file with dialogue boxes for team members to use
#           a. Maybe some customizable features, like filtering by Division
#       3. 


## ------------ IMPORT FUNCTIONS ------------
import RFP_Functions as rfp

## ------------ GET DATA ------------

# UCB
UCB_rfp_library = rfp.UCB_csv_reader("UCB Requests for Proposals (RFPs).csv")

# Stanford
Stanford_rfp_library = rfp.web_scraper_stanford("https://cfr.stanford.edu/faculty/funding-opportunities-resources/rfps")

# Princeton
Princeton_rfp_library = rfp.web_scraper_princeton(r"https://princeton.infoready4.com/#engagementHubResults?fields=Funding+Source%7COR%7CFoundation%5EIndustry")

# Get all of the libraries gathered into one library for the final report
combined_raw_library = UCB_rfp_library + Princeton_rfp_library

# This function adds displine data based on included keywords and simple (but non robust) rules
complete_rfp_library = rfp.add_discipline_data_from_keywords(combined_raw_library)

# This function makes the dates all the same format and checks for (and deletes) duplicate recources
clean_library = rfp.clean_and_deduplicate(combined_raw_library)

## ------------ CREATE REPORT ------------

output_folder = r"C:\Users\ballen47.ASURITE\Documents\OFFICE AIDE WORK\RFP LISTS"
rfp.report_maker(complete_rfp_library, output_folder)

