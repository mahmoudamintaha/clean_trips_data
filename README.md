# Project Description

This Python script is designed to create two Excel files named `offer_contents` and `offer_templates`.

## Script Overview

1. The script reads an Excel workbook and adds all its sheets into one DataFrame.
2. Country names are extracted from sheet names and validated against the backend data stored in the input file.
3. Then, country names are corrected using an online file containing information about countries and capitals.
4. Using regex, the departure port and arrival port are extracted and validated against port data stored in the backend file stored in the input folder.
5. Any misspelling in the port names is corrected using the `fuzzywuzzy` module.
6. Price and trip type are extracted using regex.
7. A function is created to translate numbers into Arabic script, and a dictionary with languages and price scripts is created.
8. The code will fetch any wrong (country or ports) mentioned in the trips data and exclude them from the final result.
9. Finally, a report with wrong entries will be generated, including misspelled entries. The report contains an `error_type` column that states whether it's entirely wrong or misspelled.

### Objective
The primary objective of this script is to meticulously clean and validate data, ensuring utmost accuracy, while also providing users with informative insights regarding any erroneous entries.