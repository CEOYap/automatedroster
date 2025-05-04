# Automated Roster Processor

## Description
This Python script automatically extract (Number, Rank, Name) from multiple Word docs and categorises them into the appropriate category.

It is designed to handle variations in document formatting, including different list styles (numbered, bulleted, symbols like '👉'), minor header variations, and common invisible characters sometimes found in copied text.

Finally, it generates into 2 Excel sheet:
1.  `Output_1.xlsx`: Contains all extracted unique names with their assigned category, unit category (derived from the filename), and a sequential number (`BIL`) across the entire dataset.
2.  `Output_sorted.xlsx`: Contains the same personnel data, but sorted alphabetically by category, with a sequential count (`COUNT`) restarting for each category.
## Key Features
* Processes multiple docx - Simply add them to `input_docs` folder .
* Fast and accurate - Uses two method `startwith`  and check for keywords
* Easily configurable - Add more categories under `SECTION_MAP`
* Cleans input text with multiple methods to ensure output are easily readable
* Assigns unique names to "NEW CHECK" category for manual review if unknown category
## Requirements
* Python 3.x
* [`pandas` library](https://pandas.pydata.org/)
* [`python-docx` library](https://python-docx.readthedocs.io)
## Usage
1.  Download repo as a zip file and extract
3. **Install Libraries:**
    ```bash
    pip install pandas python-docx
    ```
2.  Place all the `.docx` files you want to process inside the `input_docs` folder. Script assumes filename (without extension). Example `BLACKOPS.docx` -> UNIT 'BLACKOPS'
3. **Run the Script:**
    ```bash
    python arr.py
    ```
4.  **Check Output:**
    * Two Excel files will be created in the same directory as the script:
        * `Output_1.xlsx`
        * `Output_sorted.xlsx`
    * Review any entries assigned to the `NEW CHECK` category to see if new headers need to be added to the `SECTION_MAP`.

5.  **Customization (Optional):**
    * Modify the `INPUT_FOLDER`, `OUTPUT_FILE`, `OUTPUT_SORTED_FILE`, and `NEW_CATEGORY_LABEL` constants if needed.
    * Edit the `SECTION_MAP` dictionary within the script to add, remove, or modify header keywords and their corresponding category mappings. Remember that the order matters for the matching logic – place more specific keywords before more general ones.
## Limitations
* **Multi-line Entries:** The script processes documents paragraph by paragraph. Personnel names or other data split across multiple lines might not be parsed correctly.
* **Complex Layouts:** Does not handle data within tables, text boxes, or complex document structures other than standard paragraphs.
* **Header Ambiguity:** Relies heavily on the keywords and order defined in `SECTION_MAP`. Ambiguous headers or very short keywords might occasionally lead to miscategorisation, although the matching logic attempts to mitigate this. Entries under `NEW CHECK` require manual verification.
## License
Released under the [MIT License](https://opensource.org/licenses/MIT).