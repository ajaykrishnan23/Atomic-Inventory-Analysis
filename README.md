# How I approached the problem

### Friday:
    - Spent the first day going over the data to put myself in the shoes of a sales engineer
    - Understood patterns and common trends 
    - Tried to get a sense of what information I should extract to save time
    - Formed the problem statement

### Saturday:
    - Built a primitive huerestic approach to try reducing the reliance on the LLM **Refer heurestic_final.py or last 3 cells from the NewApproach.ipynb notebook**
    - Said script essentially:
        - Scrape tables from excel using empty rows as delimiters
        - Does formula reconciliation
        - Does column extraction
        - Performs Table Merging
    - Disadvantages with this script:
        - Can't distinguish between horizontal and vertical tables
        - Very heurestics based, might not scale on the longer run for different formats
        - Tables extracted lacked contextual awareness of columns and such to end to the LLM
        - LLM gave out poor results due to this while testing by making chatgpt act as a sales engineer and extract information
        

### Sunday
    - Since I understood the issues and forthcomings of a heurestic based approach, I switched to a fully agentic approach **Refer last 2 cells of pass_thru_LLM.ipynb notebook**
    - Built an agent that generates custom code for each sheet given a portion of the sheet
    - Added retries for the agent to handle errors (can add feedback mechanism for it to learn from the errors)
    - Logged all statements being generated for observability (can be pushed to a better logstore like logfire in the future)
    - Wrote a similar table analysis script too.




# Table Extraction

## **Code Logic**

1. **Sheet Data Chunk Extraction**:
   - For each sheet in the workbook, the first 50 rows (50xN chunk) are extracted to analyze and identify potential data tables.
   - This is done to ensure the script handles large sheets efficiently by focusing on the most relevant portion of the data.
   - An issue with this approach is that all tables might not get captured / sometimes the data we send to the LLM can be more than its context length so this would require some additional chunking while making sure that the data is clean

2. **LLM-Powered Code Generation**:
   - The extracted chunk is sent to a language model (`gpt-4o-mini`) via LangChain with a well-defined prompt.
   - The prompt instructs the LLM to:
     - Identify tables within the chunk.
     - Generate Python code to extract the tables as separate CSV files and save them in the designated folder.
   - This code is generated dynamically for each page encountered to account for variability in formats. 
   - The folder for saving CSV files is explicitly passed to ensure the output is correctly placed in the intended location.

3. **Generated Code Execution**:
   - The generated Python code is executed.
   - The results are logged.

4. **Retry Mechanism**:
   - If an error occurs during the execution of the generated code, the script retries up to 10 times.



### **Advantages**

1. **LLM-Driven Table Extraction**:
   - Leverages the power of GPT-4o-mini to dynamically analyze sheet data and generate code tailored to the specific structure of each sheet.

2. **Scalability**:
   - Handles workbooks with multiple sheets seamlessly.
   - Processes each sheet independently, making it suitable for large datasets.

3. **Error Resilience**:
   - Incorporates a retry mechanism to handle and recover from errors during code generation and execution.

4. **Customizable Workflow**:
   - The script is modular and built natively, allowing for easy customization of prompts, directory structures, and retry logic.

5. **Clear Logging**:
   - Logs each step of the process, from data extraction to code execution, providing transparency and aiding in debugging.



# Table Analysis Script

This script performs automated analysis of CSV files in multiple folders, determines if they are related to inventory planning, and extracts relevant details about their structure. It leverages **Pydantic** for validation and **GPT-4o-mini** for reasoning and analysis. The results are saved in a consolidated CSV for further review.

---

## **Workflow**

1. **Structured Output Schema Definition**:
   - Uses Pydantic to define a strict schema for the analysis results.
   - Ensures outputs conform to a consistent format with the following fields:
     - `is_inventory_planning`: Indicates whether the file is related to inventory planning (1 for Yes, 0 for No).
     - `details`: A dictionary mapping column names to their roles (e.g., SKU, Location/Warehouse, etc.).
     - `description`: Explains why the file is or isn’t identified as inventory planning.

2. **Sheet Analysis**:
   - The function `analyze_sheet` analyzes the first 5 rows of a CSV file converted into Markdown format.
   - GPT-4o-mini is prompted with these rows and tasked to:
     - Classify the file as inventory planning or not.
     - Map relevant column names to their roles.
     - Provide a description of the decision-making process.
   - Includes a retry mechanism to handle validation errors or failures, retrying up to 10 times if necessary.

3. **CSV File Processing**:
   - The function `analyze_csv_files` iterates through all folders and files in a base directory.
   - For each CSV file, it:
     - Loads the data into a DataFrame.
     - Passes the data to `analyze_sheet` for analysis.
     - Collects results into a DataFrame, unpacking the `details` dictionary into separate columns for easier readability.

## **Advantages**

1. **Automated Analysis**:
   - The script automates the classification of CSV files for inventory planning tasks using GPT-4o-mini.
   - Eliminates manual effort in identifying and mapping inventory-related columns.

2. **Error Handling and Retries**:
   - Includes a retry mechanism that retries up to 10 times for validation errors.
   - Ensures robustness in processing and parsing LLM responses.

3. **Structured Output**:
   - Uses Pydantic for schema validation to enforce consistent and reliable results.
   - Ensures outputs conform to predefined fields like `is_inventory_planning`, `details`, and `description`.

4. **Consolidated Results**:
   - Generates a final consolidated CSV containing analysis results for all processed files.
   - Unpacks details into separate columns for easier readability and usability.

5. **Transparency**:
   - Logs every processing step, including validation errors and retries.
   - Enables clear tracking of progress and debugging, if necessary.



# Setup and Usage

## Installation
Make sure the environment variable OPENAI_API_KEY has been set.
```bash
# Clone the repository and navigate to directory
git clone 
cd <repository-name>

# Create conda environment from file
conda env create -f environment.yml

# Activate the environment
conda activate atomic
```

## Running the Scripts

```bash
# Extract tables from documents
python table_extraction.py

# Analyze extracted tables
python table_analysis.py
```
The notebooks are present for me experimenting with different approaches only, don't bother reading through them. 

# Two Cents on What's Next
    - Table extraction works intuitively
    - But the LLM Has not yet "Understood" the data to properly pick out the business logic I feel
        - I can get this done if I have a clear understanding of what the Sales Engineers look for by business logic 
    - Better logging (Logging all events + Cost)
