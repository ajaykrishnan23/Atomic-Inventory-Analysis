import os
import pandas as pd
from pydantic import BaseModel, Field, ValidationError
from langchain.chat_models import ChatOpenAI
import json


# Define the structured output schema
class InventorySheetAnalysis(BaseModel):
    is_inventory_planning: int = Field(
        ..., description="1 if the sheet is for inventory planning, 0 otherwise"
    )
    details: dict = Field(
        ...,
        description=(
            "A dictionary mapping columns to roles: SKU, Location/Warehouse, Quantity, "
            "Total Inventory, Current Inventory, Sales Forecast. Set to None if not found."
        ),
    )
    description: str = Field(
        ..., description="Explanation of why this sheet is for inventory planning."
    )


# Function to clean the LLM response
def clean_response(response: str) -> str:
    if response.startswith("```json"):
        response = response[7:]  # Remove ```json
    if response.endswith("```"):
        response = response[:-3]  # Remove ```
    return response.strip()  # Strip any extra whitespace


# Function to analyze a sheet using GPT-4o-mini
def analyze_sheet(dataframe, llm, retries=10):
    markdown_preview = dataframe.head(5).to_markdown()

    prompt = f"""
    You are a data analysis expert. Below is a preview of a sheet:

    {markdown_preview}

    Task:
    1. Determine if this sheet is used for inventory planning (1 for Yes, 0 for No).
    2. If Yes, provide:
       - 'details': A dictionary mapping column names to these roles: SKU, Location/Warehouse, Quantity, Total Inventory, Current Inventory, Sales Forecast. If no related column exists, set the value to None.
       - 'description': A brief explanation explaining why this sheet is for inventory planning.

    Respond strictly in the following JSON format:
    {{
        "is_inventory_planning": 1,
        "details": {{
            "SKU": "Column Name or None",
            "Location/Warehouse": "Column Name or None",
            "Quantity": "Column Name or None",
            "Total Inventory": "Column Name or None",
            "Current Inventory": "Column Name or None",
            "Sales Forecast": "Column Name or None"
        }},
        "description": "Brief explanation"
    }}
    """

    while retries > 0:
        try:
            response = llm.predict(prompt)
            cleaned_response = clean_response(response)
            structured_output = InventorySheetAnalysis.model_validate_json(cleaned_response)
            return structured_output
        except ValidationError as ve:
            print(f"Validation Error: {ve}")
            retries -= 1
            print(f"Retries left: {retries}")
            if retries == 0:
                print("Max retries reached. Skipping this sheet.")
                return None
        except Exception as e:
            print(f"Unexpected Error: {e}")
            return None


# Function to process all CSV files in a base folder
def analyze_csv_files(base_folder, llm):
    results = []

    for folder in os.listdir(base_folder):
        folder_path = os.path.join(base_folder, folder)
        if os.path.isdir(folder_path):
            for file in os.listdir(folder_path):
                if file.endswith(".csv"):
                    file_path = os.path.join(folder_path, file)
                    print(f"Processing file: {file_path}")

                    df = pd.read_csv(file_path)
                    response = analyze_sheet(dataframe=df, llm=llm)
                    if response:
                        # Unpack the details dictionary into individual columns
                        details = response.details
                        results.append(
                            {
                                'sheet_name': folder,
                                "file": file,
                                "is_inventory_planning": response.is_inventory_planning,
                                "description": response.description,
                                "SKU": details.get("SKU"),
                                "Location/Warehouse": details.get("Location/Warehouse"),
                                "Quantity": details.get("Quantity"),
                                "Total Inventory": details.get("Total Inventory"),
                                "Current Inventory": details.get("Current Inventory"),
                                "Sales Forecast": details.get("Sales Forecast"),
                            }
                        )
                    else:
                        print(f"Failed to analyze file: {file}")

    return pd.DataFrame(results)


# Main Execution
llm = ChatOpenAI(model="gpt-4o-mini", temperature=0)
# provide paths to the folders containing the csvs
base_folders = ["/Users/ajay/Documents/Atomic/inventory_analysis_2/Company 1 - Inventory Planning", "/Users/ajay/Documents/Atomic/inventory_analysis_2/Company 2 - Supply Management", "/Users/ajay/Documents/Atomic/inventory_analysis_2/Company 3 - Inventory Dashboard _V2"]
for base_folder in base_folders:
    analysis_results_df = analyze_csv_files(base_folder, llm)

    # Save results to a CSV file
    output_file = f"{base_folder.split('/')[-1]}.csv"
    analysis_results_df.to_csv(output_file, index=False)
    print(f"Results saved to {output_file}")


