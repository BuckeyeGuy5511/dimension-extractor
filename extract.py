import pandas as pd
import spacy
import re
from dateutil.parser import parse
import os
import glob

nlp = spacy.load("en_core_web_md")

recency_concepts = ["recent", "last", "latest", "newest", "most recent"]


def calculate_recency_score(cell):
    doc = nlp(cell)
    score = 0

    dates = []
    for ent in doc.ents:
        if ent.label_ == "DATE":
            try:
                parsed_date = parse(ent.text, fuzzy=True)
                dates.append(parsed_date)
            except ValueError:
                continue

    if dates:
        print(dates)
        most_recent_date = max(dates)
        score += 1

    for token in doc:
        if token.dep_ in ["amod", "advmod", "nummod"] and token.head.text in cell:
            score += 1

    for concept in recency_concepts:
        concept_doc = nlp(concept)
        similarity = doc.similarity(concept_doc)
        score += similarity
    print("Score: ", score, "Cell: ", cell)
    return score


def extract_most_recent_dimensions(file_path):
    df = pd.read_excel(file_path, sheet_name=None)

    most_recent_dimensions = None
    most_recent_score = 0
    dimensions_list = []

    for sheet_name, sheet_data in df.items():
        for row in sheet_data.iterrows():
            for cell in row[1]:
                if isinstance(cell, str):
                    doc = nlp(cell)

                    dimensions_match = re.search(
                        r'\b(\d*\.?\d+)\s*x\s*(\d*\.?\d+)\s*x\s*(\d*\.?\d+)\b', cell)

                    if dimensions_match:
                        dimensions = dimensions_match.group(0)
                        dimensions_list.append(dimensions)

                        score = calculate_recency_score(cell)

                        if score > most_recent_score:
                            most_recent_dimensions = dimensions
                            most_recent_score = score

    return most_recent_dimensions, dimensions_list


def process_all_files_in_documents():
    documents_dir = 'documents'
    results_dir = 'results'
    results_file = os.path.join(results_dir, 'results.txt')

    os.makedirs(results_dir, exist_ok=True)

    excel_files = glob.glob(os.path.join(documents_dir, '*.xlsx'))

    all_results = []

    for file_path in excel_files:
        most_recent_dimensions, dimensions_list = extract_most_recent_dimensions(
            file_path)
        result = f"File: {os.path.basename(file_path)}\n"
        result += f"The most recent dimensions of the gland are: {most_recent_dimensions}\n"
        result += f"All matched dimensions: {dimensions_list}\n\n"

    with open(results_file, 'w', encoding='utf-8') as file:
        file.writelines(all_results)

    print(f"Results written to {results_file}")


process_all_files_in_documents()
