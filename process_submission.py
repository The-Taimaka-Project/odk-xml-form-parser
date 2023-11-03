import re
import xml.etree.ElementTree as ET
import pandas as pd
import argparse
import logging

logging.basicConfig(filename='scripts/process_submission.log', level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s', filemode='w')

def get_choice_label(list_name, choice_name, choices_data):
    choice_row = choices_data[(choices_data['list_name'] == list_name) & (choices_data['name'] == choice_name)]
    if not choice_row.empty:
        return choice_row['label'].iloc[0]
    return None

def get_multiple_choice_labels(list_name, choices_str, choices_data):
    choice_names = choices_str.split() if choices_str else []
    labels = [get_choice_label(list_name, choice_name, choices_data) for choice_name in choice_names]
    return ", ".join([label for label in labels if label])

def replace_references_with_answers_nested(label, root):
    if not isinstance(label, str):
        return label
    references = re.findall(r"\$\{(.*?)\}", label)
    for ref_name in references:
        tag = root.find(".//{}".format(ref_name))
        if tag is not None and tag.text is not None:
            label = label.replace("${" + ref_name + "}", tag.text)
    return label

def main(xml_path, xlsx_path, output_html_path):
    logging.info(f"Started processing with XML: {xml_path}, XLSX: {xlsx_path}, and output to: {output_html_path}")

    # Load XML data
    tree = ET.parse(xml_path)
    root = tree.getroot()

    # Load form definition from Excel
    form_data = pd.read_excel(xlsx_path, sheet_name="survey")
    choices_data = pd.read_excel(xlsx_path, sheet_name="choices")
    form_dict = dict(zip(form_data['name'], form_data['label']))

    # Generate the HTML content
    html_content = "<html><head><title>Submission Data</title></head><body>"

    for tag in root.findall(".//*"):
        if len(tag) == 0 and tag.tag in form_dict:
            label = form_dict[tag.tag]
            label = replace_references_with_answers_nested(label, root)
            question_type = form_data[form_data['name'] == tag.tag]['type'].iloc[0] if tag.tag in form_data['name'].values else None
            question_appearance = form_data[form_data['name'] == tag.tag]['appearance'].iloc[0] if tag.tag in form_data['name'].values else None

            if question_type and question_type.startswith("select_multiple") and not question_type.startswith("select_multiple_from_file"):
                list_name = question_type.split()[-1]
                choice_labels = get_multiple_choice_labels(list_name, tag.text, choices_data)
                if choice_labels:
                    tag.text = choice_labels
            
            if question_type and question_type.startswith("select_one") and not question_type.startswith("select_one_from_file"):
                list_name = question_type.split()[-1]
                choice_label = get_choice_label(list_name, tag.text, choices_data)
                if choice_label:
                    tag.text = choice_label

            if question_type == "note" or question_appearance == "label":
                html_content += f"<p><strong>{label}</strong></p>"
            elif pd.notna(label):
                if label.endswith(":"):
                    html_content += f"<p><strong>{label}</strong> {tag.text}</p>"
                else:
                    html_content += f"<p><strong>{label}:</strong> {tag.text}</p>"

    html_content += "</body></html>"

    # Write the generated HTML content to an output file
    with open(output_html_path, "w") as file:
        file.write(html_content)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Process ODK submission.")
    parser.add_argument('xml_path', help='Path to the XML submission.')
    parser.add_argument('xlsx_path', help='Path to the Excel form definition.')
    parser.add_argument('output_html_path', help='Path for the output HTML file.')

    args = parser.parse_args()

    main(args.xml_path, args.xlsx_path, args.output_html_path)