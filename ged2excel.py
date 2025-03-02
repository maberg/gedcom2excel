import argparse
import pandas as pd
from gedcom.element.individual import IndividualElement
from gedcom.element.family import FamilyElement
from gedcom.element.element import Element
from gedcom.parser import Parser
import tempfile
import os
import re

def repair_gedcom_numbering(input_file_path, temp_file_path):
    """Repairs incorrect line numbering in a GEDCOM file."""
    with open(input_file_path, 'r', encoding='utf-8') as infile, \
         open(temp_file_path, 'w', encoding='utf-8') as outfile:
        current_level = 0
        previous_level = -1
        for line in infile:
            line = line.strip()
            if not line:
                continue
            parts = line.split(' ', 2)
            try:
                level = int(parts[0])
            except ValueError:
                level = previous_level + 1 if previous_level >= 0 else 0
            if level > previous_level + 1:
                level = previous_level + 1
            elif level < 0:
                level = 0
            if len(parts) == 3:
                repaired_line = f"{level} {parts[1]} {parts[2]}"
            else:
                repaired_line = f"{level} {parts[1]}"
            outfile.write(repaired_line + '\n')
            previous_level = level
    print(f"Repaired GEDCOM saved to temporary file: {temp_file_path}")

def sanitize_string(text):
    """Remove or replace invalid XML characters that Excel cannot handle."""
    if not isinstance(text, str):
        return text
    # Remove control characters (0x00-0x1F, except 0x09, 0x0A, 0x0D: tab, newline, carriage return)
    text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', text)
    # Replace problematic characters with alternatives or remove them
    text = text.replace('^', ' ')  # Replace ^ with space (could be a separator)
    text = text.replace('·', '.')  # Replace middle dot with period
    # Keep Ç and other valid Unicode characters unless they cause issues
    return text

def gedcom_to_excel(gedcom_file_path, excel_file_path):
    with tempfile.NamedTemporaryFile(mode='w+', delete=False, suffix='.ged') as temp_file:
        repair_gedcom_numbering(gedcom_file_path, temp_file.name)
        repaired_gedcom_path = temp_file.name
    
    try:
        gedcom_parser = Parser()
        gedcom_parser.parse_file(repaired_gedcom_path)
        
        individuals_data = {
            'ID': [], 'Name': [], 'Gender': [], 'Birth Date': [], 'Birth Place': [],
            'Death Date': [], 'Death Place': [], 'Father ID': [], 'Mother ID': [],
            'Father Name': [], 'Mother Name': [], 'Spouse ID': [], 'Spouse Name': [],
            'Occupation': [], 'Education': [], 'Religion': [], 'Nationality': [],
            'Physical Description': [], 'SSN': [], 'Title': [], 'Cause of Death': [],
            'Residence': [], 'Children IDs': [], 'Children Names': [], 'Notes': [],
            'Change Date': [], 'Change Time': []
        }
        
        families_data = {
            'Family ID': [], 'Husband ID': [], 'Husband Name': [], 'Wife ID': [], 'Wife Name': [],
            'Marriage Date': [], 'Marriage Place': [], 'Divorce Date': [], 'Divorce Place': [],
            'Engagement Date': [], 'Engagement Place': [], 'Marriage Contract Date': [],
            'Marriage Contract Place': [], 'Marriage Settlement Date': [], 'Marriage Settlement Place': [],
            'Children IDs': [], 'Children Names': [], 'Notes': [], 'Change Date': [], 'Change Time': []
        }
        
        events_data = {
            'Record ID': [], 'Record Type': [], 'Event Type': [], 'Date': [], 'Place': [],
            'Cause': [], 'Notes': [], 'Source IDs': []
        }
        
        sources_data = {
            'Source ID': [], 'Title': [], 'Author': [], 'Publication': [], 'Page': [],
            'Repository': [], 'Data': [], 'Notes': []
        }
        
        notes_data = {
            'Note ID': [], 'Text': [], 'Referenced By': []
        }
        
        multimedia_data = {
            'Object ID': [], 'File': [], 'Format': [], 'Title': [], 'Notes': []
        }
        
        associations_data = {
            'Individual ID': [], 'Associated ID': [], 'Relationship': [], 'Notes': []
        }
        
        submitter_data = {
            'Submitter ID': [], 'Name': [], 'Address': [], 'Phone': [], 'Email': []
        }
        
        header_data = {
            'Source Software': [], 'Source Version': [], 'Date': [], 'Character Set': [],
            'GEDCOM Version': []
        }
        
        id_to_name = {}
        child_to_parents = {}
        id_to_spouses = {}
        id_to_children = {}
        
        # Header
        header = gedcom_parser.get_root_element().get_child_elements()
        for elem in header:
            if elem.get_tag() == "HEAD":
                source = version = date = charset = gedc_version = ""
                for child in elem.get_child_elements():
                    if child.get_tag() == "SOUR":
                        source = child.get_value()
                        for sub in child.get_child_elements():
                            if sub.get_tag() == "VERS": version = sub.get_value()
                    elif child.get_tag() == "DATE": date = child.get_value()
                    elif child.get_tag() == "CHAR": charset = child.get_value()
                    elif child.get_tag() == "GEDC":
                        for sub in child.get_child_elements():
                            if sub.get_tag() == "VERS": gedc_version = sub.get_value()
                header_data['Source Software'].append(sanitize_string(source))
                header_data['Source Version'].append(sanitize_string(version))
                header_data['Date'].append(sanitize_string(date))
                header_data['Character Set'].append(sanitize_string(charset))
                header_data['GEDCOM Version'].append(sanitize_string(gedc_version))
        
        # Individuals
        for element in gedcom_parser.get_element_list():
            if isinstance(element, IndividualElement):
                id = element.get_pointer()
                name = element.get_name()[0] + " " + element.get_name()[1] if element.get_name() else "Unknown"
                gender = element.get_gender()
                birth_date = birth_place = death_date = death_place = occupation = ""
                education = religion = nationality = description = ssn = title = cause = residence = ""
                notes = []
                change_date = change_time = ""
                
                for child in element.get_child_elements():
                    if child.get_tag() == "BIRT":
                        for sub in child.get_child_elements():
                            if sub.get_tag() == "DATE": birth_date = sub.get_value()
                            if sub.get_tag() == "PLAC": birth_place = sub.get_value()
                    elif child.get_tag() == "DEAT":
                        for sub in child.get_child_elements():
                            if sub.get_tag() == "DATE": death_date = sub.get_value()
                            if sub.get_tag() == "PLAC": death_place = sub.get_value()
                            if sub.get_tag() == "CAUS": cause = sub.get_value()
                    elif child.get_tag() == "OCCU": occupation = child.get_value()
                    elif child.get_tag() == "EDUC": education = child.get_value()
                    elif child.get_tag() == "RELI": religion = child.get_value()
                    elif child.get_tag() == "NATI": nationality = child.get_value()
                    elif child.get_tag() == "DSCR": description = child.get_value()
                    elif child.get_tag() == "SSN": ssn = child.get_value()
                    elif child.get_tag() == "TITL": title = child.get_value()
                    elif child.get_tag() == "RESI": residence = child.get_value()
                    elif child.get_tag() == "NOTE": notes.append(child.get_value())
                    elif child.get_tag() == "CHAN":
                        for sub in child.get_child_elements():
                            if sub.get_tag() == "DATE": change_date = sub.get_value()
                            if sub.get_tag() == "TIME": change_time = sub.get_value()
                    elif child.get_tag() in ["BAPM", "CHR", "BURI", "CREM", "ADOP", "GRAD", "RETI",
                                            "NATU", "EMIG", "IMMI", "CENS", "WILL", "PROB",
                                            "CONF", "FCOM", "BARM", "BASM", "BAPL", "ENDL",
                                            "SLGC", "SLGS"]:
                        event_type = child.get_tag()
                        date = place = cause = event_notes = ""
                        source_ids = []
                        for sub in child.get_child_elements():
                            if sub.get_tag() == "DATE": date = sub.get_value()
                            if sub.get_tag() == "PLAC": place = sub.get_value()
                            if sub.get_tag() == "CAUS": cause = sub.get_value()
                            if sub.get_tag() == "NOTE": event_notes = sub.get_value()
                            if sub.get_tag() == "SOUR": source_ids.append(sub.get_value())
                        events_data['Record ID'].append(id)
                        events_data['Record Type'].append("INDI")
                        events_data['Event Type'].append(event_type)
                        events_data['Date'].append(sanitize_string(date))
                        events_data['Place'].append(sanitize_string(place))
                        events_data['Cause'].append(sanitize_string(cause))
                        events_data['Notes'].append(sanitize_string(event_notes))
                        events_data['Source IDs'].append(sanitize_string(", ".join(source_ids)))
                
                id_to_name[id] = name
                individuals_data['ID'].append(sanitize_string(id))
                individuals_data['Name'].append(sanitize_string(name))
                individuals_data['Gender'].append(sanitize_string(gender))
                individuals_data['Birth Date'].append(sanitize_string(birth_date))
                individuals_data['Birth Place'].append(sanitize_string(birth_place))
                individuals_data['Death Date'].append(sanitize_string(death_date))
                individuals_data['Death Place'].append(sanitize_string(death_place))
                individuals_data['Occupation'].append(sanitize_string(occupation))
                individuals_data['Education'].append(sanitize_string(education))
                individuals_data['Religion'].append(sanitize_string(religion))
                individuals_data['Nationality'].append(sanitize_string(nationality))
                individuals_data['Physical Description'].append(sanitize_string(description))
                individuals_data['SSN'].append(sanitize_string(ssn))
                individuals_data['Title'].append(sanitize_string(title))
                individuals_data['Cause of Death'].append(sanitize_string(cause))
                individuals_data['Residence'].append(sanitize_string(residence))
                individuals_data['Father ID'].append("")
                individuals_data['Mother ID'].append("")
                individuals_data['Father Name'].append("")
                individuals_data['Mother Name'].append("")
                individuals_data['Spouse ID'].append("")
                individuals_data['Spouse Name'].append("")
                individuals_data['Children IDs'].append("")
                individuals_data['Children Names'].append("")
                individuals_data['Notes'].append(sanitize_string(", ".join(notes)))
                individuals_data['Change Date'].append(sanitize_string(change_date))
                individuals_data['Change Time'].append(sanitize_string(change_time))
        
        # Families
        for element in gedcom_parser.get_element_list():
            if isinstance(element, FamilyElement):
                fam_id = element.get_pointer()
                husband_id = wife_id = ""
                children_ids = []
                marr_date = marr_place = div_date = div_place = ""
                eng_date = eng_place = marc_date = marc_place = mars_date = mars_place = ""
                notes = []
                change_date = change_time = ""
                
                for child in element.get_child_elements():
                    if child.get_tag() == "HUSB": husband_id = child.get_value()
                    elif child.get_tag() == "WIFE": wife_id = child.get_value()
                    elif child.get_tag() == "CHIL": children_ids.append(child.get_value())
                    elif child.get_tag() == "MARR":
                        for sub in child.get_child_elements():
                            if sub.get_tag() == "DATE": marr_date = sub.get_value()
                            if sub.get_tag() == "PLAC": marr_place = sub.get_value()
                    elif child.get_tag() == "DIV":
                        for sub in child.get_child_elements():
                            if sub.get_tag() == "DATE": div_date = sub.get_value()
                            if sub.get_tag() == "PLAC": div_place = sub.get_value()
                    elif child.get_tag() == "ENGA":
                        for sub in child.get_child_elements():
                            if sub.get_tag() == "DATE": eng_date = sub.get_value()
                            if sub.get_tag() == "PLAC": eng_place = sub.get_value()
                    elif child.get_tag() == "MARC":
                        for sub in child.get_child_elements():
                            if sub.get_tag() == "DATE": marc_date = sub.get_value()
                            if sub.get_tag() == "PLAC": marc_place = sub.get_value()
                    elif child.get_tag() == "MARS":
                        for sub in child.get_child_elements():
                            if sub.get_tag() == "DATE": mars_date = sub.get_value()
                            if sub.get_tag() == "PLAC": mars_place = sub.get_value()
                    elif child.get_tag() == "NOTE": notes.append(child.get_value())
                    elif child.get_tag() == "CHAN":
                        for sub in child.get_child_elements():
                            if sub.get_tag() == "DATE": change_date = sub.get_value()
                            if sub.get_tag() == "TIME": change_time = sub.get_value()
                
                for child_id in children_ids:
                    child_to_parents[child_id] = (husband_id, wife_id)
                if husband_id:
                    id_to_spouses.setdefault(husband_id, []).append(wife_id)
                    id_to_children.setdefault(husband_id, []).extend(children_ids)
                if wife_id:
                    id_to_spouses.setdefault(wife_id, []).append(husband_id)
                    id_to_children.setdefault(wife_id, []).extend(children_ids)
                
                families_data['Family ID'].append(sanitize_string(fam_id))
                families_data['Husband ID'].append(sanitize_string(husband_id))
                families_data['Husband Name'].append(sanitize_string(id_to_name.get(husband_id, "")))
                families_data['Wife ID'].append(sanitize_string(wife_id))
                families_data['Wife Name'].append(sanitize_string(id_to_name.get(wife_id, "")))
                families_data['Marriage Date'].append(sanitize_string(marr_date))
                families_data['Marriage Place'].append(sanitize_string(marr_place))
                families_data['Divorce Date'].append(sanitize_string(div_date))
                families_data['Divorce Place'].append(sanitize_string(div_place))
                families_data['Engagement Date'].append(sanitize_string(eng_date))
                families_data['Engagement Place'].append(sanitize_string(eng_place))
                families_data['Marriage Contract Date'].append(sanitize_string(marc_date))
                families_data['Marriage Contract Place'].append(sanitize_string(marc_place))
                families_data['Marriage Settlement Date'].append(sanitize_string(mars_date))
                families_data['Marriage Settlement Place'].append(sanitize_string(mars_place))
                families_data['Children IDs'].append(sanitize_string(", ".join(children_ids)))
                families_data['Children Names'].append(sanitize_string(", ".join(id_to_name.get(cid, "") for cid in children_ids)))
                families_data['Notes'].append(sanitize_string(", ".join(notes)))
                families_data['Change Date'].append(sanitize_string(change_date))
                families_data['Change Time'].append(sanitize_string(change_time))
        
        # Sources
        for element in gedcom_parser.get_element_list():
            if element.get_tag() == "SOUR":
                src_id = element.get_pointer()
                title = author = pub = page = repo = data = ""
                notes = []
                for child in element.get_child_elements():
                    if child.get_tag() == "TITL": title = child.get_value()
                    elif child.get_tag() == "AUTH": author = child.get_value()
                    elif child.get_tag() == "PUBL": pub = child.get_value()
                    elif child.get_tag() == "PAGE": page = child.get_value()
                    elif child.get_tag() == "REPO": repo = child.get_value()
                    elif child.get_tag() == "DATA": data = child.get_value()
                    elif child.get_tag() == "NOTE": notes.append(child.get_value())
                sources_data['Source ID'].append(sanitize_string(src_id))
                sources_data['Title'].append(sanitize_string(title))
                sources_data['Author'].append(sanitize_string(author))
                sources_data['Publication'].append(sanitize_string(pub))
                sources_data['Page'].append(sanitize_string(page))
                sources_data['Repository'].append(sanitize_string(repo))
                sources_data['Data'].append(sanitize_string(data))
                sources_data['Notes'].append(sanitize_string(", ".join(notes)))
        
        # Notes
        for element in gedcom_parser.get_element_list():
            if element.get_tag() == "NOTE":
                note_id = element.get_pointer()
                notes_data['Note ID'].append(sanitize_string(note_id))
                notes_data['Text'].append(sanitize_string(element.get_value()))
                notes_data['Referenced By'].append("")
        
        # Multimedia
        for element in gedcom_parser.get_element_list():
            if element.get_tag() == "OBJE":
                obj_id = element.get_pointer()
                file = form = title = ""
                notes = []
                for child in element.get_child_elements():
                    if child.get_tag() == "FILE": file = child.get_value()
                    if child.get_tag() == "FORM": form = child.get_value()
                    if child.get_tag() == "TITL": title = child.get_value()
                    if child.get_tag() == "NOTE": notes.append(child.get_value())
                multimedia_data['Object ID'].append(sanitize_string(obj_id))
                multimedia_data['File'].append(sanitize_string(file))
                multimedia_data['Format'].append(sanitize_string(form))
                multimedia_data['Title'].append(sanitize_string(title))
                multimedia_data['Notes'].append(sanitize_string(", ".join(notes)))
        
        # Associations
        for element in gedcom_parser.get_element_list():
            if isinstance(element, IndividualElement):
                id = element.get_pointer()
                for child in element.get_child_elements():
                    if child.get_tag() == "ASSO":
                        assoc_id = child.get_value()
                        rela = ""
                        notes = []
                        for sub in child.get_child_elements():
                            if sub.get_tag() == "RELA": rela = sub.get_value()
                            if sub.get_tag() == "NOTE": notes.append(sub.get_value())
                        associations_data['Individual ID'].append(sanitize_string(id))
                        associations_data['Associated ID'].append(sanitize_string(assoc_id))
                        associations_data['Relationship'].append(sanitize_string(rela))
                        associations_data['Notes'].append(sanitize_string(", ".join(notes)))
        
        # Submitter
        for element in gedcom_parser.get_element_list():
            if element.get_tag() == "SUBM":
                subm_id = element.get_pointer()
                name = addr = phone = email = ""
                for child in element.get_child_elements():
                    if child.get_tag() == "NAME": name = child.get_value()
                    if child.get_tag() == "ADDR": addr = child.get_value()
                    if child.get_tag() == "PHON": phone = child.get_value()
                    if child.get_tag() == "EMAIL": email = child.get_value()
                submitter_data['Submitter ID'].append(sanitize_string(subm_id))
                submitter_data['Name'].append(sanitize_string(name))
                submitter_data['Address'].append(sanitize_string(addr))
                submitter_data['Phone'].append(sanitize_string(phone))
                submitter_data['Email'].append(sanitize_string(email))
        
        # Update individuals with relationships
        for i, indiv_id in enumerate(individuals_data['ID']):
            if indiv_id in child_to_parents:
                father_id, mother_id = child_to_parents[indiv_id]
                individuals_data['Father ID'][i] = sanitize_string(father_id)
                individuals_data['Mother ID'][i] = sanitize_string(mother_id)
                individuals_data['Father Name'][i] = sanitize_string(id_to_name.get(father_id, ""))
                individuals_data['Mother Name'][i] = sanitize_string(id_to_name.get(mother_id, ""))
            if indiv_id in id_to_spouses:
                spouse_ids = id_to_spouses[indiv_id]
                individuals_data['Spouse ID'][i] = sanitize_string(", ".join(spouse_ids))
                individuals_data['Spouse Name'][i] = sanitize_string(", ".join(id_to_name.get(sid, "") for sid in spouse_ids))
            if indiv_id in id_to_children:
                children_ids = id_to_children[indiv_id]
                individuals_data['Children IDs'][i] = sanitize_string(", ".join(children_ids))
                individuals_data['Children Names'][i] = sanitize_string(", ".join(id_to_name.get(cid, "") for cid in children_ids))
        
        # Create DataFrames
        individuals_df = pd.DataFrame(individuals_data)
        families_df = pd.DataFrame(families_data)
        events_df = pd.DataFrame(events_data)
        sources_df = pd.DataFrame(sources_data)
        notes_df = pd.DataFrame(notes_data)
        multimedia_df = pd.DataFrame(multimedia_data)
        associations_df = pd.DataFrame(associations_data)
        submitter_df = pd.DataFrame(submitter_data)
        header_df = pd.DataFrame(header_data)
        
        # Write to Excel
        with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
            individuals_df.to_excel(writer, sheet_name='Individuals', index=False)
            families_df.to_excel(writer, sheet_name='Families', index=False)
            events_df.to_excel(writer, sheet_name='Events', index=False)
            sources_df.to_excel(writer, sheet_name='Sources', index=False)
            notes_df.to_excel(writer, sheet_name='Notes', index=False)
            multimedia_df.to_excel(writer, sheet_name='Multimedia', index=False)
            associations_df.to_excel(writer, sheet_name='Associations', index=False)
            submitter_df.to_excel(writer, sheet_name='Submitter', index=False)
            header_df.to_excel(writer, sheet_name='Header', index=False)
        
        print(f"Conversion complete. Excel file saved as: {excel_file_path}")
    
    finally:
        if os.path.exists(repaired_gedcom_path):
            os.remove(repaired_gedcom_path)
            print(f"Temporary repaired GEDCOM file deleted: {repaired_gedcom_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert a GEDCOM file to an Excel file, repairing line numbering and sanitizing data.")
    parser.add_argument("input_file", help="Path to the input GEDCOM file (e.g., family.ged)")
    parser.add_argument("output_file", help="Path to the output Excel file (e.g., output.xlsx)")
    args = parser.parse_args()
    
    try:
        gedcom_to_excel(args.input_file, args.output_file)
    except FileNotFoundError:
        print(f"Error: GEDCOM file '{args.input_file}' not found. Please check the file path.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
